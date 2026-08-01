#!/usr/bin/env node
// Backfill daily snapshots from git history. We don't have a snapshot
// store yet, but the daily-refresh workflow has been committing
// public/data/technicals.json (the canonical daily close source) for
// months. Walk that file's git log, replay the composite scorer
// against each historical commit, and write one snapshot per date.
//
// One-shot: usually run once locally or in a workflow_dispatch job, not
// on every refresh. Subsequent days are handled by write-snapshot.mjs
// inside daily-refresh.yml.
//
// Caveats:
//   - Re-scores with TODAY'S framework rules / weights, stamped via
//     framework_version on each snapshot. If thresholds change later,
//     old snapshots are still readable but won't match the new framework
//     1:1; bump framework_version and rerun this when that happens.
//   - One snapshot per calendar date — if multiple tech commits land
//     on the same day, the LAST one for that date wins.
//   - Skips dates whose snapshot file already exists, so re-runs are
//     cheap. Pass --force to overwrite.
//
// Usage:
//   node screener-test/backfill-snapshots.mjs            # last 60 days
//   node screener-test/backfill-snapshots.mjs --days 30
//   node screener-test/backfill-snapshots.mjs --force    # overwrite

import { execSync } from "node:child_process";
import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { scoreCompositeBatch, PILLAR_WEIGHTS } from "../public/js/composite-scoring.js";
import * as techScoring from "../public/js/tech-scoring.js";
import * as fundScoring from "../public/js/scoring.js";
import { enrichCompanies } from "./lib/enrich-companies.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT     = resolve(__dirname, "..");
const DATA_DIR      = resolve(REPO_ROOT, "public/data");
const SNAPSHOTS_DIR = resolve(REPO_ROOT, "public/data/snapshots");
const INDEX_PATH    = resolve(SNAPSHOTS_DIR, "index.json");
const FRAMEWORK_VERSION = "v1";

const args = process.argv.slice(2);
const FORCE = args.includes("--force");
const daysIdx = args.indexOf("--days");
const DAYS = daysIdx >= 0 ? Math.max(1, parseInt(args[daysIdx + 1], 10) || 60) : 60;

run().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  process.exit(1);
});

async function run() {
  mkdirSync(SNAPSHOTS_DIR, { recursive: true });

  console.log(`Walking git log of public/data/technicals.json (last ${DAYS} days)...`);
  const commits = listTechCommits(DAYS);
  console.log(`Found ${commits.length} candidate commits.`);

  // Group by calendar date — last commit per date wins.
  const byDate = new Map();
  for (const c of commits) byDate.set(c.date, c);
  const dates = [...byDate.keys()].sort();           // oldest → newest
  console.log(`After de-dup: ${dates.length} distinct dates (${dates[0] || "—"} → ${dates[dates.length - 1] || "—"})`);

  let written = 0, skipped = 0, failed = 0;

  for (const date of dates) {
    const out = resolve(SNAPSHOTS_DIR, `${date}.json`);
    if (!FORCE && existsSync(out)) { skipped++; continue; }
    const c = byDate.get(date);
    try {
      const snapshot = await buildSnapshotFromCommit(c.sha, date);
      writeFileSync(out, JSON.stringify(snapshot, null, 2));
      written++;
      const sb = snapshot.stocks.filter((s) => s.rating === "STRONG BUY").length;
      const buy = snapshot.stocks.filter((s) => s.rating === "BUY").length;
      console.log(`  ${date}  sha=${c.sha.slice(0,7)}  ${snapshot.stocks.length} stocks · ${sb} SB · ${buy} BUY`);
    } catch (e) {
      failed++;
      console.warn(`  ${date}  sha=${c.sha.slice(0,7)}  FAILED: ${e.message}`);
    }
  }

  rebuildIndex();

  console.log(`\nDone. Written ${written} · skipped (existed) ${skipped} · failed ${failed}`);
}

// `git log` filtered to commits that touch technicals.json. Each entry
// gets a date stamp (commit author date, normalized to UTC YYYY-MM-DD).
function listTechCommits(days) {
  const since = `${days}.days.ago`;
  const raw = execSync(
    `git log --since='${since}' --format='%H %aI' -- public/data/technicals.json`,
    { cwd: REPO_ROOT, encoding: "utf8" },
  ).trim();
  if (!raw) return [];
  return raw.split("\n").map((line) => {
    const [sha, iso] = line.split(" ");
    return { sha, date: iso.slice(0, 10) };
  });
}

// Read a file's contents at a specific commit. Returns null if the file
// didn't exist at that commit (newly added files at later dates).
function gitShow(sha, path) {
  try {
    return execSync(`git show ${sha}:${path}`, { cwd: REPO_ROOT, encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  } catch {
    return null;
  }
}

async function buildSnapshotFromCommit(sha, date) {
  const fundRaw  = gitShow(sha, "public/data/screener-companies.json");
  const techRaw  = gitShow(sha, "public/data/technicals.json");
  const macroRaw = gitShow(sha, "public/data/macro.json");
  if (!fundRaw || !techRaw) throw new Error("missing fundamentals or technicals at commit");

  const fund  = JSON.parse(fundRaw);
  const tech  = JSON.parse(techRaw);
  const macro = macroRaw ? JSON.parse(macroRaw) : {};

  const fundCompanies = Array.isArray(fund) ? fund : (fund.companies || []);
  const techCompanies = tech.companies || tech || [];

  // Enrich with the same auxiliary data the live tab / daily writer use.
  // Historical fund/tech/macro come from the commit; the aux sources
  // (auditor / governance / insider / revenue-mix / ATR) use their current
  // versions — they change slowly, so this is the best available proxy for
  // what the tab would have shown that day.
  enrichCompanies(fundCompanies, techCompanies, macro, DATA_DIR);

  const scored = scoreCompositeBatch(fundCompanies, techCompanies, macro);

  const techByTicker = {};
  for (const t of techCompanies) {
    if (t && t.ticker) techByTicker[String(t.ticker).toUpperCase()] = t;
  }
  const extractTicker = (url) => {
    const m = String(url || "").match(/\/company\/([^/]+)/);
    return m ? m[1].toUpperCase() : null;
  };

  let withClose = 0, withoutClose = 0;
  const stocks = scored.map((s) => {
    const fc = s.company || {};
    const ticker = extractTicker(fc["Screener URL"]);
    const tc = ticker ? techByTicker[ticker] : null;
    const close = typeof tc?.cmp === "number" ? tc.cmp : null;
    if (close != null) withClose++; else withoutClose++;
    const p = s.pillars || {};
    const pillars = {
      fundamentals: leanPillar(p.fundamentals),
      technicals:   leanPillar(p.technicals),
      macro:        leanPillar(p.macro),
      sentiment:    leanPillar(p.sentiment),
      liquidity:    leanPillar(p.liquidity),
    };
    // Same per-indicator / per-ratio fields the daily writer emits, so a
    // backfilled snapshot is interchangeable with a live-written one (the
    // Custom Lab back-test reads techVals / fundVals).
    let techVals = null;
    try { techVals = techScoring.techVals(tc); } catch { techVals = null; }
    let fundVals = null;
    try { fundVals = fundScoring.fundVals(fc); } catch { fundVals = null; }
    return {
      ticker,
      name: fc.Company || null,
      slug: slugify(fc.Company || ticker || ""),
      sector: fc.Sector || fc["Broad Industry"] || null,
      industry: fc.Industry || fc["Broad Industry"] || null,
      composite: s.composite == null ? null : Number(s.composite.toFixed(2)),
      rating: s.rating || null,
      hardFailed: !!s.hardFailed,
      dataComplete: !!s.dataComplete,
      close,
      pillars,
      techVals,
      fundVals,
    };
  });

  return {
    date,
    framework_version: FRAMEWORK_VERSION,
    weights: PILLAR_WEIGHTS,
    generated_at: new Date().toISOString(),
    source: { kind: "backfill", commit: sha },
    stock_count: stocks.length,
    coverage: { with_close: withClose, without_close: withoutClose },
    stocks,
  };
}

function rebuildIndex() {
  const files = readdirSync(SNAPSHOTS_DIR)
    .filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .map((f) => f.replace(".json", ""))
    .sort();
  writeFileSync(INDEX_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    framework_version: FRAMEWORK_VERSION,
    count: files.length,
    dates: files,
  }, null, 2));
  console.log(`Manifest: ${files.length} snapshot dates (${files[0] || "—"} → ${files[files.length - 1] || "—"})`);
}

function slugify(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function leanPillar(p) {
  if (!p) return null;
  return {
    pct:      p.pct      == null ? null : Number(p.pct.toFixed(2)),
    weighted: p.weighted == null ? null : Number(p.weighted.toFixed(3)),
  };
}
