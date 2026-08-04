#!/usr/bin/env node
// Daily snapshot writer. Captures what the dashboard said about each
// company today — composite score, rating band, close price — so the
// History tab can later show "we said STRONG BUY at ₹444; today it's
// ₹555 — predicted well." One file per calendar date.
//
// Reads: public/data/{screener-companies.json,technicals.json,macro.json}
// Writes:
//   public/data/snapshots/YYYY-MM-DD.json   (per-day record)
//   public/data/snapshots/index.json        (manifest of available dates)
//
// Idempotent — if today's file exists, it's overwritten only when the
// content actually differs; the manifest is rebuilt from disk each run
// so re-ordering or hand-deletes are picked up automatically.

import { readFileSync, writeFileSync, mkdirSync, readdirSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { scoreCompositeBatch, PILLAR_WEIGHTS } from "../public/js/composite-scoring.js";
import * as techScoring from "../public/js/tech-scoring.js";
import * as fundScoring from "../public/js/scoring.js";
import { enrichCompanies } from "./lib/enrich-companies.mjs";
import { mergePinned } from "./lib/pinned.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR       = resolve(__dirname, "../public/data");
const FUND_PATH      = resolve(__dirname, "../public/data/screener-companies.json");
const TECH_PATH      = resolve(__dirname, "../public/data/technicals.json");
const MACRO_PATH     = resolve(__dirname, "../public/data/macro.json");
const SNAPSHOTS_DIR  = resolve(__dirname, "../public/data/snapshots");
const INDEX_PATH     = resolve(SNAPSHOTS_DIR, "index.json");

const FRAMEWORK_VERSION = "v1";

run().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  process.exit(1);
});

async function run() {
  const dateArg = process.argv[2];                    // optional override (used by backfill)
  const date = dateArg || new Date().toISOString().slice(0, 10);

  console.log(`Building snapshot for ${date}...`);
  const fund  = readJson(FUND_PATH);
  const tech  = readJson(TECH_PATH);
  // macro.json is no longer produced -- the Macro pillar was removed and its
  // scraper deleted. It must therefore be OPTIONAL: readJson() throws ENOENT,
  // which would kill the snapshot writer outright and leave the dashboard with
  // no history at all. macro-scoring.js is null-safe (returns 0/17, all N/A)
  // and the pillar is zero-weighted, so null costs nothing.
  const macro = readJsonSafe(MACRO_PATH);

  const liveCompanies = Array.isArray(fund) ? fund : (fund.companies || []);
  // Cohort members that have left the market-cap band still need a snapshot
  // row every day, otherwise the tracker loses the position mid-hold and the
  // month's return is measured over a shorter window than it was held.
  const pinnedMerge = mergePinned(liveCompanies, DATA_DIR);
  const fundCompanies = pinnedMerge.companies;
  if (pinnedMerge.added) {
    console.log(`  pinned (outside band, still held): ${pinnedMerge.addedTickers.join(", ")}`);
  }
  const techCompanies = tech.companies || tech || [];

  // Fuse the same auxiliary data the live SPIP Basket tab merges before
  // scoring (auditor / governance / insider / revenue-mix / macro overlays
  // / ATR) so the snapshot's composite, rating and hard-fails match what
  // the dashboard renders for the same day.
  const enrich = enrichCompanies(fundCompanies, techCompanies, macro, DATA_DIR);
  console.log(`  enrichment: auditor ${enrich.auditorLoaded ? enrich.counts.auditor : "off"} · governance ${enrich.governanceLoaded ? enrich.counts.governance : "off"} · insider ${enrich.insiderLoaded ? enrich.counts.insider : "off"} · revenue-mix ${enrich.counts.revenueMix}`);

  const scored = scoreCompositeBatch(fundCompanies, techCompanies, macro);

  // Tech rows are keyed by ticker — same path scoreCompositeBatch takes
  // when looking up a fund company's tech match, so the close prices we
  // attach here are guaranteed to match the score.
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

    // Pillar breakdown — lean per-pillar shape so forensics can decompose
    // composite deltas day-over-day. pct = 0..100, weighted = contribution
    // to the 100-point composite (sum of all five = composite).
    const p = s.pillars || {};
    // Keyed off PILLAR_WEIGHTS so a pillar that is removed from the model
    // stops appearing in the snapshot as a permanent null column.
    const pillars = {};
    for (const k of Object.keys(PILLAR_WEIGHTS)) pillars[k] = leanPillar(p[k]);

    // Per-indicator raw values — lets the Custom Lab's tweakable
    // parameter filters (RSI ≥ 65, within 5% of high, Beta 0.7–1.3…)
    // backtest historically. The indicator backtest "accrues forward" as
    // snapshots gain this field. null when there's no usable tech row.
    let techVals = null;
    try { techVals = techScoring.techVals(tc); } catch { techVals = null; }
    // Fundamental ratios (Custom Lab "Fundamentals" params) + industry
    // (finer than sector — powers the sector/industry rebalance-timing view).
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

  const snapshot = {
    date,
    framework_version: FRAMEWORK_VERSION,
    weights: PILLAR_WEIGHTS,
    generated_at: new Date().toISOString(),
    stock_count: stocks.length,
    coverage: { with_close: withClose, without_close: withoutClose },
    stocks,
  };

  mkdirSync(SNAPSHOTS_DIR, { recursive: true });
  const outPath = resolve(SNAPSHOTS_DIR, `${date}.json`);
  writeIfChanged(outPath, JSON.stringify(snapshot, null, 2));

  rebuildIndex();

  console.log(`Wrote ${outPath}`);
  console.log(`  ${stocks.length} stocks · ${withClose} with close · ${withoutClose} without`);
}

function readJson(path) {
  return JSON.parse(readFileSync(path, "utf8"));
}

// Same, but returns null instead of throwing — for inputs whose absence is a
// valid state rather than a broken run.
function readJsonSafe(path) {
  try { return JSON.parse(readFileSync(path, "utf8")); } catch { return null; }
}

function writeIfChanged(path, content) {
  if (existsSync(path)) {
    const old = readFileSync(path, "utf8");
    if (old === content) {
      console.log(`  (unchanged, skipping write)`);
      return;
    }
  }
  writeFileSync(path, content);
}

// Browser code can't list a directory, so we maintain a manifest of
// dates the History tab can fetch. Order: oldest → newest.
function rebuildIndex() {
  const files = readdirSync(SNAPSHOTS_DIR)
    .filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .map((f) => f.replace(".json", ""))
    .sort();
  // Preserve live_from: the reconstructed/real boundary written by
  // backfill-from-history.mjs. Rewriting the manifest without it wiped the
  // marker every daily run, so the dashboard lost the "history before X is
  // rebuilt" note. Derive it from the files themselves as the fallback.
  let liveFrom = null;
  try { liveFrom = JSON.parse(readFileSync(INDEX_PATH, "utf8")).live_from ?? null; } catch {}
  if (!liveFrom) {
    for (const d of files) {
      try { if (!JSON.parse(readFileSync(resolve(SNAPSHOTS_DIR, `${d}.json`), "utf8")).reconstructed) { liveFrom = d; break; } } catch {}
    }
  }
  const idx = {
    generated_at: new Date().toISOString(),
    framework_version: FRAMEWORK_VERSION,
    count: files.length,
    live_from: liveFrom,
    dates: files,
  };
  writeFileSync(INDEX_PATH, JSON.stringify(idx, null, 2));
  console.log(`Manifest: ${files.length} snapshot dates (${files[0] || "—"} → ${files[files.length - 1] || "—"})`);
}

function slugify(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

// Drop {raw, max} from each pillar — keep just pct (0..100) and weighted
// (contribution to the composite). Round both to 2 decimals.
function leanPillar(p) {
  if (!p) return null;
  return {
    pct:      p.pct      == null ? null : Number(p.pct.toFixed(2)),
    weighted: p.weighted == null ? null : Number(p.weighted.toFixed(3)),
  };
}
