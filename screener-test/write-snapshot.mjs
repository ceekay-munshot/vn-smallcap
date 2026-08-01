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
  const macro = readJson(MACRO_PATH);

  const fundCompanies = Array.isArray(fund) ? fund : (fund.companies || []);
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
    const pillars = {
      fundamentals: leanPillar(p.fundamentals),
      technicals:   leanPillar(p.technicals),
      macro:        leanPillar(p.macro),
      sentiment:    leanPillar(p.sentiment),
      liquidity:    leanPillar(p.liquidity),
    };

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
  const idx = {
    generated_at: new Date().toISOString(),
    framework_version: FRAMEWORK_VERSION,
    count: files.length,
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
