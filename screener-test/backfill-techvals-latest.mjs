#!/usr/bin/env node
// One-off: add the Custom Lab per-stock signal records (techVals +
// fundVals) and `industry` to the LATEST snapshot only, by joining it
// with today's technicals.json + screener-companies.json. The latest
// snapshot and those files are the same trading day, so this is accurate.
// Earlier snapshots are intentionally left untouched — we have no
// historical granular data for them, so the Custom Lab backtest falls
// back to composite/pillar-only on those days and "accrues forward" as
// the daily writer (write-snapshot.mjs) stamps these on each new snapshot.
//
// Purely additive: only techVals / fundVals / industry are added per
// stock; every other field is preserved by re-serializing the same object.
import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import * as techScoring from "../public/js/tech-scoring.js";
import * as fundScoring from "../public/js/scoring.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAP_DIR = resolve(__dirname, "../public/data/snapshots");
const TECH_PATH = resolve(__dirname, "../public/data/technicals.json");
const FUND_PATH = resolve(__dirname, "../public/data/screener-companies.json");

const idx = JSON.parse(readFileSync(resolve(SNAP_DIR, "index.json"), "utf8"));
const latest = idx.dates[idx.dates.length - 1];
const snapPath = resolve(SNAP_DIR, `${latest}.json`);
const snap = JSON.parse(readFileSync(snapPath, "utf8"));
const tech = JSON.parse(readFileSync(TECH_PATH, "utf8"));
const fundRaw = JSON.parse(readFileSync(FUND_PATH, "utf8"));
const fundCompanies = Array.isArray(fundRaw) ? fundRaw : (fundRaw.companies || []);

const techByTicker = new Map((tech.companies || []).map((c) => [String(c.ticker).toUpperCase(), c]));
const extractTicker = (url) => { const m = String(url || "").match(/\/company\/([^/]+)/); return m ? m[1].toUpperCase() : null; };
const fundByTicker = new Map();
for (const fc of fundCompanies) { const t = extractTicker(fc["Screener URL"]); if (t) fundByTicker.set(t, fc); }

let withTech = 0, withFund = 0;
for (const s of snap.stocks) {
  const key = s.ticker ? String(s.ticker).toUpperCase() : null;
  const tc = key ? techByTicker.get(key) : null;
  const fc = key ? fundByTicker.get(key) : null;
  delete s.techPass;   // superseded by techVals
  try { s.techVals = techScoring.techVals(tc); } catch { s.techVals = null; }
  try { s.fundVals = fundScoring.fundVals(fc); } catch { s.fundVals = null; }
  if (fc && (fc.Industry || fc["Broad Industry"])) s.industry = fc.Industry || fc["Broad Industry"];
  if (s.techVals) withTech++;
  if (s.fundVals) withFund++;
}

writeFileSync(snapPath, JSON.stringify(snap, null, 2));
console.log(`${latest}: techVals ${withTech}/${snap.stocks.length}, fundVals ${withFund}/${snap.stocks.length}, industry added.`);
