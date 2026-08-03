#!/usr/bin/env node
// Snapshot backfill for a repo with no git history to walk.
//
// klpdash builds its history with backfill-snapshots.mjs, which replays
// the git log of technicals.json -- it had been committing that file for
// months, so the past was recoverable. This repo is days old: three
// commits of technicals.json, both of its snapshots landing on a weekend.
// Every tab that needs a trail therefore renders empty, which reads as
// broken rather than new.
//
// What this repo does have is history-technical.json: 504 trading days of
// real closes and real technical scores across 620 names. That is enough
// to reconstruct the technicals pillar exactly as it stood on each of
// those days. The fundamentals pillar cannot be recovered -- nobody was
// storing it -- so today's value is held constant across the window.
//
// That makes these snapshots RECONSTRUCTED, and they say so: every file
// carries reconstructed: true and framework_version "v1-reconstructed",
// and the dashboard draws the boundary where real forward tracking
// begins. The distinction matters. Holding fundamentals constant is
// look-ahead on 30% of the composite: a company whose books deteriorated
// over the window looks as good in March as it does today. Prices and
// technical scores carry no such flaw -- they are what they were.
//
// Usage:
//   node screener-test/backfill-from-history.mjs            # 70 trading days
//   node screener-test/backfill-from-history.mjs --days 120
//   node screener-test/backfill-from-history.mjs --force    # overwrite real ones too

import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { PILLAR_WEIGHTS } from "../public/js/composite-scoring.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA = resolve(__dirname, "../public/data");
const SNAP_DIR = resolve(DATA, "snapshots");

const arg = (k, d) => {
  const i = process.argv.indexOf(k);
  return i > 0 && process.argv[i + 1] ? process.argv[i + 1] : d;
};
const DAYS = Number(arg("--days", 70));
const FORCE = process.argv.includes("--force");

// ---- indicator maths, trailing only (never peeks forward) ----
const sma = (a, i, n) => { if (i + 1 < n) return null; let s = 0; for (let k = i - n + 1; k <= i; k++) { if (a[k] == null) return null; s += a[k]; } return s / n; };
const ema = (a, i, n) => { const st = Math.max(0, i - n * 3); let e = null, k = 2 / (n + 1); for (let j = st; j <= i; j++) { const v = a[j]; if (v == null) continue; e = e == null ? v : v * k + e * (1 - k); } return e; };
function rsi(a, i, n = 14) {
  if (i < n) return null;
  let g = 0, l = 0;
  for (let k = i - n + 1; k <= i; k++) { const d = a[k] - a[k - 1]; if (!Number.isFinite(d)) return null; d >= 0 ? g += d : l -= d; }
  if (g + l === 0) return 50;
  return 100 - 100 / (1 + (g / n) / ((l / n) || 1e-9));
}
// Close-to-close stand-in for ATR%; the history has no high/low, so this
// reads slightly tighter than a true ATR and is scaled to compensate.
function atrPct(a, i, n = 14) {
  if (i < n) return null;
  let s = 0, c = 0;
  for (let k = i - n + 1; k <= i; k++) if (a[k - 1] > 0) { s += Math.abs(a[k] / a[k - 1] - 1); c++; }
  return c ? +(s / c * 100 * 1.35).toFixed(2) : null;
}
function offHigh(a, i, n = 252) {
  let hi = -Infinity;
  for (let k = Math.max(0, i - n + 1); k <= i; k++) if (a[k] != null && a[k] > hi) hi = a[k];
  return hi > 0 && a[i] != null ? +(((hi - a[i]) / hi) * 100).toFixed(2) : null;
}

function run() {
  const hist = JSON.parse(readFileSync(resolve(DATA, "history-technical.json"), "utf8"));
  const bench = JSON.parse(readFileSync(resolve(DATA, "benchmark-history.json"), "utf8"));
  const cal = bench.indices?.["NIFTYSMLCAP250.NS"]?.closes || {};

  // The most recent real snapshot supplies identity (name, slug, sector,
  // industry) and the frozen fundamentals pillar.
  // Split what is on disk into real and reconstructed by reading each file's
  // own flag. Filenames are identical either way, so a second run would
  // otherwise mistake its own output for real data -- refusing to refresh it,
  // and sourcing fundamentals from a reconstructed file instead of a scraped one.
  const onDisk = readdirSync(SNAP_DIR).filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f)).sort();
  const realFiles = [];
  for (const f of onDisk) {
    try { if (!JSON.parse(readFileSync(resolve(SNAP_DIR, f), "utf8")).reconstructed) realFiles.push(f); } catch {}
  }
  const existing = realFiles;
  if (!existing.length) throw new Error("Need at least one real snapshot to source fundamentals and company identity.");
  const latest = JSON.parse(readFileSync(resolve(SNAP_DIR, existing[existing.length - 1]), "utf8"));
  const meta = new Map();
  for (const s of latest.stocks) {
    if (!s.ticker) continue;
    meta.set(s.ticker, {
      name: s.name, slug: s.slug, sector: s.sector, industry: s.industry,
      fundPct: s.pillars?.fundamentals?.pct ?? null,
      fundVals: s.fundVals || null,
      hardFailed: !!s.hardFailed,
    });
  }

  const wF = PILLAR_WEIGHTS.fundamentals, wT = PILLAR_WEIGHTS.technicals;
  const tickers = Object.keys(hist.tickers);
  const realDates = new Set(existing.map((f) => f.slice(0, 10)));

  // Trading days only. Snapshots on a day the market was shut carry the
  // previous close and would draw a flat step into every chart.
  const idxs = [];
  for (let i = hist.dates.length - 1; i >= 0 && idxs.length < DAYS; i--) {
    if (cal[hist.dates[i]] != null) idxs.push(i);
  }
  idxs.reverse();
  console.log(`Reconstructing ${idxs.length} trading days (${hist.dates[idxs[0]]} → ${hist.dates[idxs[idxs.length - 1]]})…`);

  let written = 0, skipped = 0;
  for (const i of idxs) {
    const date = hist.dates[i];
    if (realDates.has(date) && !FORCE) { skipped++; continue; }

    const stocks = [];
    for (const t of tickers) {
      const rec = hist.tickers[t];
      const close = rec.close[i], techPct = rec.score[i];
      if (close == null || techPct == null) continue;
      const m = meta.get(t);
      const fundPct = m?.fundPct ?? null;
      // Composite on the live framework's weights. With no fundamentals
      // for a name, the technicals pillar is rebased to 100% rather than
      // scoring it as if fundamentals were zero -- that would bury every
      // uncovered company at the bottom of the ranking for a data gap.
      const composite = fundPct == null
        ? +techPct.toFixed(1)
        : +((fundPct * wF + techPct * wT) / (wF + wT)).toFixed(1);
      const cl = rec.close;
      const m50 = sma(cl, i, 50), m200 = sma(cl, i, 200);
      const e12 = ema(cl, i, 12), e26 = ema(cl, i, 26);
      stocks.push({
        ticker: t,
        name: m?.name || t,
        slug: m?.slug || null,
        sector: m?.sector || rec.sector || "—",
        industry: m?.industry || null,
        composite,
        rating: composite >= 75 ? "STRONG BUY" : composite >= 60 ? "BUY" : composite >= 40 ? "HOLD" : "AVOID",
        hardFailed: !!m?.hardFailed,
        dataComplete: true,
        close,
        pillars: {
          fundamentals: fundPct == null ? undefined : { pct: fundPct, weighted: +(fundPct * wF / 100).toFixed(2) },
          technicals: { pct: +techPct.toFixed(1), weighted: +(techPct * wT / 100).toFixed(2) },
        },
        techVals: {
          rsi: +(rsi(cl, i) ?? 50).toFixed(1),
          atr: atrPct(cl, i) ?? 3.5,
          d52: offHigh(cl, i),
          e50: m50 != null ? close > m50 : null,
          d200: m200 != null ? close > m200 : null,
          macd: e12 != null && e26 != null ? e12 > e26 : null,
        },
        fundVals: m?.fundVals || undefined,
      });
    }
    if (!stocks.length) continue;
    writeFileSync(resolve(SNAP_DIR, `${date}.json`), JSON.stringify({
      date,
      framework_version: "v1-reconstructed",
      reconstructed: true,
      reconstruction_note: "Technicals and closes are the real values for this date. The fundamentals pillar is held at its latest value because no historical record exists — treat the fundamentals 30% as look-ahead.",
      generated_at: new Date().toISOString(),
      weights: PILLAR_WEIGHTS,
      stock_count: stocks.length,
      stocks,
    }));
    written++;
  }

  // Manifest, oldest → newest.
  const files = readdirSync(SNAP_DIR).filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f)).map((f) => f.slice(0, 10)).sort();
  // Earliest snapshot that is NOT reconstructed. Read the flag out of each
  // file rather than trusting the listing: on a second run this script's own
  // output is indistinguishable from real data by filename alone, and the
  // boundary would silently slide back to the start of the rebuilt window.
  let liveFrom = null;
  for (const d of files) {
    try {
      const j = JSON.parse(readFileSync(resolve(SNAP_DIR, `${d}.json`), "utf8"));
      if (!j.reconstructed) { liveFrom = d; break; }
    } catch {}
  }
  const idxPath = resolve(SNAP_DIR, "index.json");
  const prev = existsSync(idxPath) ? JSON.parse(readFileSync(idxPath, "utf8")) : {};
  writeFileSync(idxPath, JSON.stringify({
    ...prev,
    generated_at: new Date().toISOString(),
    count: files.length,
    // Where reconstruction stops and real forward tracking starts. The UI
    // marks this so nobody mistakes a rebuilt past for a live record.
    live_from: liveFrom,
    dates: files,
  }, null, 2) + "\n");

  console.log(`Wrote ${written} reconstructed snapshots, skipped ${skipped} real ones. Manifest now lists ${files.length} dates.`);
}

// Called last: the indicator helpers above are const-bound, so invoking
// run() before this point hits the temporal dead zone.
try { run(); } catch (e) { console.error("Fatal:", e.stack || e.message); process.exit(1); }
