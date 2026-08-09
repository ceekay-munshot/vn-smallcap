#!/usr/bin/env node
// Monthly cohort + entry-price capture for the Strategy basket.
//
// THE RULE (locked with the client):
//   1. SELECTION — the basket is the top-7 by composite on the LAST TRADING DAY
//      of the previous month (its close). Not the 1st, not a hardcoded "31st" —
//      the last trading day, whatever date that lands on.
//   2. ENTRY — on the FIRST TRADING DAY of the new month we enter each name at
//      the AVERAGE of its first 15-minute candle's HIGH and LOW (09:15-09:30
//      IST):  entry = (first15High + first15Low) / 2.
//
// Both pieces are captured here into ONE file, public/data/cohort-entries.json,
// which becomes the single source of truth the whole dashboard reads:
//   - app.js               — cohort membership + cost basis (the entry price)
//   - pin-holdings.mjs      — which departed names to keep tracking
//   - scrape-live-prices.mjs — which tickers need a live quote
// so selection, pinning and live quotes can never disagree again.
//
// Why capture instead of compute: the entry price lives in intraday data that
// is NOT in the daily snapshot, and "first trading day" is a market fact, not a
// calendar one (holidays move it). Running this each trading morning and only
// writing once the real first-15-min bar exists makes the anchor holiday-proof
// automatically — the first day the market actually trades becomes the anchor.
//
// Shape written:
//   { generated_at, cohort_size,
//     months: { "2026-08": {
//       selectionDate: "2026-07-31",   // last trading day of the prior month
//       anchorDate:    "2026-08-03",   // first trading day of this month
//       tickers: ["SJS", ...],         // top-7 by the selection-day close
//       entries: { "SJS": 2380.55, ... } // (first15High+first15Low)/2
//     } } }
//
// Usage:
//   node scrape-first15-entry.mjs                         # auto (run each trading morning)
//   node scrape-first15-entry.mjs --month 2026-08 \
//        --anchor 2026-08-03 --selection 2026-07-31       # backfill / force
//   node scrape-first15-entry.mjs --dry                   # print, don't write

import { readFileSync, writeFileSync, existsSync, readdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname   = dirname(fileURLToPath(import.meta.url));
const DATA_DIR    = resolve(__dirname, "../public/data");
const SNAP_DIR    = resolve(DATA_DIR, "snapshots");
const OUT_PATH    = resolve(DATA_DIR, "cohort-entries.json");
const IST_OFFSET  = 19800; // +5:30 in seconds
const COHORT_SIZE = Number(process.env.COHORT_SIZE || 7);

const args = process.argv.slice(2);
const getArg = (k) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : null; };
const DRY = args.includes("--dry");
const forceMonth     = getArg("--month");
const forceAnchor    = getArg("--anchor");
const forceSelection = getArg("--selection");

const istDateOf = (ms) => new Date(ms + IST_OFFSET * 1000).toISOString().slice(0, 10);
const istToday  = () => istDateOf(Number(getArg("--now")) || Date.now());
const isWeekday = (d) => { const dow = new Date(d + "T00:00:00Z").getUTCDay(); return dow >= 1 && dow <= 5; };
const readJson  = (p) => { try { return JSON.parse(readFileSync(p, "utf8")); } catch { return null; } };

// All snapshot dates on disk, ascending. Weekend snapshots (historical
// leftovers) are dropped so "last trading day" can never land on a Saturday.
function snapshotDates() {
  if (!existsSync(SNAP_DIR)) return [];
  return readdirSync(SNAP_DIR)
    .filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .map((f) => f.slice(0, 10))
    .filter(isWeekday)
    .sort();
}

run().catch((e) => { console.error("Fatal:", e.stack || e.message); process.exit(1); });

async function run() {
  const dates = snapshotDates();
  if (!dates.length) { console.log("No snapshots yet — nothing to capture."); return; }

  const today = istToday();
  const month = forceMonth || today.slice(0, 7);

  const store = readJson(OUT_PATH) || {};
  store.months = store.months || {};
  const existing = store.months[month] || null;

  // ANCHOR — the first trading day of `month`. Forced value wins (backfill);
  // otherwise it's today, but only once the market has actually traded (a real
  // first-15-min bar exists). We only ever set the anchor for the FIRST such
  // day: once entries are recorded for the month we leave it alone.
  let anchorDate = forceAnchor || existing?.anchorDate || null;
  if (!anchorDate) {
    if (today.slice(0, 7) !== month) {
      console.log(`Auto mode: today (${today}) is not in ${month}; nothing to do.`);
      return;
    }
    if (!isWeekday(today)) { console.log(`${today} is a weekend — market shut, nothing to capture.`); return; }
    anchorDate = today; // provisional; confirmed below only if a real bar exists
  }

  // SELECTION — last trading day strictly before this month. Forced value wins.
  let selectionDate = forceSelection || existing?.selectionDate || null;
  if (!selectionDate) {
    const monthStart = `${month}-01`;
    selectionDate = [...dates].reverse().find((d) => d < monthStart) || null;
  }
  if (!selectionDate) { console.log(`No prior-month snapshot before ${month}; cannot select a cohort yet.`); return; }

  // Cohort — top-COHORT_SIZE by composite from the selection-day snapshot.
  const tickers = existing?.tickers?.length ? existing.tickers : pickTop7(selectionDate);
  if (!tickers.length) { console.log(`Selection snapshot ${selectionDate} has no rankable stocks.`); return; }

  console.log(`Month ${month} · selection ${selectionDate} · anchor ${anchorDate}`);
  console.log(`  cohort (${tickers.length}): ${tickers.join(", ")}`);

  // If every entry is already captured, this is a no-op (safe to run daily).
  const prevEntries = existing?.entries || {};
  const todo = tickers.filter((t) => prevEntries[t] == null);
  if (!todo.length && existing) {
    console.log("  all entries already captured — nothing to do.");
    return;
  }
  console.log(`  to capture: ${todo.join(", ")}`);

  const entries = { ...prevEntries };
  let gotABar = Object.keys(prevEntries).length > 0; // backfill of an existing partial
  for (const t of todo) {
    process.stdout.write(`  ${t}... `);
    try {
      const hl = await first15HL(`${t}.NS`, anchorDate);
      if (!hl) { console.log("no first-15-min bar yet"); continue; }
      const entry = Number(((hl.high + hl.low) / 2).toFixed(2));
      entries[t] = entry;
      gotABar = true;
      console.log(`H ${hl.high} · L ${hl.low} → entry ${entry}`);
    } catch (e) { console.log(`FAILED: ${e.message}`); }
  }

  // Auto mode with no forced anchor: if the market has not traded yet today
  // (no bar for anyone), do NOT record — this is a non-trading day (weekend
  // already excluded, so most likely a holiday). Try again next morning; the
  // real first trading day will be the one that finally produces bars.
  if (!forceAnchor && !existing && !gotABar) {
    console.log(`No first-15-min bars on ${anchorDate} — not a trading day (holiday). Will retry next morning.`);
    return;
  }

  const captured = tickers.filter((t) => entries[t] != null).length;
  const complete = captured === tickers.length;
  console.log(`\nCaptured ${captured}/${tickers.length}${complete ? " — complete" : " — partial (rerun to fill the rest)"}`);

  store.generated_at = new Date().toISOString();
  store.cohort_size  = COHORT_SIZE;
  store.months[month] = { selectionDate, anchorDate, tickers, entries };

  if (DRY) { console.log("\n[dry run — not writing]\n" + JSON.stringify(store.months[month], null, 2)); return; }
  writeFileSync(OUT_PATH, JSON.stringify(store, null, 2) + "\n");
  console.log(`Wrote ${OUT_PATH}`);
}

// Top-COHORT_SIZE tickers by composite from a day's snapshot (rated,
// data-complete, not hard-failed). Mirror of app.js pickTop7 selection.
function pickTop7(dateStr) {
  const snap = readJson(resolve(SNAP_DIR, `${dateStr}.json`));
  if (!snap) return [];
  return (snap.stocks || [])
    .filter((s) => s.ticker && typeof s.composite === "number" && s.dataComplete && !s.hardFailed)
    .sort((a, b) => b.composite - a.composite)
    .slice(0, COHORT_SIZE)
    .map((s) => String(s.ticker).toUpperCase());
}

// High AND low of the FIRST 15-min candle (09:15-09:30 IST) on dateStr, Yahoo.
async function first15HL(symbol, dateStr) {
  const dayStartUTC = Math.floor(Date.UTC(+dateStr.slice(0, 4), +dateStr.slice(5, 7) - 1, +dateStr.slice(8, 10)) / 1000) - IST_OFFSET;
  const period1 = dayStartUTC - 2 * 86400, period2 = dayStartUTC + 2 * 86400;
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?period1=${period1}&period2=${period2}&interval=15m`;
  const headers = { "User-Agent": "Mozilla/5.0 (compatible; VNSmallcapBot/1.0)" };
  let attempt = 0, lastErr;
  while (attempt < 3) {
    try {
      const r = await fetch(url, { headers });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const j = await r.json();
      const res = j?.chart?.result?.[0];
      const ts = res?.timestamp || [];
      const highs = res?.indicators?.quote?.[0]?.high || [];
      const lows  = res?.indicators?.quote?.[0]?.low  || [];
      // Earliest bar whose IST date matches dateStr = the 09:15 candle.
      let best = null;
      for (let i = 0; i < ts.length; i++) {
        if (highs[i] == null || lows[i] == null) continue;
        const istDate = istDateOf(ts[i] * 1000);
        if (istDate !== dateStr) continue;
        if (best == null || ts[i] < best.ts) best = { ts: ts[i], high: highs[i], low: lows[i] };
      }
      return best ? { high: Number(best.high.toFixed(2)), low: Number(best.low.toFixed(2)) } : null;
    } catch (e) { lastErr = e; attempt++; await new Promise((r) => setTimeout(r, 800 * attempt)); }
  }
  throw lastErr || new Error("fetch failed");
}
