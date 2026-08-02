#!/usr/bin/env node
// Live intraday price feed for the CURRENT MONTH'S COHORT. Calls the
// Munshot quote API (no auth) for each held ticker and writes a compact
// public/data/live-prices.json the dashboard reads to:
//   - show the current intraday price instead of yesterday's close, and
//   - detect target / stop-loss TOUCHES from the day's high / low, so a
//     level counts as hit the moment price touches it intraday rather
//     than only when it closes through.
//
// The cohort is derived from the snapshot trail with the same rule the
// dashboard uses (app.js pickTop7): top COHORT_SIZE by composite among
// rated, data-complete, non-hard-failed stocks in the FIRST snapshot of
// the current month. This used to read lkp-manual-picks.json -- a client
// basket that does not exist here -- so the workflow was disabled.
//
// Pinned holdings are included too: a pick that has crossed out of the
// market-cap band is still held, and still needs a live price.
//
// Reads:  public/data/snapshots/  + public/data/pinned-companies.json
// Writes: public/data/live-prices.json
//
// Usage: node screener-test/scrape-live-prices.mjs [EXTRA,TICKERS]

import { readFileSync, writeFileSync, mkdirSync, readdirSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR  = resolve(__dirname, "../public/data");
const SNAP_DIR  = resolve(DATA_DIR, "snapshots");
// 7 held + 3 buffer. The buffers need live prices too: the whole point is
// to know, on entry morning, which core name we missed and what to buy
// instead -- which is impossible without a live quote for the substitute.
const COHORT_SIZE = Number(process.env.LIVE_COHORT_SIZE || 10);
const OUT_PATH  = resolve(DATA_DIR, "live-prices.json");
const API       = "https://fastapi.muns.io/stock-data";

run().catch((e) => { console.error("Fatal:", e.stack || e.message); process.exit(1); });

// Mirror of app.js pickTop7 — keep the two in step. Returns [] when no
// snapshot exists yet, which makes this a safe no-op before tracking starts.
function currentCohortTickers() {
  if (!existsSync(SNAP_DIR)) return [];
  const dates = readdirSync(SNAP_DIR)
    .filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .map((f) => f.slice(0, 10))
    .sort();
  if (!dates.length) return [];
  // First snapshot of the LATEST month present — the month-start basket.
  const ym = dates[dates.length - 1].slice(0, 7);
  const first = dates.find((d) => d.slice(0, 7) === ym);
  let snap;
  try { snap = JSON.parse(readFileSync(resolve(SNAP_DIR, `${first}.json`), "utf8")); }
  catch { return []; }
  return (snap.stocks || [])
    .filter((x) => x.composite != null && x.dataComplete && !x.hardFailed)
    .sort((a, b) => b.composite - a.composite)
    .slice(0, COHORT_SIZE)
    .map((x) => String(x.ticker || "").toUpperCase())
    .filter(Boolean);
}

async function run() {
  // This month's cohort, plus any explicitly-listed extras.
  const tickers = new Set(currentCohortTickers());
  for (const t of (process.argv[2] || "").split(",").map((s) => s.trim().toUpperCase()).filter(Boolean)) {
    tickers.add(t);
  }

  const list = [...tickers];
  console.log(`Fetching live quotes for ${list.length} tickers: ${list.join(", ")}`);

  const prices = {};
  let ok = 0, fail = 0;
  for (const ticker of list) {
    const q = await fetchQuote(ticker);
    if (q) { prices[ticker] = q; ok++; process.stdout.write("."); }
    else   { fail++; process.stdout.write("x"); }
    await new Promise((r) => setTimeout(r, 200));
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: "Munshot quote API (fastapi.muns.io/stock-data)",
    note: "Live intraday snapshot per basket ticker. dayHigh/dayLow drive intraday target/SL touch detection; current is the latest traded price.",
    ticker_count: Object.keys(prices).length,
    prices,
  };
  mkdirSync(DATA_DIR, { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\nWrote ${OUT_PATH} — ${ok} ok, ${fail} failed.`);
}

// One quote → { current, open, prevClose, dayHigh, dayLow, week52High,
// week52Low, ma50, ma200, vol10d, marketCap, yearlyChangePct }. Returns
// null on any error so a single bad ticker never aborts the run.
async function fetchQuote(ticker) {
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const r = await fetch(API, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ticker_symbol: ticker, type: "stockquote", country: "india" }),
      });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const body = await r.json();               // API returns a quoted CSV string
      if (typeof body !== "string") throw new Error("unexpected shape");
      return parseQuote(body);
    } catch (e) {
      if (attempt === 2) return null;
      await new Promise((res) => setTimeout(res, 600 * (attempt + 1)));
    }
  }
  return null;
}

// "Current Price=1341.8,...,Day Range=1268.9 - 1359.0,..." → structured.
function parseQuote(str) {
  const kv = {};
  for (const part of str.split(",")) {
    const i = part.indexOf("=");
    if (i < 0) continue;
    kv[part.slice(0, i).trim()] = part.slice(i + 1).trim();
  }
  const num = (v) => { const n = Number(String(v).replace(/[^0-9.\-]/g, "")); return Number.isFinite(n) ? n : null; };
  const range = (v) => {
    const m = String(v || "").split("-").map((s) => num(s));
    return m.length === 2 && m[0] != null && m[1] != null ? { lo: Math.min(m[0], m[1]), hi: Math.max(m[0], m[1]) } : null;
  };
  const day = range(kv["Day Range"]);
  const wk  = range(kv["52-Week Range"]);
  const current = num(kv["Current Price"]);
  if (current == null) return null;
  return {
    current,
    open: num(kv["Opening Price"]),
    prevClose: num(kv["Previous Close"]),
    dayHigh: day?.hi ?? null,
    dayLow: day?.lo ?? null,
    week52High: wk?.hi ?? null,
    week52Low: wk?.lo ?? null,
    ma50: num(kv["50-Day Moving Average"]),
    ma200: num(kv["200-Day Moving Average"]),
    vol10d: num(kv["10-Day Average Volume"]),
    marketCap: num(kv["Market Cap"]),
    yearlyChangePct: num(kv["Yearly Change (%)"]),
  };
}
