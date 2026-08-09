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
// Per-request ceiling. Without one, a hanging endpoint blocks fetch()
// forever: on 3 Aug the quote API stopped responding, the job sat until
// its 10-minute limit, was cancelled, and wrote nothing at all -- so the
// dashboard showed Friday's closes through a live session and called them
// current. A slow API must degrade to partial data, never to no data.
const REQ_TIMEOUT_MS = 8000;


const readJsonSafe = (p) => { try { return JSON.parse(readFileSync(p, "utf8")); } catch { return null; } };

// The tickers the dashboard is actually holding this month, so every basket
// name gets a live quote. Prefer the captured cohort-entries record for the
// current month (the locked framework basket, whose selection day and
// cost-basis day differ, so it can't be re-derived from one snapshot). Fall
// back to the first-snapshot-of-month top-N for a month not yet captured, and
// to [] before tracking starts (safe no-op). Mirror of the app's resolution.
function currentCohortTickers() {
  if (!existsSync(SNAP_DIR)) return [];
  const dates = readdirSync(SNAP_DIR)
    .filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .map((f) => f.slice(0, 10))
    .sort();
  if (!dates.length) return [];
  const ym = dates[dates.length - 1].slice(0, 7);

  // Captured cohort for this month wins -- but only when COMPLETE (every held
  // name has an entry), so a partial capture never diverges from the dashboard,
  // which also falls back until the record fills. Then append the next-ranked
  // SELECTION names as substitutes, so ranks 8..COHORT_SIZE still get quotes
  // (the 7-plus-3 rule: on entry morning we must be able to price a substitute
  // when a core order is missed).
  const ce = readJsonSafe(resolve(DATA_DIR, "cohort-entries.json"));
  const rec = ce?.months?.[ym];
  const recReady = rec && Array.isArray(rec.tickers) && rec.tickers.length
    && rec.tickers.every((t) => typeof rec.entries?.[t] === "number");
  if (recReady) {
    const held = rec.tickers.map((t) => String(t).toUpperCase());
    const out = [...held];
    const heldSet = new Set(held);
    const selSnap = rec.selectionDate ? readJsonSafe(resolve(SNAP_DIR, `${rec.selectionDate}.json`)) : null;
    for (const t of (selSnap?.stocks || [])
      .filter((x) => x.composite != null && x.dataComplete && !x.hardFailed)
      .sort((a, b) => b.composite - a.composite)
      .map((x) => String(x.ticker || "").toUpperCase())) {
      if (out.length >= COHORT_SIZE) break;
      if (t && !heldSet.has(t)) { out.push(t); heldSet.add(t); }
    }
    return out;
  }

  // Fallback: first snapshot of the latest month, top-N by composite.
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
  const srcCounts = { munshot: 0, yahoo: 0 };
  let ok = 0, fail = 0;
  for (const ticker of list) {
    const q = await fetchQuote(ticker);
    if (q) { srcCounts[q.source] = (srcCounts[q.source] || 0) + 1; delete q.source; prices[ticker] = q; ok++; process.stdout.write(q === null ? "x" : "."); }
    else   { fail++; process.stdout.write("x"); }
    await new Promise((r) => setTimeout(r, 200));
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: srcCounts.yahoo ? `Munshot quote API (${srcCounts.munshot} tickers) + Yahoo fallback (${srcCounts.yahoo})` : "Munshot quote API (fastapi.muns.io/stock-data)",
    note: "Live intraday snapshot per basket ticker. dayHigh/dayLow drive intraday target/SL touch detection; current is the latest traded price.",
    ticker_count: Object.keys(prices).length,
    prices,
  };
  mkdirSync(DATA_DIR, { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\nWrote ${OUT_PATH} — ${ok} ok (${srcCounts.munshot} munshot, ${srcCounts.yahoo} yahoo), ${fail} failed.`);
  // A run that fetched nothing must not be mistaken for a successful one:
  // it would overwrite good quotes with an empty file.
  if (ok === 0) { console.error("No quotes retrieved from either source — leaving the previous file in place."); process.exit(1); }
}

// One quote → { current, open, prevClose, dayHigh, dayLow, week52High,
// week52Low, ma50, ma200, vol10d, marketCap, yearlyChangePct }. Returns
// null on any error so a single bad ticker never aborts the run.
// Munshot first, Yahoo second. Yahoo's chart meta carries every field this
// file needs -- last price, the day's high and low, previous close -- so a
// dead primary costs freshness of source, not the feed itself.
// Once the primary has timed out twice in a row it is treated as down for
// the rest of the run. Retrying it per ticker costs 24 seconds each and
// pushed a ten-ticker job past its timeout -- the exact failure this
// fallback exists to prevent.
let munshotDown = false, munshotMisses = 0;

async function fetchQuote(ticker) {
  const primary = munshotDown ? null : await fetchFromMunshot(ticker);
  if (!primary) {
    if (++munshotMisses >= 2 && !munshotDown) {
      munshotDown = true;
      console.log("\n  primary quote API unresponsive — using Yahoo for the rest of this run");
    }
  } else munshotMisses = 0;
  if (primary) return { ...primary, source: "munshot" };
  const fallback = await fetchFromYahoo(ticker);
  return fallback ? { ...fallback, source: "yahoo" } : null;
}

async function fetchFromYahoo(ticker) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}.NS?range=1d&interval=5m`;
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const r = await fetch(url, {
        headers: { "User-Agent": "Mozilla/5.0 (compatible; VNSmallcapBot/1.0)" },
        signal: AbortSignal.timeout(REQ_TIMEOUT_MS),
      });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const m = (await r.json())?.chart?.result?.[0]?.meta;
      if (!m || m.regularMarketPrice == null) throw new Error("no meta");
      return {
        current: m.regularMarketPrice,
        open: m.regularMarketOpen ?? null,
        prevClose: m.chartPreviousClose ?? m.previousClose ?? null,
        dayHigh: m.regularMarketDayHigh ?? null,
        dayLow: m.regularMarketDayLow ?? null,
        week52High: m.fiftyTwoWeekHigh ?? null,
        week52Low: m.fiftyTwoWeekLow ?? null,
        ma50: null, ma200: null, vol10d: null,
        marketCap: null, yearlyChangePct: null,
      };
    } catch {
      if (attempt === 1) return null;
      await new Promise((res) => setTimeout(res, 500));
    }
  }
  return null;
}

async function fetchFromMunshot(ticker) {
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const r = await fetch(API, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ticker_symbol: ticker, type: "stockquote", country: "india" }),
        signal: AbortSignal.timeout(REQ_TIMEOUT_MS),
      });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const body = await r.json();               // API returns a quoted CSV string
      if (typeof body !== "string") throw new Error("unexpected shape");
      return parseQuote(body);
    } catch (e) {
      if (attempt === 1) return null;
      await new Promise((res) => setTimeout(res, 400));
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

// Invoked last: the circuit-breaker state above is let-bound, so calling
// run() earlier in the file hits the temporal dead zone.
run().catch((e) => { console.error("Fatal:", e.stack || e.message); process.exit(1); });
