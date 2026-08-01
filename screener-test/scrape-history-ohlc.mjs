#!/usr/bin/env node
// 1-year technical-history backfill. For every ticker in our universe it
// pulls ~2 years of daily OHLC from Yahoo (enough warm-up for SMA200 etc.),
// computes the technical indicators + a 0-100 technical score per day, and
// stores a compact per-ticker (date, close, score) series over the most
// recent KEEP trading days. The Custom-tab technical back-test replays over
// this, so we can validate technical strategies over a full year without
// waiting for live data to accumulate.
//
// Usage: node scrape-history-ohlc.mjs [limit]   (limit = # tickers, for testing)

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { sma, ema, rsi, roc, macd, adx, highProximity, technicalScore } from "./lib/indicators.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAP_DIR = resolve(__dirname, "../public/data/snapshots");
const OUT_PATH = resolve(__dirname, "../public/data/history-technical.json");
const FETCH_DAYS = 1120;   // ~3 calendar years → ~760 trading days (warm-up + 2-yr window)
const KEEP = 504;          // trading days retained per ticker (~2 years)
const LIMIT = Number(process.argv[2]) || Infinity;

run().catch((e) => { console.error("Fatal:", e.stack || e.message); process.exit(1); });

async function run() {
  const idx = JSON.parse(readFileSync(resolve(SNAP_DIR, "index.json"), "utf8"));
  const latest = JSON.parse(readFileSync(resolve(SNAP_DIR, `${idx.dates[idx.dates.length - 1]}.json`), "utf8"));
  const universe = latest.stocks
    .filter((s) => s.ticker)
    .map((s) => ({ ticker: s.ticker, sector: s.sector || null }))
    .slice(0, LIMIT === Infinity ? undefined : LIMIT);

  console.log(`Backfilling technical history for ${universe.length} tickers (fetch ${FETCH_DAYS}d, keep ${KEEP})...`);
  const perTicker = {};
  let ok = 0, fail = 0;
  const allDates = new Set();

  for (const { ticker, sector } of universe) {
    try {
      const bars = await fetchOHLC(`${ticker}.NS`, FETCH_DAYS);
      if (!bars || bars.dates.length < 220) { fail++; process.stdout.write("."); continue; }
      const series = computeSeries(bars);
      perTicker[ticker] = { sector, ...series };
      for (const d of series.dates) allDates.add(d);
      ok++;
      if (ok % 25 === 0) process.stdout.write(` ${ok} `);
    } catch { fail++; process.stdout.write("x"); }
    await new Promise((r) => setTimeout(r, 250));
  }

  // Shared calendar = union of kept dates, most recent KEEP.
  const dates = [...allDates].sort().slice(-KEEP);
  const dateIdx = new Map(dates.map((d, i) => [d, i]));
  const tickers = {};
  for (const [ticker, s] of Object.entries(perTicker)) {
    const close = new Array(dates.length).fill(null);
    const score = new Array(dates.length).fill(null);
    for (let i = 0; i < s.dates.length; i++) {
      const j = dateIdx.get(s.dates[i]);
      if (j == null) continue;
      close[j] = s.close[i];
      score[j] = s.score[i];
    }
    tickers[ticker] = { sector: s.sector, close, score };
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: "Yahoo Finance daily OHLC → indicators computed in screener-test/lib/indicators.mjs",
    note: `Per-ticker daily close + 0-100 technical score over the last ${dates.length} trading days. Score = trend (price vs 50/200 MA) + momentum (RSI zone, MACD, ROC) + ADX + 52-week-high proximity. Powers the Custom-tab 1-year technical back-test.`,
    days: dates.length,
    ticker_count: Object.keys(tickers).length,
    dates,
    tickers,
  };
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload) + "\n");
  console.log(`\nWrote ${OUT_PATH} — ${ok} ok, ${fail} skipped, ${dates.length} days (${dates[0]} → ${dates[dates.length - 1]}).`);

  miniBacktest(payload);
}

function computeSeries(bars) {
  const { dates, open, high, low, close, volume } = bars;
  const rsi14 = rsi(close, 14);
  const s50 = sma(close, 50), s200 = sma(close, 200);
  const { hist } = macd(close);
  const roc20 = roc(close, 20);
  const adx14 = adx(high, low, close, 14);
  const hp = highProximity(close, 252);
  const score = close.map((c, i) => technicalScore({
    close: c, sma50: s50[i], sma200: s200[i], rsi: rsi14[i],
    macdHist: hist[i], roc20: roc20[i], adx: adx14[i], highProx: hp[i],
  }));
  // keep only rows with a valid score (indicators warmed up), last KEEP
  const rows = [];
  for (let i = 0; i < dates.length; i++) if (score[i] != null && close[i] != null) rows.push({ d: dates[i], c: Math.round(close[i] * 100) / 100, s: score[i] });
  const kept = rows.slice(-KEEP);
  return { dates: kept.map((r) => r.d), close: kept.map((r) => r.c), score: kept.map((r) => r.s) };
}

async function fetchOHLC(symbol, days) {
  const end = Math.floor(Date.now() / 1000), start = end - days * 86400;
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?period1=${start}&period2=${end}&interval=1d`;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const r = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (compatible; KLPDashboardBot/1.0)" } });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const res = (await r.json())?.chart?.result?.[0];
      if (!res?.timestamp) throw new Error("no data");
      const q = res.indicators?.quote?.[0] || {};
      const dates = [], open = [], high = [], low = [], close = [], volume = [];
      for (let i = 0; i < res.timestamp.length; i++) {
        const c = q.close?.[i];
        if (c == null || c <= 0) continue;                 // drop holidays / bad prints
        dates.push(new Date(res.timestamp[i] * 1000).toISOString().slice(0, 10));
        open.push(q.open?.[i] ?? c); high.push(q.high?.[i] ?? c); low.push(q.low?.[i] ?? c);
        close.push(c); volume.push(q.volume?.[i] ?? 0);
      }
      return { dates, open, high, low, close, volume };
    } catch (e) {
      if (attempt === 2) throw e;
      await new Promise((r) => setTimeout(r, 700 * (attempt + 1)));
    }
  }
}

// Quick sanity replay: held top-7 by technical score at the window start,
// equal-weight, tracked to the end. Prints return + drawdown.
function miniBacktest(payload) {
  const { dates, tickers } = payload;
  const day0 = 0, dayN = dates.length - 1;
  const ranked = Object.entries(tickers)
    .filter(([, t]) => t.score[day0] != null && t.close[day0] != null)
    .sort((a, b) => b[1].score[day0] - a[1].score[day0])
    .slice(0, 7);
  if (!ranked.length) { console.log("mini-backtest: no ranked names"); return; }
  let peak = -Infinity, maxDD = 0, lastRet = 0;
  const entry = ranked.map(([, t]) => t.close[day0]);
  for (let j = day0; j <= dayN; j++) {
    let sum = 0, n = 0;
    ranked.forEach(([, t], k) => { const c = t.close[j] ?? t.close[day0]; if (c != null) { sum += c / entry[k]; n++; } });
    if (!n) continue;
    lastRet = (sum / n - 1) * 100;
    const f = 1 + lastRet / 100; if (f > peak) peak = f; const dd = (f / peak - 1) * 100; if (dd < maxDD) maxDD = dd;
  }
  console.log(`mini-backtest (held top-7 by tech score, ${dates[0]} → ${dates[dayN]}): ${lastRet.toFixed(2)}% return, ${maxDD.toFixed(2)}% max drawdown`);
  console.log("  picks:", ranked.map(([tk, t]) => `${tk}(${t.score[day0]})`).join(", "));
}
