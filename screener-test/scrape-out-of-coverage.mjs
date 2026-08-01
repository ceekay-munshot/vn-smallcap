#!/usr/bin/env node
// Out-of-coverage OHLCV scraper. The LKP manual basket sometimes
// includes stocks below our ~Rs 12,500 Cr universe (e.g. ELECON). They
// don't appear in our daily snapshots, so the Performance Tracker
// silently excludes them from the Manual average — which biases the
// average toward the 6-of-7 stocks we DO cover.
//
// This script reads lkp-manual-picks.json, collects every ticker
// flagged `in_universe: false`, and pulls daily closes from Yahoo
// Finance. The dashboard merges these into the Manual cohort + the
// Accuracy hit tracker, so every client pick contributes to the
// average return.
//
// Run inside daily-refresh.yml; output is committed back to main.

import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const LKP_PATH = resolve(__dirname, "../public/data/lkp-manual-picks.json");
const OUT_PATH = resolve(__dirname, "../public/data/out-of-coverage-history.json");
const DAYS = 180;            // matches the snapshot trail window — gives 6 months of context

run().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  process.exit(1);
});

async function run() {
  if (!existsSync(LKP_PATH)) {
    console.log("No LKP manual picks file — nothing to scrape.");
    return;
  }
  const lkp = JSON.parse(readFileSync(LKP_PATH, "utf8"));

  // Collect every out-of-coverage ticker mentioned in the file (current
  // top-level `picks` array + any historic month-keyed entries under
  // `picksByMonth`). For LKP rows where `ticker` is null because we
  // couldn't match it to our coverage universe, the `selection` field
  // usually IS the NSE symbol (e.g. "ELECON").
  const tickers = new Set();
  const collect = (arr) => {
    for (const p of (arr || [])) {
      if (p.in_universe) continue;
      const t = (p.ticker || p.selection || "").toString().trim().toUpperCase();
      if (t && /^[A-Z][A-Z0-9-]*$/.test(t)) tickers.add(t);
    }
  };
  collect(lkp.picks);
  for (const m of Object.values(lkp.picksByMonth || {})) collect(m);

  if (tickers.size === 0) {
    console.log("No out-of-coverage tickers in LKP picks. Nothing to fetch.");
    // Write an empty payload so the dashboard's fetch doesn't 404.
    writeEmpty();
    return;
  }

  console.log(`Fetching ${tickers.size} out-of-coverage ticker${tickers.size === 1 ? "" : "s"} from Yahoo...`);
  const closesByTicker = {};
  for (const t of tickers) {
    process.stdout.write(`  ${t}.NS ... `);
    try {
      const closes = await fetchDailyCloses(`${t}.NS`, DAYS);
      closesByTicker[t] = closes;
      const dates = Object.keys(closes);
      console.log(`${dates.length} days (${dates[0] || "—"} → ${dates[dates.length - 1] || "—"})`);
    } catch (err) {
      console.log(`FAILED: ${err.message}`);
    }
    await new Promise((r) => setTimeout(r, 350));   // be nice to Yahoo
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: "Yahoo Finance · query1 chart endpoint",
    note: "Daily closes for LKP manual-pick stocks that fall outside our coverage universe. Used by the Performance Tracker + Accuracy view so out-of-cap picks count in the Manual average return and hit tracking.",
    ticker_count: Object.keys(closesByTicker).length,
    tickers: closesByTicker,
  };
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\nWrote ${OUT_PATH}`);
}

function writeEmpty() {
  const payload = {
    generated_at: new Date().toISOString(),
    source: "Yahoo Finance · query1 chart endpoint",
    note: "No out-of-coverage tickers detected.",
    ticker_count: 0,
    tickers: {},
  };
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
}

async function fetchDailyCloses(symbol, days) {
  const end = Math.floor(Date.now() / 1000);
  const start = end - days * 86400;
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?period1=${start}&period2=${end}&interval=1d&events=history`;
  const headers = { "User-Agent": "Mozilla/5.0 (compatible; KLPDashboardBot/1.0)" };
  let attempt = 0, lastErr;
  while (attempt < 3) {
    try {
      const r = await fetch(url, { headers });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const j = await r.json();
      const result = j?.chart?.result?.[0];
      if (!result?.timestamp) throw new Error("no data");
      const ts = result.timestamp;
      const closes = result.indicators?.quote?.[0]?.close || [];
      const out = {};
      for (let i = 0; i < ts.length; i++) {
        if (closes[i] == null) continue;
        const date = new Date(ts[i] * 1000).toISOString().slice(0, 10);
        out[date] = Number(closes[i].toFixed(2));
      }
      return out;
    } catch (err) {
      lastErr = err;
      attempt++;
      await new Promise((r) => setTimeout(r, 800 * attempt));
    }
  }
  throw lastErr || new Error("fetch failed");
}
