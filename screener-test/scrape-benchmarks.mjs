#!/usr/bin/env node
// Benchmark history scraper. Pulls daily close prices for the indices we
// compare cohort returns against. Output is read by the Strategy /
// History tab to compute alpha (pick return − benchmark return) over the
// same window as each monthly top-7 cohort.
//
// Cheap: one Yahoo chart endpoint call per index, gets the full 90-day
// series in one go. No paid APIs.

import { writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_PATH = resolve(__dirname, "../public/data/benchmark-history.json");

// Smallcap 250 is the primary benchmark — it covers NSE ranks 251-500,
// which is where a Rs 2,000-12,500 Cr universe actually lives. Midcap 150
// is secondary, because the Rs 12,500 Cr ceiling straddles into it.
//
// Nifty 50 and Nifty 500 are retained because app.js still looks them up
// by literal symbol key (benchmark.indices["^NSEI"] / ["^CRSLDX"]) —
// dropping them here would silently blank those chart lines. Repoint
// app.js at Smallcap 250 in the UI pass, then these can go.
// Bank Nifty removed: nothing reads it.
const INDICES = [
  { symbol: "NIFTYSMLCAP250.NS", label: "Nifty Smallcap 250" },
  { symbol: "NIFTYMIDCAP150.NS", label: "Nifty Midcap 150" },
  { symbol: "^NSEI",             label: "Nifty 50" },
  { symbol: "^CRSLDX",           label: "Nifty 500" },
  // India VIX — the regime overlay's second input, and the only one of the
  // founder's "market momentum" asks with a free, stable daily feed.
  { symbol: "^INDIAVIX",         label: "India VIX" },
];
const DAYS = 1120;        // ~760 trading days. 420 gave ~280, enough for the
                          // Trade Plan's 200 DMA but far short of the 504-day
                          // strategy backtest, which left the regime overlay
                          // and the VIX band blind over most of its window.

run().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  process.exit(1);
});

async function run() {
  const indices = {};
  for (const { symbol, label } of INDICES) {
    process.stdout.write(`Fetching ${label} (${symbol})... `);
    try {
      const closes = await fetchDailyCloses(symbol, DAYS);
      indices[symbol] = { label, closes };
      const dates = Object.keys(closes);
      console.log(`${dates.length} days (${dates[0]} → ${dates[dates.length - 1]})`);
    } catch (err) {
      console.log(`FAILED: ${err.message}`);
    }
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: "Yahoo Finance · query1 chart endpoint",
    indices,
  };

  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\nWrote ${OUT_PATH}`);
}

// Yahoo chart endpoint → { date(YYYY-MM-DD) → close } map.
// Filters out NaN closes (Yahoo emits null for non-trading days
// inside the range; safer to drop than to keep the gap).
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
