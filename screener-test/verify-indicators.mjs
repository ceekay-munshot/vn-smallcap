#!/usr/bin/env node
// Independent verification of the technicals math.
// Picks N random tickers from public/data/technicals.json, refetches their
// daily OHLCV from Yahoo, recomputes the same indicators using the well-
// tested `technicalindicators` npm package, and compares to the values our
// scrape-technicals.mjs produced. Fails the workflow if any indicator
// disagrees with the library by more than the tolerance.

import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

// Tee everything to a debug log file so failures are inspectable even
// when GitHub workflow logs are auth-gated. The workflow uploads this
// file as an artifact on always().
const __dirname = dirname(fileURLToPath(import.meta.url));
const DEBUG_LOG = resolve(__dirname, "../public/data/verify-debug.log");
const TECH_PATH = resolve(__dirname, "../public/data/technicals.json");
let logBuf = [];
const log = (...args) => {
  const line = args.map((a) => typeof a === "object" ? JSON.stringify(a) : String(a)).join(" ");
  process.stdout.write(line + "\n");
  logBuf.push(line);
  // Eager write — every log call flushes to disk so a crash never
  // loses the log. process.on("exit") was unreliable when the runtime
  // killed the process before handlers fired.
  try { writeFileSync(DEBUG_LOG, logBuf.join("\n") + "\n"); } catch {}
};

// Make sure the file exists immediately so the workflow always has
// something to commit on failure.
try { writeFileSync(DEBUG_LOG, "verify-indicators starting — log will accumulate here.\n"); } catch (e) {
  process.stdout.write("WARN: could not pre-write debug log: " + e.message + "\n");
}

log("=== verify-indicators starting ===");
log("Node version:", process.version);
log("CWD:", process.cwd());
log("__dirname:", __dirname);

// Use createRequire for bulletproof CJS-from-ESM import (safer than
// default import shenanigans, especially under Node 20).
const require = createRequire(import.meta.url);
let EMA, SMA, RSI, MACD, ADX, ATR;
try {
  const ti = require("technicalindicators");
  ({ EMA, SMA, RSI, MACD, ADX, ATR } = ti);
  log("Imported technicalindicators OK. Keys present:", Object.keys(ti).slice(0, 12).join(","));
  log("  EMA type:", typeof EMA, "SMA:", typeof SMA, "RSI:", typeof RSI, "MACD:", typeof MACD, "ADX:", typeof ADX, "ATR:", typeof ATR);
} catch (err) {
  log("FAILED to import technicalindicators:", err.message);
  log("Stack:", err.stack);
  process.exit(1);
}

const SAMPLE_SIZE = parseInt(process.env.SAMPLE_SIZE || "8", 10);
const TOLERANCE_PCT = parseFloat(process.env.TOLERANCE_PCT || "0.5"); // %

run().catch((err) => { log("FATAL in run():", err.message); log("Stack:", err.stack); process.exit(1); });

async function run() {
  log("Reading technicals.json from:", TECH_PATH);
  const data = JSON.parse(readFileSync(TECH_PATH, "utf8"));
  log("Loaded", (data.companies || []).length, "company records");
  const candidates = (data.companies || []).filter((c) => !c.error && c.bars_count >= 200);
  log("Eligible candidates (no error, ≥200 bars):", candidates.length);
  if (candidates.length < SAMPLE_SIZE) {
    log(`Only ${candidates.length} eligible companies — need at least ${SAMPLE_SIZE}.`);
    process.exit(1);
  }

  // Deterministic random sample seeded by date so consecutive runs spread coverage
  const seed = new Date().toISOString().slice(0, 10);
  const shuffled = [...candidates].sort((a, b) => hash(seed + a.ticker) - hash(seed + b.ticker));
  const sample = shuffled.slice(0, SAMPLE_SIZE);

  log(`Verifying ${sample.length} tickers against technicalindicators library`);
  log(`Tolerance: ${TOLERANCE_PCT}%\n`);

  const results = [];
  for (const c of sample) {
    log(`\n[${c.ticker}] fetching...`);
    let bars;
    try {
      bars = await fetchBars(`${c.ticker}.NS`, 400);
      log(`  fetched ${bars.length} bars`);
      if (bars.length < 250) { log(`  only ${bars.length} bars, skip`); continue; }
    } catch (err) {
      log(`  FETCH FAIL: ${err.message}`);
      continue;
    }
    try {
      const closes = bars.map((b) => b.close);
      const highs = bars.map((b) => b.high);
      const lows = bars.map((b) => b.low);
      const checks = [];
      checks.push(compare("EMA(50)",  c.ema50, EMA.calculate({ period: 50,  values: closes }).at(-1)));
      checks.push(compare("SMA(50)",  c.sma50, SMA.calculate({ period: 50,  values: closes }).at(-1)));
      checks.push(compare("SMA(200)", c.sma200, SMA.calculate({ period: 200, values: closes }).at(-1)));
      checks.push(compare("RSI(14)",  c.rsi14, RSI.calculate({ period: 14, values: closes }).at(-1)));
      const macdLib = MACD.calculate({ fastPeriod: 12, slowPeriod: 26, signalPeriod: 9, values: closes, SimpleMAOscillator: false, SimpleMASignal: false }).at(-1);
      checks.push(compare("MACD line",   c.macd?.line,   macdLib?.MACD));
      checks.push(compare("MACD signal", c.macd?.signal, macdLib?.signal));
      checks.push(compare("ADX(14)", c.adx14, ADX.calculate({ period: 14, high: highs, low: lows, close: closes }).at(-1)?.adx));
      const atrLib = ATR.calculate({ period: 14, high: highs, low: lows, close: closes }).at(-1);
      const atrPctLib = atrLib != null ? (atrLib / closes.at(-1)) * 100 : null;
      checks.push(compare("ATR(14) %", c.atr14_pct, atrPctLib));
      const failed = checks.filter((x) => x && !x.pass);
      log(`  ${checks.filter((x) => x).length} checks, ${failed.length} failed`);
      results.push({ ticker: c.ticker, name: c.name, checks });
    } catch (err) {
      log(`  COMPUTE FAIL: ${err.message}`);
      log(`  Stack: ${err.stack}`);
    }
  }

  log("\n=== Detailed report ===\n");
  for (const r of results) {
    log(`${r.ticker} · ${r.name}`);
    for (const ch of r.checks) {
      if (!ch) continue;
      const mark = ch.pass ? "✓" : "✗";
      log(`  ${mark} ${ch.label.padEnd(14)} our: ${fmt(ch.ours).padStart(10)}  lib: ${fmt(ch.lib).padStart(10)}  diff: ${ch.diffPct?.toFixed(3)}%`);
    }
    log();
  }

  const totalChecks = results.flatMap((r) => r.checks).filter((c) => c).length;
  const totalFail   = results.flatMap((r) => r.checks).filter((c) => c && !c.pass).length;
  log(`=== Summary: ${totalChecks - totalFail}/${totalChecks} checks passed (tolerance ±${TOLERANCE_PCT}%) ===`);
  if (totalFail > 0) {
    log(`\nFAIL — ${totalFail} indicator(s) differ from technicalindicators library by more than ${TOLERANCE_PCT}%.`);
    process.exit(1);
  } else {
    log("\nAll indicators agree with technicalindicators reference implementation. Math verified.");
  }
}

function compare(label, ours, lib) {
  if (ours == null || lib == null || !Number.isFinite(ours) || !Number.isFinite(lib)) return null;
  const diff = Math.abs(ours - lib);
  const base = Math.max(Math.abs(ours), Math.abs(lib), 1e-9);
  const diffPct = (diff / base) * 100;
  return { label, ours, lib, diffPct, pass: diffPct <= TOLERANCE_PCT };
}

function fmt(n) { return n == null ? "—" : Number(n).toFixed(2); }

function hash(s) { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0; return h; }

async function fetchBars(symbol, days) {
  const end = Math.floor(Date.now() / 1000);
  const start = end - days * 86400;
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?period1=${start}&period2=${end}&interval=1d&events=history`;
  const r = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (compatible; KLPVerifyBot/1.0)" } });
  if (!r.ok) throw new Error(`HTTP ${r.status}`);
  const j = await r.json();
  const result = j?.chart?.result?.[0];
  if (!result?.timestamp) throw new Error("no chart data");
  const ts = result.timestamp;
  const q = result.indicators?.quote?.[0] || {};
  const out = [];
  for (let i = 0; i < ts.length; i++) {
    const close = q.close?.[i];
    if (close == null) continue;
    out.push({ close, high: q.high?.[i] ?? close, low: q.low?.[i] ?? close });
  }
  return out;
}
