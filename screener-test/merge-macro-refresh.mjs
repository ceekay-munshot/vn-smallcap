#!/usr/bin/env node
// Merge weekly macro refresh from Claude.ai routine into the production
// macro-context.json. Each indicator is sanity-checked against bounds —
// if validation fails, the existing value is preserved instead of being
// overwritten with bad data. This guards against hallucinated extractions.
//
// Inputs:
//   - public/data/weekly-macro-refresh.json  (routine output, schema below)
//   - screener-test/static/macro-context.json  (production source-of-truth)
//
// Output:
//   - Mutates the `economic`, `regime`, and `sentiment.put_call_ratio` fields
//     in macro-context.json. Sector themes, PLI/renewable lists, etc. are
//     untouched — those remain hand-maintained.
//
// The next daily-refresh's scrape-macro.mjs picks up the new context values
// and writes them into public/data/macro.json which the dashboard reads.

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REFRESH_PATH = resolve(__dirname, "../public/data/weekly-macro-refresh.json");
const CONTEXT_PATH = resolve(__dirname, "static/macro-context.json");

// Sanity bounds per indicator — values outside these are rejected as
// likely-bad extractions and the existing context value is preserved.
const BOUNDS = {
  gdp_yoy:        { min: -5,  max: 15 },
  cpi:            { min: -2,  max: 25 },
  gsec10y:        { min: 4,   max: 12 },
  put_call_ratio: { min: 0.3, max: 3.0 },
};
const RBI_INFLATION_BAND_UPPER = 6.0;   // 4 ± 2% tolerance band
const MAX_STALE_DAYS = 120;             // routine claim of source-date older than this → rejected
const MAX_PCR_STALE_DAYS = 14;          // PCR is intraday — stricter

function inRange(v, b) {
  return typeof v === "number" && Number.isFinite(v) && v >= b.min && v <= b.max;
}
function ageDays(asOf) {
  if (!asOf) return Infinity;
  const t = new Date(asOf).getTime();
  return Number.isFinite(t) ? (Date.now() - t) / 86400_000 : Infinity;
}

if (!existsSync(REFRESH_PATH)) {
  console.log(`No routine output at ${REFRESH_PATH} — nothing to merge.`);
  process.exit(0);
}
const refresh = JSON.parse(readFileSync(REFRESH_PATH, "utf8"));
const context = JSON.parse(readFileSync(CONTEXT_PATH, "utf8"));
const indicators = refresh.indicators || {};

const summary = { accepted: [], rejected: [], skipped: [] };
const accept = (key, oldV, newV) => summary.accepted.push(`${key}: ${oldV} → ${newV}`);
const reject = (key, reason) => summary.rejected.push(`${key}: ${reason}`);

// ---------- GDP ----------
if (indicators.gdp_yoy) {
  const i = indicators.gdp_yoy;
  if (!inRange(i.value, BOUNDS.gdp_yoy))      reject("gdp_yoy", `value ${i.value} outside ${BOUNDS.gdp_yoy.min}–${BOUNDS.gdp_yoy.max}%`);
  else if (ageDays(i.as_of) > MAX_STALE_DAYS) reject("gdp_yoy", `as_of ${i.as_of} >${MAX_STALE_DAYS}d old`);
  else {
    accept("gdp_yoy", context.economic.gdp_yoy, i.value);
    context.economic.gdp_yoy = i.value;
    if (typeof i.trending_up === "boolean") context.economic.gdp_trending_up = i.trending_up;
  }
}

// ---------- CPI ----------
if (indicators.cpi) {
  const i = indicators.cpi;
  if (!inRange(i.value, BOUNDS.cpi))          reject("cpi", `value ${i.value} outside ${BOUNDS.cpi.min}–${BOUNDS.cpi.max}%`);
  else if (ageDays(i.as_of) > MAX_STALE_DAYS) reject("cpi", `as_of ${i.as_of} >${MAX_STALE_DAYS}d old`);
  else {
    accept("cpi", context.economic.cpi, i.value);
    context.economic.cpi = i.value;
    // Computed from the value — don't trust the routine to classify the band.
    context.economic.cpi_within_rbi_band = i.value <= RBI_INFLATION_BAND_UPPER;
  }
}

// ---------- 10Y G-Sec ----------
if (indicators.gsec10y) {
  const i = indicators.gsec10y;
  if (!inRange(i.value, BOUNDS.gsec10y))      reject("gsec10y", `value ${i.value} outside ${BOUNDS.gsec10y.min}–${BOUNDS.gsec10y.max}%`);
  else if (ageDays(i.as_of) > MAX_STALE_DAYS) reject("gsec10y", `as_of ${i.as_of} >${MAX_STALE_DAYS}d old`);
  else {
    accept("gsec10y", context.economic.gsec10y, i.value);
    context.economic.gsec10y = i.value;
    if (["rising", "declining", "stable"].includes(i.trend)) context.economic.gsec10y_trend = i.trend;
  }
}

// ---------- Regime flags (rate cut / gov capex / china+1) ----------
const regimeMap = [
  { key: "rate_cut_cycle", contextFlag: "rate_cut_cycle",        contextNote: "rate_cut_cycle_note" },
  { key: "gov_capex",      contextFlag: "gov_capex_active",      contextNote: "gov_capex_note" },
  { key: "china_plus_one", contextFlag: "china_plus_one_active", contextNote: "china_plus_one_note" },
];
for (const r of regimeMap) {
  const i = indicators[r.key];
  if (!i) continue;
  if (typeof i.active !== "boolean")          { reject(r.key, `active not boolean`); continue; }
  if (ageDays(i.as_of) > MAX_STALE_DAYS)      { reject(r.key, `as_of ${i.as_of} >${MAX_STALE_DAYS}d old`); continue; }
  accept(r.key, context.regime[r.contextFlag], i.active);
  context.regime[r.contextFlag] = i.active;
  if (i.note) context.regime[r.contextNote] = i.note;
}

// ---------- Rural recovery (signal enum, not boolean) ----------
if (indicators.rural_recovery) {
  const i = indicators.rural_recovery;
  const validSignals = ["good", "neutral", "weak"];
  if (!validSignals.includes(i.signal))       reject("rural_recovery", `signal "${i.signal}" not one of ${validSignals.join("/")}`);
  else if (ageDays(i.as_of) > MAX_STALE_DAYS) reject("rural_recovery", `as_of ${i.as_of} >${MAX_STALE_DAYS}d old`);
  else {
    accept("rural_recovery", context.regime.rural_recovery_signal, i.signal);
    context.regime.rural_recovery_signal = i.signal;
    if (i.note) context.regime.rural_recovery_note = i.note;
  }
}

// ---------- Put-Call Ratio ----------
if (indicators.put_call_ratio) {
  const i = indicators.put_call_ratio;
  if (!inRange(i.value, BOUNDS.put_call_ratio)) reject("put_call_ratio", `value ${i.value} outside ${BOUNDS.put_call_ratio.min}–${BOUNDS.put_call_ratio.max}`);
  else if (ageDays(i.as_of) > MAX_PCR_STALE_DAYS) reject("put_call_ratio", `as_of ${i.as_of} >${MAX_PCR_STALE_DAYS}d old (PCR is intraday)`);
  else {
    accept("put_call_ratio", context.sentiment.put_call_ratio, i.value);
    context.sentiment.put_call_ratio = i.value;
    context.sentiment.put_call_ratio_note = `Latest NSE option-chain PCR (refreshed weekly by routine). Source: ${i.source || "NSE"}, as of ${i.as_of}.`;
  }
}

// Routine-side skips passed through to the summary.
if (Array.isArray(refresh.skipped)) {
  for (const s of refresh.skipped) summary.skipped.push(typeof s === "string" ? s : JSON.stringify(s));
}

// Bookkeeping: stamp the run.
context._last_routine_update = refresh.generated_at || new Date().toISOString();

writeFileSync(CONTEXT_PATH, JSON.stringify(context, null, 2) + "\n");

console.log("=== Macro refresh merge ===");
console.log(`Routine generated_at: ${refresh.generated_at || "(missing)"}`);
console.log(`\nAccepted (${summary.accepted.length}):`);
summary.accepted.forEach((s) => console.log(`  ✓ ${s}`));
console.log(`\nRejected (${summary.rejected.length}):`);
summary.rejected.forEach((s) => console.log(`  ✗ ${s}`));
if (summary.skipped.length) {
  console.log(`\nRoutine-side skips (${summary.skipped.length}):`);
  summary.skipped.forEach((s) => console.log(`  - ${s}`));
}
console.log(`\nWrote ${CONTEXT_PATH}`);
