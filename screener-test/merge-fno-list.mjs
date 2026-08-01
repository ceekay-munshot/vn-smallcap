#!/usr/bin/env node
// Merge weekly F&O list refresh from Claude.ai routine into the production
// fno-stocks.json. The routine fetches the latest NSE F&O underlyings list
// quarterly (or whenever NSE issues a circular adding/removing names); this
// merger validates the new ticker list and replaces the production list if
// the sanity checks pass.
//
// Inputs:
//   - public/data/weekly-fno-refresh.json    (routine output, schema below)
//   - screener-test/static/fno-stocks.json   (production source-of-truth)
//
// Output:
//   - Mutates `tickers` (+ bookkeeping fields) in fno-stocks.json
//
// Routine output schema:
//   {
//     "generated_at": "ISO",
//     "source": "...",
//     "as_of": "YYYY-MM-DD",
//     "circular_reference": "NSE circular dated YYYY-MM-DD",
//     "tickers": ["RELIANCE", "TCS", ...]
//   }

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REFRESH_PATH = resolve(__dirname, "../public/data/weekly-fno-refresh.json");
const STOCKS_PATH  = resolve(__dirname, "static/fno-stocks.json");

const MIN_LIST_SIZE = 100;      // NSE F&O list is ~200; <100 means extraction failed
const MAX_LIST_SIZE = 500;      // Sanity upper bound — list is never this large
const MAX_CHURN_PCT = 20;       // If routine wants to add/remove >20% of list, reject (likely a bug)
const MAX_STALE_DAYS = 120;     // Routine's `as_of` older than this → reject
const VALID_TICKER  = /^[A-Z0-9&-]{1,15}$/;

if (!existsSync(REFRESH_PATH)) {
  console.log(`No routine output at ${REFRESH_PATH} — nothing to merge.`);
  process.exit(0);
}
const refresh = JSON.parse(readFileSync(REFRESH_PATH, "utf8"));
const stocks  = JSON.parse(readFileSync(STOCKS_PATH, "utf8"));

const newTickers = Array.isArray(refresh.tickers) ? refresh.tickers : [];
const oldTickers = Array.isArray(stocks.tickers)  ? stocks.tickers  : [];
const oldSet     = new Set(oldTickers);

// ---------- Validation ----------
const reasons = [];
if (newTickers.length < MIN_LIST_SIZE)   reasons.push(`size ${newTickers.length} below floor ${MIN_LIST_SIZE}`);
if (newTickers.length > MAX_LIST_SIZE)   reasons.push(`size ${newTickers.length} above ceiling ${MAX_LIST_SIZE}`);

const invalid = newTickers.filter((t) => !VALID_TICKER.test(t));
if (invalid.length) reasons.push(`${invalid.length} invalid-format tickers (first 5: ${invalid.slice(0, 5).join(", ")})`);

const asOfDays = refresh.as_of ? (Date.now() - new Date(refresh.as_of).getTime()) / 86400_000 : Infinity;
if (!Number.isFinite(asOfDays))          reasons.push(`as_of missing or unparseable`);
else if (asOfDays > MAX_STALE_DAYS)      reasons.push(`as_of ${refresh.as_of} >${MAX_STALE_DAYS}d old`);

// Churn guard — if more than MAX_CHURN_PCT% of the old list disappeared OR
// more than MAX_CHURN_PCT% of the new list is brand-new, that's suspicious.
const newSet     = new Set(newTickers);
const added      = newTickers.filter((t) => !oldSet.has(t));
const removed    = oldTickers.filter((t) => !newSet.has(t));
const churn      = Math.max(added.length, removed.length);
const churnPct   = oldTickers.length ? (churn / oldTickers.length) * 100 : 0;
if (churnPct > MAX_CHURN_PCT) reasons.push(`churn ${churnPct.toFixed(1)}% > cap ${MAX_CHURN_PCT}% (added ${added.length}, removed ${removed.length})`);

if (reasons.length) {
  console.log("=== F&O refresh REJECTED ===");
  console.log(`Routine generated_at: ${refresh.generated_at || "(missing)"}`);
  console.log(`Routine as_of:        ${refresh.as_of || "(missing)"}`);
  console.log(`New list size:        ${newTickers.length}`);
  console.log(`Current list size:    ${oldTickers.length}`);
  console.log(`Rejection reasons:`);
  reasons.forEach((r) => console.log(`  ✗ ${r}`));
  console.log(`\nPreserving existing fno-stocks.json (no changes).`);
  process.exit(0);   // not an error — preserved old data
}

// ---------- Accept + write ----------
stocks.tickers              = [...newTickers].sort();
stocks._last_routine_update = refresh.generated_at || new Date().toISOString();
stocks._last_reviewed       = refresh.as_of || stocks._last_reviewed;
if (refresh.circular_reference) {
  stocks._last_routine_source = refresh.circular_reference;
}

writeFileSync(STOCKS_PATH, JSON.stringify(stocks, null, 2) + "\n");

console.log("=== F&O refresh accepted ===");
console.log(`Routine generated_at: ${refresh.generated_at}`);
console.log(`Source:               ${refresh.circular_reference || refresh.source || "(unspecified)"}`);
console.log(`As of:                ${refresh.as_of}`);
console.log(`Tickers: ${oldTickers.length} → ${newTickers.length} (churn ${churnPct.toFixed(1)}%)`);
if (added.length)   console.log(`  + Added (${added.length}):   ${added.slice(0, 15).join(", ")}${added.length > 15 ? ", ..." : ""}`);
if (removed.length) console.log(`  − Removed (${removed.length}): ${removed.slice(0, 15).join(", ")}${removed.length > 15 ? ", ..." : ""}`);
console.log(`\nWrote ${STOCKS_PATH}`);
