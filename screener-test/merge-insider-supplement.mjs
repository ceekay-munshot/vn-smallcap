#!/usr/bin/env node
// Merge daily insider-refresh from Claude.ai routine into a SUPPLEMENT file
// that the dashboard reads alongside insider-trades.json.
//
// Why a supplement (instead of editing insider-trades.json directly):
// the daily NSE PIT + Firecrawl scrape REGENERATES insider-trades.json from
// scratch every morning based on the NSE API's 180-day window. Anything we
// write into that file would be wiped on the next daily run. The supplement
// (public/data/insider-recent-supplement.json) lives alongside, is owned by
// this routine, and the dashboard merges both files at load time.
//
// Routine output schema (public/data/daily-insider-refresh.json):
//   {
//     "generated_at": "ISO",
//     "source": "BSE — Reg 7(2) latest disclosures",
//     "as_of": "YYYY-MM-DD",
//     "filings": [
//       {
//         "ticker": "OLAELEC",       // NSE symbol (uppercase)
//         "company_name": "Ola Electric",
//         "bse_code": "544225",      // string
//         "date": "2026-06-02",      // ISO date of disclosure
//         "type": "Sell",            // "Buy" | "Sell" | "Pledge" | "Pledge Release"
//         "shares": 25000,
//         "value": 3500000,          // rupees
//         "person": "ABC",           // who traded
//         "category": "Promoter",    // "Promoter" | "KMP" | "Designated Person"
//         "filing_url": "https://www.bseindia.com/..."
//       },
//       ...
//     ]
//   }
//
// Pledge / Pledge Release filings are kept in `seen_filings` for dedupe but
// excluded from buy/sell aggregates (matches the scoring.js rule's behaviour).

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REFRESH_PATH    = resolve(__dirname, "../public/data/daily-insider-refresh.json");
const SUPPLEMENT_PATH = resolve(__dirname, "../public/data/insider-recent-supplement.json");

// Sanity bounds — a single filing outside these is rejected (likely bad parse).
const BOUNDS = {
  shares: { min: 1,  max: 1e10 },     // up to 10 billion shares
  value:  { min: 1,  max: 1e13 },     // up to ₹10 lakh crore (some PSU disposals are massive)
};
const MAX_STALE_DAYS = 30;            // filings older than 30 days are dropped (NSE PIT covers history)
const MAX_FUTURE_DAYS = 2;            // tolerate timezone slop

const isAgg = (t) => /^(Buy|Sell)$/i.test(String(t || "").trim());
const isPledge = (t) => /pledge/i.test(String(t || ""));

// Stable fingerprint for dedupe across runs: ticker + date + type + shares + person.
function fingerprint(f) {
  return [
    String(f.ticker || "").toUpperCase().trim(),
    String(f.date || "").slice(0, 10),
    String(f.type || "").trim(),
    Number(f.shares || 0),
    String(f.person || "").trim().toLowerCase(),
  ].join("|");
}

function inRange(v, b) {
  return typeof v === "number" && Number.isFinite(v) && v >= b.min && v <= b.max;
}

if (!existsSync(REFRESH_PATH)) {
  console.log(`No routine output at ${REFRESH_PATH} — nothing to merge.`);
  process.exit(0);
}
const refresh = JSON.parse(readFileSync(REFRESH_PATH, "utf8"));
const incoming = Array.isArray(refresh.filings) ? refresh.filings : [];

// Load existing supplement (or seed empty).
let supplement;
if (existsSync(SUPPLEMENT_PATH)) {
  try { supplement = JSON.parse(readFileSync(SUPPLEMENT_PATH, "utf8")); }
  catch { supplement = null; }
}
if (!supplement || typeof supplement !== "object") {
  supplement = { generated_at: null, source: "BSE Reg 7(2) — daily routine", companies: {} };
}
if (!supplement.companies) supplement.companies = {};

const summary = { accepted: 0, rejected: [], dedup: 0, by_type: {} };

const today = new Date();
for (const f of incoming) {
  // -- Validation gates --
  const tk = String(f.ticker || "").toUpperCase().trim();
  if (!tk) { summary.rejected.push(`missing ticker: ${JSON.stringify(f).slice(0, 80)}`); continue; }

  const dt = new Date(f.date);
  if (!Number.isFinite(dt.getTime())) { summary.rejected.push(`${tk}: bad date ${f.date}`); continue; }
  const ageDays = (today.getTime() - dt.getTime()) / 86400_000;
  if (ageDays > MAX_STALE_DAYS) { summary.rejected.push(`${tk}: filing ${ageDays.toFixed(0)}d old — covered by daily NSE scrape`); continue; }
  if (ageDays < -MAX_FUTURE_DAYS) { summary.rejected.push(`${tk}: filing dated in future (${f.date})`); continue; }

  if (!inRange(f.shares, BOUNDS.shares)) { summary.rejected.push(`${tk}: shares ${f.shares} out of range`); continue; }
  if (f.value != null && !inRange(f.value, BOUNDS.value)) {
    // Value is optional (some filings disclose only shares); only reject if explicit and bad.
    summary.rejected.push(`${tk}: value ${f.value} out of range`); continue;
  }

  // -- Per-company entry --
  if (!supplement.companies[tk]) {
    supplement.companies[tk] = {
      buy_shares: 0, sell_shares: 0,
      buy_value: 0,  sell_value: 0,
      transactions: 0, pledges_excluded: 0,
      last_date: null,
      seen_filings: [],
    };
  }
  const entry = supplement.companies[tk];

  // Dedupe — same filing on a re-run is silently skipped.
  const fp = fingerprint(f);
  if (entry.seen_filings.includes(fp)) { summary.dedup++; continue; }
  entry.seen_filings.push(fp);

  summary.by_type[f.type] = (summary.by_type[f.type] || 0) + 1;

  if (isPledge(f.type)) {
    entry.pledges_excluded += 1;
  } else if (isAgg(f.type)) {
    const isBuy = /^buy$/i.test(f.type);
    if (isBuy) {
      entry.buy_shares += Number(f.shares) || 0;
      entry.buy_value  += Number(f.value)  || 0;
    } else {
      entry.sell_shares += Number(f.shares) || 0;
      entry.sell_value  += Number(f.value)  || 0;
    }
    entry.transactions += 1;
  }

  // Track most recent disclosure date.
  if (!entry.last_date || new Date(f.date) > new Date(entry.last_date)) {
    entry.last_date = f.date;
  }
  summary.accepted++;
}

// Recompute net columns + top-level counters.
for (const tk of Object.keys(supplement.companies)) {
  const e = supplement.companies[tk];
  e.net_shares = (e.buy_shares || 0) - (e.sell_shares || 0);
  e.net_value  = (e.buy_value  || 0) - (e.sell_value  || 0);
}

supplement.generated_at = refresh.generated_at || new Date().toISOString();
supplement.as_of        = refresh.as_of || new Date().toISOString().slice(0, 10);
supplement.source       = "BSE Reg 7(2) — daily routine supplement to NSE PIT";
supplement.total_companies_supplemented = Object.keys(supplement.companies).length;
supplement.routine_filings_seen = Object.values(supplement.companies)
  .reduce((acc, e) => acc + (e.seen_filings?.length || 0), 0);

writeFileSync(SUPPLEMENT_PATH, JSON.stringify(supplement, null, 2) + "\n");

console.log("=== Insider supplement merge ===");
console.log(`Routine generated_at: ${refresh.generated_at || "(missing)"}`);
console.log(`Incoming filings:     ${incoming.length}`);
console.log(`Accepted:             ${summary.accepted}`);
console.log(`Duplicates skipped:   ${summary.dedup}`);
console.log(`Rejected:             ${summary.rejected.length}`);
if (summary.rejected.length) summary.rejected.slice(0, 10).forEach((r) => console.log(`  ✗ ${r}`));
console.log(`By type:              ${JSON.stringify(summary.by_type)}`);
console.log(`\nSupplement covers ${supplement.total_companies_supplemented} companies with ${supplement.routine_filings_seen} cumulative filings.`);
console.log(`Wrote ${SUPPLEMENT_PATH}`);
