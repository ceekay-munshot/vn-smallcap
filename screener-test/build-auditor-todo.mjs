#!/usr/bin/env node
// Build the routine's "to-do list" — public/data/auditor-todo.csv —
// based on which companies need a fresh auditor opinion.
//
// Inclusion rule per company in the current universe:
//   - We have NO opinion for it yet, OR
//   - Our opinion was fetched > AUDITOR_STALE_AFTER_DAYS ago
//     (auditor opinions tied to annual reports — refresh once a year)
//
// BSE-code resolution priority:
//   1. screener-companies.json `bseCode` field (future — once scrape.mjs
//      enhancement lands)
//   2. URL slug if numeric (BSE-only stocks already use BSE codes as slug)
//   3. static/bse-codes-supplement.json (populated by the backfill workflow)
//
// Output format matches the routine prompt's requirement:
//   Company Name, BSECode
//   (one per line, no header)

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SCREENER_PATH   = resolve(__dirname, "../public/data/screener-companies.json");
const AUDITOR_PATH    = resolve(__dirname, "../public/data/auditor-opinions.json");
const SUPPLEMENT_PATH = resolve(__dirname, "static/bse-codes-supplement.json");
const OUT_PATH        = resolve(__dirname, "../public/data/auditor-todo.csv");

const AUDITOR_STALE_AFTER_DAYS = Number(process.env.AUDITOR_STALE_AFTER_DAYS || 330);
//                                          ≈ 11 months — captures the
//                                          window where the new annual
//                                          report should be out.

const screener = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));
const auditor  = JSON.parse(readFileSync(AUDITOR_PATH, "utf8"));
let supplement = {};
if (existsSync(SUPPLEMENT_PATH)) {
  try { supplement = JSON.parse(readFileSync(SUPPLEMENT_PATH, "utf8")).codes || {}; }
  catch { supplement = {}; }
}

function extractSlug(url) {
  const m = String(url || "").match(/\/company\/([^/]+)/);
  return m ? m[1].toUpperCase() : null;
}

function bseFor(company, slug) {
  if (company.bseCode) return String(company.bseCode);
  if (slug && /^\d+$/.test(slug)) return slug;
  if (slug && supplement[slug]) return String(supplement[slug]);
  return null;
}

function isStale(entry) {
  if (!entry || !entry.fetched_at) return true;
  // If we have a real opinion (not "Could not extract"), only refresh
  // if it's older than the stale threshold.
  const ageDays = (Date.now() - new Date(entry.fetched_at).getTime()) / 86400_000;
  if (entry.opinion && !/could not extract|not\s+(disclosed|provided|available)/i.test(entry.opinion)) {
    return ageDays > AUDITOR_STALE_AFTER_DAYS;
  }
  // No real opinion — always re-try (it was an error or "could not extract"
  // last time, and that may have been a transient).
  return true;
}

let included = 0, skippedNoBSE = 0, skippedFresh = 0;
const todo = [];
const noBSE = [];
for (const c of screener) {
  const slug = extractSlug(c["Screener URL"]);
  if (!slug) continue;
  const bse = bseFor(c, slug);
  if (!bse) {
    skippedNoBSE++;
    noBSE.push(c.Company);
    continue;
  }
  const entry = auditor.companies?.[slug];
  if (!isStale(entry)) { skippedFresh++; continue; }
  // CSV-safe: replace commas in name with semicolons so the parser
  // treats each line as exactly "name,code".
  const safeName = String(c.Company).replace(/,/g, ";").trim();
  todo.push(`${safeName}, ${bse}`);
  included++;
}

const header = [
  `# Auditor opinion to-do list — auto-generated`,
  `# Generated: ${new Date().toISOString()}`,
  `# Stale threshold: ${AUDITOR_STALE_AFTER_DAYS} days`,
  `# Total in universe: ${screener.length}`,
  `# Already fresh (skipped): ${skippedFresh}`,
  `# No BSE code yet (skipped — run bse-codes-backfill workflow): ${skippedNoBSE}`,
  `# To-do this week: ${included}`,
  `#`,
  `# Format: Company Name, BSECode`,
].join("\n");
writeFileSync(OUT_PATH, header + "\n" + todo.join("\n") + "\n");

console.log(`Universe:           ${screener.length}`);
console.log(`Already fresh:      ${skippedFresh}`);
console.log(`No BSE code yet:    ${skippedNoBSE}`);
console.log(`Included in to-do:  ${included}`);
console.log(`\nWrote ${OUT_PATH}`);
if (noBSE.length && noBSE.length <= 20) {
  console.log(`\nCompanies skipped for missing BSE code:`);
  noBSE.forEach((n) => console.log(`  - ${n}`));
} else if (noBSE.length) {
  console.log(`\n${noBSE.length} companies skipped for missing BSE code. First 10:`);
  noBSE.slice(0, 10).forEach((n) => console.log(`  - ${n}`));
  console.log(`(Run the bse-codes-backfill workflow to capture these.)`);
}
