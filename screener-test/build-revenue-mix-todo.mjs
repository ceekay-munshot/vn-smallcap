#!/usr/bin/env node
// Build the routine's "to-do list" — public/data/revenue-mix-todo.csv —
// listing companies whose annual report needs per-company macro-honesty
// extraction (geographic revenue, China+1 narrative, rural/govt-capex/
// rate-sensitive/crude/INR/PLI exposure mix).
//
// Priority order (highest first):
//   1. SPIP STRONG BUY companies with no extraction yet
//   2. SPIP BUY companies with no extraction yet
//   3. Any other universe company with no extraction yet
//   4. STALE — existing extraction > 365 days old (annual report cycle)
//
// The daily routine processes TOP N rows (N defaults to 10). After each
// run, that batch falls off the TODO (fresh) and the next priority block
// floats up — new universe additions (e.g., recent IPOs) get inserted at
// the top of their priority tier on the next daily regen.
//
// Composite scoring is run inline so priority reflects the live SPIP
// basket. We mirror app.js's loadTab data enrichments (insider,
// governance, auditor, ATR history) before scoring so the composite
// matches what the dashboard would show.

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import * as composite from "../public/js/composite-scoring.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SCREENER_PATH    = resolve(__dirname, "../public/data/screener-companies.json");
const TECH_PATH        = resolve(__dirname, "../public/data/technicals.json");
const MACRO_PATH       = resolve(__dirname, "../public/data/macro.json");
const INSIDER_PATH     = resolve(__dirname, "../public/data/insider-trades.json");
const GOVERNANCE_PATH  = resolve(__dirname, "../public/data/governance-flags.json");
const AUDITOR_PATH     = resolve(__dirname, "../public/data/auditor-opinions.json");
const ATR_HISTORY_PATH = resolve(__dirname, "../public/data/atr-history.json");
const EXTRACTED_PATH   = resolve(__dirname, "../public/data/company-revenue-mix.json");
const SUPPLEMENT_PATH  = resolve(__dirname, "static/bse-codes-supplement.json");
const OUT_PATH         = resolve(__dirname, "../public/data/revenue-mix-todo.csv");

const STALE_AFTER_DAYS = Number(process.env.REVENUE_MIX_STALE_AFTER_DAYS || 365);

// ---------- Load all data files ----------
const screener   = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));
const techRaw    = JSON.parse(readFileSync(TECH_PATH, "utf8"));
const macroData  = JSON.parse(readFileSync(MACRO_PATH, "utf8"));
const techCompanies = techRaw.companies || techRaw;

let extracted = { companies: {} };
if (existsSync(EXTRACTED_PATH)) {
  try { extracted = JSON.parse(readFileSync(EXTRACTED_PATH, "utf8")); } catch {}
}

let supplement = {};
if (existsSync(SUPPLEMENT_PATH)) {
  try { supplement = JSON.parse(readFileSync(SUPPLEMENT_PATH, "utf8")).codes || {}; }
  catch {}
}

// ---------- Load enrichments (mirrors app.js loadTab for composite) ----------
let insiderByTicker = {}, governanceByTicker = {}, auditorByTicker = {}, atrHistory = {};
try { insiderByTicker    = (JSON.parse(readFileSync(INSIDER_PATH, "utf8")) || {}).companies || {}; } catch {}
try { governanceByTicker = (JSON.parse(readFileSync(GOVERNANCE_PATH, "utf8")) || {}).flagged_companies || {}; } catch {}
try { auditorByTicker    = (JSON.parse(readFileSync(AUDITOR_PATH, "utf8")) || {}).companies || {}; } catch {}
try { atrHistory         =  JSON.parse(readFileSync(ATR_HISTORY_PATH, "utf8")) || {}; } catch {}

// ---------- Enrich rows in-place (matches what app.js does on tab load) ----------
function extractSlug(url) {
  const m = String(url || "").match(/\/company\/([^/]+)/);
  return m ? m[1].toUpperCase() : null;
}

for (const row of screener) {
  const slug = extractSlug(row["Screener URL"]);
  const ticker = slug;
  // Insider
  const ins = ticker ? insiderByTicker[ticker] : null;
  row.insider_loaded = Object.keys(insiderByTicker).length > 0;
  if (ins) {
    row.insider_transactions = ins.transactions || 0;
    row.insider_buy_shares   = ins.buy_shares  || 0;
    row.insider_sell_shares  = ins.sell_shares || 0;
    row.insider_buy_value    = ins.buy_value   || 0;
    row.insider_sell_value   = ins.sell_value  || 0;
    row.insider_net_shares   = ins.net_shares  || 0;
    row.insider_net_value    = ins.net_value   || 0;
    row.insider_last_date    = ins.last_date;
  } else { row.insider_transactions = 0; }
  // Governance
  row.governance_loaded = Object.keys(governanceByTicker).length > 0;
  row.governance_flag = ticker && governanceByTicker[ticker] ? governanceByTicker[ticker] : null;
  // Auditor
  row.auditor_opinions_loaded = Object.keys(auditorByTicker).length > 0;
  const aud = ticker ? auditorByTicker[ticker] : null;
  row.auditor_opinion = aud?.opinion || null;
  row.auditor_opinion_source = aud?.source || null;
}
for (const t of techCompanies) {
  if (t && t.ticker && atrHistory[t.ticker]) t.atr_history = atrHistory[t.ticker];
}

// ---------- Composite scoring ----------
const compositeResults = composite.scoreCompositeBatch(screener, techCompanies, macroData);
const compositeBySlug = new Map();
for (const r of compositeResults) {
  const slug = extractSlug(r.company?.["Screener URL"]);
  if (slug) compositeBySlug.set(slug, { composite: r.composite, rating: r.rating });
}

// ---------- BSE-code resolution ----------
function bseFor(company, slug) {
  if (company.bseCode) return String(company.bseCode);
  if (slug && /^\d+$/.test(slug)) return slug;
  if (slug && supplement[slug]) return String(supplement[slug]);
  return null;
}

function isStale(entry) {
  if (!entry || !entry.extracted_at) return true;
  const ageDays = (Date.now() - new Date(entry.extracted_at).getTime()) / 86400_000;
  return ageDays > STALE_AFTER_DAYS;
}

// ---------- Build priority-tagged TODO ----------
const todo = [];
const noBSE = [];
let skippedFresh = 0;
for (const c of screener) {
  const slug = extractSlug(c["Screener URL"]);
  if (!slug) continue;
  const bse = bseFor(c, slug);
  if (!bse) { noBSE.push(c.Company); continue; }
  const entry = extracted.companies?.[slug];
  const stale = isStale(entry);
  if (entry && !stale) { skippedFresh++; continue; }

  const comp = compositeBySlug.get(slug);
  const rating = comp?.rating || "—";
  let priority;
  if (!entry) {
    // NEW
    if (rating === "STRONG BUY") priority = 1;
    else if (rating === "BUY")    priority = 2;
    else                          priority = 3;
  } else {
    priority = 4; // STALE
  }

  todo.push({
    name: String(c.Company).replace(/,/g, ";").trim(),
    bse,
    slug,
    priority,
    composite: comp?.composite ?? 0,
    rating,
  });
}

// Sort: priority asc (1 first), then composite desc
todo.sort((a, b) => a.priority - b.priority || (b.composite - a.composite));

// ---------- Write CSV ----------
const counts = {
  1: todo.filter((t) => t.priority === 1).length,
  2: todo.filter((t) => t.priority === 2).length,
  3: todo.filter((t) => t.priority === 3).length,
  4: todo.filter((t) => t.priority === 4).length,
};
const header = [
  `# Revenue-mix extraction TODO — auto-generated`,
  `# Generated:        ${new Date().toISOString()}`,
  `# Stale threshold:  ${STALE_AFTER_DAYS} days`,
  `# Universe total:   ${screener.length}`,
  `# Skipped (fresh):  ${skippedFresh}`,
  `# Skipped (no BSE): ${noBSE.length}`,
  `# Priority breakdown:`,
  `#   1 — SPIP STRONG BUY (NEW): ${counts[1]}`,
  `#   2 — SPIP BUY (NEW):        ${counts[2]}`,
  `#   3 — Other NEW:             ${counts[3]}`,
  `#   4 — STALE (>${STALE_AFTER_DAYS}d):           ${counts[4]}`,
  `# Routine processes TOP 10 rows per daily run.`,
  `#`,
  `# Format: Company Name, BSECode`,
].join("\n");
const lines = todo.map((r) => `${r.name}, ${r.bse}`);
writeFileSync(OUT_PATH, header + "\n" + lines.join("\n") + "\n");

console.log(`Universe:           ${screener.length}`);
console.log(`Already fresh:      ${skippedFresh}`);
console.log(`No BSE code yet:    ${noBSE.length}`);
console.log(`Included in to-do:  ${todo.length}`);
console.log(`  Priority 1 (STRONG BUY): ${counts[1]}`);
console.log(`  Priority 2 (BUY):        ${counts[2]}`);
console.log(`  Priority 3 (Other NEW):  ${counts[3]}`);
console.log(`  Priority 4 (STALE):      ${counts[4]}`);
console.log(`\nWrote ${OUT_PATH}`);
