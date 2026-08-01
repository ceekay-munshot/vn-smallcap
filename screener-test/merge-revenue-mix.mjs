#!/usr/bin/env node
// Merge daily revenue-mix refresh (from Claude.ai routine) into the
// production per-company file company-revenue-mix.json.
//
// Schema per company (output by routine in daily-revenue-mix-refresh.json):
//   {
//     "PIIND": {
//       "report_year": "FY2024-25",
//       "geography": {
//         "domestic_pct": 12,
//         "export_pct": 88,
//         "export_us_pct": 35,                     // optional sub-breakdown
//         "export_europe_pct": 28,
//         "export_asia_ex_china_pct": 18
//       },
//       "china_plus_one": {
//         "narrative_present": true,
//         "strength": "strong",                     // strong | moderate | weak | none
//         "evidence": "..."
//       },
//       "rural":          { "revenue_pct": null, "evidence": null },
//       "govt_capex":     { "revenue_pct": null, "evidence": null },
//       "rate_sensitive": { "floating_rate_liability_pct": null, "evidence": null },
//       "crude_exposure": { "raw_material_pct": null, "evidence": null },
//       "inr_exposure":   { "usd_revenue_pct": 88, "hedging_policy": "...", "evidence": "..." },
//       "pli_scheme":     { "is_beneficiary": false, "revenue_pct": null, "evidence": null },
//       "confidence": "High"
//     }
//   }
//
// Each per-company entry is independently validated — bad ones are
// rejected and the previous extraction is preserved.

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REFRESH_PATH = resolve(__dirname, "../public/data/daily-revenue-mix-refresh.json");
const PROD_PATH    = resolve(__dirname, "../public/data/company-revenue-mix.json");

const ALLOWED_STRENGTH   = ["strong", "moderate", "weak", "none"];
const ALLOWED_CONFIDENCE = ["High", "Medium", "Low"];
const MAX_REPORT_AGE_YEARS = 2;
const PCT_TOLERANCE = 5;   // domestic_pct + export_pct should sum to ~100 within ±5pp

function inPctRange(v) {
  return typeof v === "number" && Number.isFinite(v) && v >= 0 && v <= 100;
}

function validate(slug, entry) {
  const errs = [];

  // report_year
  const yMatch = String(entry.report_year || "").match(/FY(\d{4})-(\d{2}|\d{4})/);
  if (!yMatch) {
    errs.push("report_year missing or unparseable");
  } else {
    const endYear = Number(yMatch[2].length === 2 ? `20${yMatch[2]}` : yMatch[2]);
    const ageYears = new Date().getFullYear() - endYear;
    if (ageYears > MAX_REPORT_AGE_YEARS) errs.push(`report_year FY ending ${endYear} is >${MAX_REPORT_AGE_YEARS}y old`);
  }

  // confidence
  if (!ALLOWED_CONFIDENCE.includes(entry.confidence)) errs.push(`confidence "${entry.confidence}" not in ${ALLOWED_CONFIDENCE.join("/")}`);

  // geography
  const g = entry.geography || {};
  if (g.domestic_pct != null && !inPctRange(g.domestic_pct)) errs.push(`geography.domestic_pct ${g.domestic_pct} out of 0-100`);
  if (g.export_pct   != null && !inPctRange(g.export_pct))   errs.push(`geography.export_pct ${g.export_pct} out of 0-100`);
  if (g.domestic_pct != null && g.export_pct != null) {
    const sum = g.domestic_pct + g.export_pct;
    if (Math.abs(sum - 100) > PCT_TOLERANCE) errs.push(`geography domestic+export = ${sum}% (expected ≈100%)`);
  }

  // china_plus_one
  const cpo = entry.china_plus_one || {};
  if (cpo.strength != null && !ALLOWED_STRENGTH.includes(cpo.strength)) {
    errs.push(`china_plus_one.strength "${cpo.strength}" not in ${ALLOWED_STRENGTH.join("/")}`);
  }
  if (cpo.narrative_present != null && typeof cpo.narrative_present !== "boolean") {
    errs.push(`china_plus_one.narrative_present not boolean`);
  }

  // Optional pct fields across other signals
  const subPct = [
    ["rural",          "revenue_pct"],
    ["govt_capex",     "revenue_pct"],
    ["rate_sensitive", "floating_rate_liability_pct"],
    ["crude_exposure", "raw_material_pct"],
    ["inr_exposure",   "usd_revenue_pct"],
    ["pli_scheme",     "revenue_pct"],
  ];
  for (const [block, field] of subPct) {
    const v = entry?.[block]?.[field];
    if (v != null && !inPctRange(v)) errs.push(`${block}.${field} ${v} out of 0-100`);
  }
  if (entry.pli_scheme?.is_beneficiary != null && typeof entry.pli_scheme.is_beneficiary !== "boolean") {
    errs.push(`pli_scheme.is_beneficiary not boolean`);
  }

  return errs;
}

if (!existsSync(REFRESH_PATH)) {
  console.log(`No routine output at ${REFRESH_PATH} — nothing to merge.`);
  process.exit(0);
}

const refresh = JSON.parse(readFileSync(REFRESH_PATH, "utf8"));
let prod;
if (existsSync(PROD_PATH)) {
  try { prod = JSON.parse(readFileSync(PROD_PATH, "utf8")); }
  catch { prod = null; }
}
if (!prod || typeof prod !== "object") {
  prod = { generated_at: null, source: "Per-company annual-report extraction (routine)", companies: {} };
}
if (!prod.companies) prod.companies = {};

const refreshCompanies = refresh.companies || {};
const summary = { accepted: [], rejected: [], skipped_unchanged: [] };

for (const [slug, entry] of Object.entries(refreshCompanies)) {
  const SLUG = slug.toUpperCase();
  const errs = validate(SLUG, entry);
  if (errs.length) {
    summary.rejected.push({ slug: SLUG, errors: errs });
    continue;
  }
  // Stamp the extraction timestamp on the entry itself so we can detect staleness later.
  const merged = {
    ...entry,
    extracted_at: refresh.generated_at || new Date().toISOString(),
  };
  prod.companies[SLUG] = merged;
  summary.accepted.push(SLUG);
}

prod.generated_at        = refresh.generated_at || new Date().toISOString();
prod.source              = "Per-company annual-report extraction (Claude.ai routine, daily batch of ~10)";
prod.total_extracted     = Object.keys(prod.companies).length;
prod.last_batch_size     = summary.accepted.length;
prod.last_batch_rejected = summary.rejected.length;

writeFileSync(PROD_PATH, JSON.stringify(prod, null, 2) + "\n");

console.log("=== Revenue-mix merge ===");
console.log(`Routine generated_at: ${refresh.generated_at || "(missing)"}`);
console.log(`Incoming entries:     ${Object.keys(refreshCompanies).length}`);
console.log(`Accepted:             ${summary.accepted.length} → ${summary.accepted.slice(0, 8).join(", ")}${summary.accepted.length > 8 ? ", ..." : ""}`);
console.log(`Rejected:             ${summary.rejected.length}`);
if (summary.rejected.length) {
  for (const r of summary.rejected.slice(0, 8)) {
    console.log(`  ✗ ${r.slug}: ${r.errors.join("; ")}`);
  }
}
console.log(`\nProduction file now covers ${prod.total_extracted} companies.`);
console.log(`Wrote ${PROD_PATH}`);
