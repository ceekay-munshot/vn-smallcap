#!/usr/bin/env node
// Merge Claude.ai routine's auditor extraction (weekly-auditor-test.json)
// into the dashboard's primary auditor-opinions.json.
//
// Routine output shape (per company):
//   { company_name, bse_code, auditor_firm_name, opinion_type,
//     opinion_date, report_year, emphasis_of_matter,
//     key_concerns_flagged_by_auditor, confidence_in_extraction }
//
// Dashboard primary shape (keyed by URL-slug = NSE ticker or BSE code):
//   { name, opinion, source, conclusion_block, error, fetched_at,
//     [new rich fields preserved as auditor_*] }
//
// Rule (`ruleAuditorOpinion` in scoring.js) reads `auditor_opinion` +
// `auditor_opinion_source` and classifies via regex:
//   - "clean"/"unqualified" → pass (2 pts)
//   - "qualified" or "emphasis-of-matter" (not "unqualified") → partial (1 pt)
//   - "adverse" or "disclaimer" → hard_fail
//   - empty/unknown → N/A
//
// So we write the `opinion` field as a string that matches what the
// rule expects, AND keep all the rich routine fields as additional
// properties for the drill-down (no separate parsing path needed).

import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROUTINE_PATH  = resolve(__dirname, "../public/data/weekly-auditor-test.json");
const SCREENER_PATH = resolve(__dirname, "../public/data/screener-companies.json");
const OUT_PATH      = resolve(__dirname, "../public/data/auditor-opinions.json");

// ---------- Helpers ----------
function extractSlug(url) {
  const m = String(url || "").match(/\/company\/([^/]+)/);
  return m ? m[1].toUpperCase() : null;
}

function normaliseName(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\b(ltd|limited|pvt|private|corporation|corp|inc|company|co|the|of|and|&)\b/g, " ")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Map the routine's opinion_type → an opinion string the existing rule
// classifies correctly. Augment with EoM where applicable so the rule's
// emphasis-of-matter branch triggers.
function buildOpinionString(routineRow) {
  const type = String(routineRow.opinion_type || "").trim();
  const eom = String(routineRow.emphasis_of_matter || "").trim();
  const hasEoM = eom && eom !== "None" && eom !== "N/A" && eom !== "Could not extract";
  if (!type || /^could not extract$/i.test(type)) return null;
  if (/^clean$/i.test(type)) {
    return hasEoM
      ? "Clean (Unmodified) with Emphasis of Matter"   // rule → partial (1 pt)
      : "Clean (Unmodified Opinion)";                  // rule → pass (2 pts)
  }
  if (/^qualified$/i.test(type)) return "Qualified Opinion";       // rule → partial
  if (/^adverse$/i.test(type)) return "Adverse Opinion";           // rule → hard_fail
  if (/disclaimer/i.test(type)) return "Disclaimer of Opinion";    // rule → hard_fail
  return type; // pass through unknown values verbatim
}

// ---------- Main ----------
console.log("Loading routine output, screener universe, and existing opinions...");
const routine  = JSON.parse(readFileSync(ROUTINE_PATH, "utf8"));
const screener = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));
const existing = JSON.parse(readFileSync(OUT_PATH, "utf8"));

// Build a lookup from BOTH normalised name AND BSE code → URL slug.
// The dashboard's auditor-opinions.json is keyed by URL slug
// (NSE ticker for NSE-listed, BSE code for BSE-only). The routine
// output carries company_name + bse_code, so we need to translate.
const slugByName = new Map();
const slugByBSE  = new Map();
for (const c of screener) {
  const slug = extractSlug(c["Screener URL"]);
  if (!slug) continue;
  const nName = normaliseName(c.Company);
  if (nName) slugByName.set(nName, slug);
  // If the slug is numeric, that IS the BSE code — index it too.
  if (/^\d+$/.test(slug)) slugByBSE.set(slug, slug);
}
console.log(`Built lookup: ${slugByName.size} by name, ${slugByBSE.size} by BSE code.`);

let updated = 0, kept = 0, unmatched = 0;
const unmatchedList = [];
const stats = {
  pass: 0, partial: 0, hard_fail: 0, na: 0,
  by_type: {},
};

for (const r of routine) {
  const nName = normaliseName(r.company_name);
  let slug = slugByName.get(nName) || slugByBSE.get(String(r.bse_code || "").trim());
  if (!slug) {
    // Last resort: try fuzzy prefix match on first 2–3 normalised words.
    const words = nName.split(/\s+/);
    for (const L of [3, 2]) {
      if (words.length >= L) {
        const prefix = words.slice(0, L).join(" ");
        const hits = [...slugByName.entries()].filter(([k]) => k.startsWith(prefix));
        if (hits.length === 1) { slug = hits[0][1]; break; }
      }
    }
  }
  if (!slug) {
    unmatched++;
    unmatchedList.push({ name: r.company_name, bse: r.bse_code });
    continue;
  }
  const opinionStr = buildOpinionString(r);
  if (opinionStr == null) {
    // "Could not extract" — preserve any prior data instead of overwriting.
    kept++;
    continue;
  }

  // Track classification this routine entry will produce in the rule.
  const opL = opinionStr.toLowerCase();
  if (/\b(adverse|disclaimer)\b/.test(opL)) stats.hard_fail++;
  else if (/\bqualified\b|emphasis[- ]of[- ]matter/.test(opL) && !/un\s*-?\s*qualified/.test(opL)) stats.partial++;
  else if (/\bclean\b|\bunqualified\b/.test(opL)) stats.pass++;
  else stats.na++;
  stats.by_type[r.opinion_type] = (stats.by_type[r.opinion_type] || 0) + 1;

  // Build the merged entry. Keep all routine-extracted rich fields as
  // auditor_* prefixed props so the dashboard drill-down can surface them.
  const existingEntry = existing.companies?.[slug] || {};
  existing.companies[slug] = {
    ...existingEntry,
    name: existingEntry.name || r.company_name,
    opinion: opinionStr,
    // Kept short — the dashboard reads `auditor_confidence` separately for
    // internal QA, but the visible source label stays plain "BSE annual
    // report" so the drill-down doesn't expose pipeline mechanics.
    source: "BSE annual report",
    conclusion_block: r.key_concerns_flagged_by_auditor || null,
    error: null,
    http_status: 200,
    error_snippet: null,
    fetched_at: new Date().toISOString(),
    // Rich routine fields preserved verbatim for the drill-down
    auditor_firm: r.auditor_firm_name || null,
    auditor_opinion_date: r.opinion_date || null,
    auditor_report_year: r.report_year || null,
    auditor_emphasis_of_matter: r.emphasis_of_matter || null,
    auditor_key_concerns: r.key_concerns_flagged_by_auditor || null,
    auditor_confidence: r.confidence_in_extraction || null,
    _origin: "routine",
  };
  updated++;
}

// Refresh top-level counters.
const allEntries = Object.values(existing.companies || {});
existing.generated_at = new Date().toISOString();
existing.total_universe = allEntries.length;
existing.total_live_opinion = allEntries.filter((e) => e.opinion && !/not\s+(disclosed|provided|available)/i.test(e.opinion)).length;
existing.total_not_disclosed = allEntries.filter((e) => e.opinion && /not\s+(disclosed|provided|available)/i.test(e.opinion)).length;
existing.total_errors = allEntries.filter((e) => e.error).length;

writeFileSync(OUT_PATH, JSON.stringify(existing, null, 2) + "\n");

console.log("");
console.log("=== Merge summary ===");
console.log(`Routine entries processed:    ${routine.length}`);
console.log(`Updated existing entries:     ${updated}`);
console.log(`Skipped (Could not extract):  ${kept}`);
console.log(`Unmatched to screener:        ${unmatched}`);
console.log("");
console.log("Routine opinion_type distribution:", stats.by_type);
console.log("");
console.log("Rule-outcome breakdown (how the dashboard will score them):");
console.log(`  pass (Clean):              ${stats.pass}`);
console.log(`  partial (Qualified / EoM): ${stats.partial}`);
console.log(`  hard_fail (Adverse/Discl): ${stats.hard_fail}`);
console.log(`  N/A:                       ${stats.na}`);
if (unmatchedList.length) {
  console.log("");
  console.log("Unmatched routine entries (no slug found in screener):");
  unmatchedList.slice(0, 10).forEach((u) => console.log(`  ${u.name}  (BSE: ${u.bse})`));
  if (unmatchedList.length > 10) console.log(`  ... and ${unmatchedList.length - 10} more`);
}
console.log("");
console.log(`Wrote ${OUT_PATH}`);
