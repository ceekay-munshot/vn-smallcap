#!/usr/bin/env node
// PLI + Renewable auto-refresh. Scrapes a small set of authoritative
// government sources (DPIIT, MNRE, PIB press releases) via Firecrawl's
// LLM extract, pulls out company names mentioned as approved
// beneficiaries, maps them to NSE 500 tickers, and appends new entries
// to screener-test/static/macro-context.json. Purely additive — old
// entries never get removed, so the lists only grow.
//
// Runs quarterly. Cost budget: ~50-80 Firecrawl credits per run.

import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const STATIC_PATH   = resolve(__dirname, "static/macro-context.json");
const SCREENER_PATH = resolve(__dirname, "../public/data/screener-companies.json");
const STATUS_PATH   = resolve(__dirname, "../public/data/pli-renewable-refresh-status.json");

// Diagnostic sidecar — written on every run regardless of whether the
// lists changed. Lets us audit Firecrawl behaviour without access to
// the GHA log viewer.
const DEBUG_PER_SOURCE = [];

const FIRECRAWL_URL = "https://api.firecrawl.dev/v1/scrape";
const API_KEY = process.env.FIRECRAWL_API_KEY;

if (!API_KEY) {
  console.error("FIRECRAWL_API_KEY env var not set — cannot auto-refresh lists.");
  process.exit(0);
}

// Source URLs to scrape. Each is a public landing page for either DPIIT
// or MNRE press releases. Firecrawl renders the page (some are JS-heavy)
// and the LLM pulls Indian listed company names out of the headlines and
// body.
const PLI_SOURCES = [
  {
    url: "https://www.pib.gov.in/PressReleseDetailm.aspx?PRID=2&MenuId=1",
    label: "PIB — recent press releases",
  },
  {
    url: "https://pib.gov.in/indexd.aspx",
    label: "PIB — home",
  },
  {
    url: "https://www.dpiit.gov.in/whats-new",
    label: "DPIIT — what's new",
  },
];

const RENEWABLE_SOURCES = [
  {
    url: "https://mnre.gov.in/whats-new/",
    label: "MNRE — what's new",
  },
  {
    url: "https://mnre.gov.in/press-releases/",
    label: "MNRE — press releases",
  },
];

const COMPANY_EXTRACT_SCHEMA = {
  type: "object",
  properties: {
    company_names: {
      type: "array",
      description: "List of Indian listed company names mentioned on this page as approved PLI beneficiaries, renewable energy auction winners, or capacity allocation recipients. Use the company name as written (e.g. 'Tata Power', 'Reliance Industries'). Only include companies that look like they're listed on NSE / BSE. Skip government entities (NTPC and similar are OK because they're publicly listed). Skip foreign companies.",
      items: { type: "string" },
    },
  },
  required: ["company_names"],
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function firecrawlExtract(url, label, category) {
  console.log(`  Firecrawl LLM extract on: ${url}`);
  const entry = { url, label, category, at: new Date().toISOString() };
  try {
    const r = await fetch(FIRECRAWL_URL, {
      method: "POST",
      headers: { "Authorization": `Bearer ${API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        url,
        formats: ["json"],
        jsonOptions: { schema: COMPANY_EXTRACT_SCHEMA },
        onlyMainContent: false,
      }),
    });
    entry.fc_status = r.status;
    const text = await r.text();
    if (!r.ok) {
      entry.outcome = "fc_http_error";
      entry.error_body = text.slice(0, 300);
      console.log(`    Firecrawl HTTP ${r.status}: ${text.slice(0, 200)}`);
      DEBUG_PER_SOURCE.push(entry);
      return [];
    }
    let j;
    try { j = JSON.parse(text); }
    catch (e) {
      entry.outcome = "non_json_response";
      entry.error_body = text.slice(0, 300);
      DEBUG_PER_SOURCE.push(entry);
      return [];
    }
    const data = j?.data || null;
    const extract = data?.json || data?.llm_extraction || data?.extract || {};
    const names = Array.isArray(extract.company_names) ? extract.company_names : [];
    entry.outcome = "ok";
    entry.upstream_status = data?.metadata?.statusCode || data?.metadata?.["sourceURL-status"] || null;
    entry.candidates_count = names.length;
    entry.candidates_sample = names.slice(0, 20); // first 20 for audit
    console.log(`    Extracted ${names.length} candidate names from ${label}`);
    DEBUG_PER_SOURCE.push(entry);
    return names;
  } catch (err) {
    entry.outcome = "fetch_error";
    entry.error = err.message;
    DEBUG_PER_SOURCE.push(entry);
    console.log(`    Error: ${err.message}`);
    return [];
  }
}

// Build a name → ticker map from the Nifty 500 universe. Normalises
// names aggressively so a press release saying "Tata Power" matches
// "Tata Power Co Ltd." in our universe.
function buildNameMap(screener) {
  const map = new Map();
  for (const c of screener) {
    const fullName = String(c.Company || "");
    const m = String(c["Screener URL"] || "").match(/\/company\/([^/]+)/);
    if (!m) continue;
    const ticker = m[1].toUpperCase();
    // Aliases: full normalised name + each significant word segment of
    // the name. The "Aarti Industries Ltd" entry yields keys
    // "aartiindustries" and "aarti industries".
    const variants = generateNameVariants(fullName);
    for (const v of variants) map.set(v, ticker);
  }
  return map;
}

function normaliseName(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\b(ltd|limited|pvt|private|corporation|corp|inc|company|co|the|of|and|&)\b/g, " ")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function generateNameVariants(name) {
  const out = new Set();
  const n = normaliseName(name);
  if (!n) return out;
  out.add(n);
  out.add(n.replace(/\s+/g, ""));
  // Add first 2 words if the name is long (catches "Aarti Industries Ltd"
  // matching to "Aarti Industries" mention in a press release).
  const parts = n.split(/\s+/);
  if (parts.length >= 2) {
    out.add(parts.slice(0, 2).join(" "));
    out.add(parts.slice(0, 2).join(""));
  }
  if (parts.length >= 3) {
    out.add(parts.slice(0, 3).join(" "));
  }
  // Prefix variant: first 8 chars of the no-space form. Bridges the gap
  // between Screener's name truncation ("Dixon Technolog.") and full
  // names in press releases ("Dixon Technologies"). Long enough to avoid
  // most false positives — "apollotyr" wouldn't match "apolloho".
  const noSpace = n.replace(/\s+/g, "");
  if (noSpace.length >= 8) out.add(noSpace.slice(0, 8));
  return out;
}

function matchToTickers(names, nameMap) {
  const tickers = new Set();
  for (const raw of names) {
    const variants = generateNameVariants(raw);
    for (const v of variants) {
      const t = nameMap.get(v);
      if (t) { tickers.add(t); break; }
    }
  }
  return tickers;
}

async function main() {
  console.log("Loading current macro-context.json and Nifty 500 universe...");
  const stat = JSON.parse(readFileSync(STATIC_PATH, "utf8"));
  const screener = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));
  const nameMap = buildNameMap(screener);
  console.log(`Built ticker map with ${nameMap.size} name variants.`);

  console.log("\n=== PLI sources ===");
  const pliCandidates = new Set();
  for (const src of PLI_SOURCES) {
    const names = await firecrawlExtract(src.url, src.label, "pli");
    const matched = matchToTickers(names, nameMap);
    matched.forEach((t) => pliCandidates.add(t));
    console.log(`    Matched ${matched.size} of ${names.length} to NSE 500 tickers.`);
    await sleep(500);
  }

  console.log("\n=== Renewable sources ===");
  const renewableCandidates = new Set();
  for (const src of RENEWABLE_SOURCES) {
    const names = await firecrawlExtract(src.url, src.label, "renewable");
    const matched = matchToTickers(names, nameMap);
    matched.forEach((t) => renewableCandidates.add(t));
    console.log(`    Matched ${matched.size} of ${names.length} to NSE 500 tickers.`);
    await sleep(500);
  }

  // Diff against existing lists — additive only.
  const currentPLI = new Set((stat.pli_companies || []).map((s) => s.toUpperCase()));
  const currentRenewable = new Set((stat.renewable_companies || []).map((s) => s.toUpperCase()));
  const addedPLI = [...pliCandidates].filter((t) => !currentPLI.has(t)).sort();
  const addedRenewable = [...renewableCandidates].filter((t) => !currentRenewable.has(t)).sort();

  console.log(`\n=== Diff ===`);
  console.log(`PLI: ${currentPLI.size} existing, ${pliCandidates.size} discovered, ${addedPLI.length} new`);
  if (addedPLI.length) console.log(`  New PLI tickers: ${addedPLI.join(", ")}`);
  console.log(`Renewable: ${currentRenewable.size} existing, ${renewableCandidates.size} discovered, ${addedRenewable.length} new`);
  if (addedRenewable.length) console.log(`  New Renewable tickers: ${addedRenewable.join(", ")}`);

  // Always write the diagnostic sidecar — committed every run so we can
  // audit Firecrawl behaviour by reading the JSON instead of needing
  // GHA log access.
  const status = {
    generated_at: new Date().toISOString(),
    pli: {
      existing: currentPLI.size,
      discovered: pliCandidates.size,
      added: addedPLI,
    },
    renewable: {
      existing: currentRenewable.size,
      discovered: renewableCandidates.size,
      added: addedRenewable,
    },
    sources: DEBUG_PER_SOURCE,
  };
  writeFileSync(STATUS_PATH, JSON.stringify(status, null, 2) + "\n");
  console.log(`\nWrote status sidecar to ${STATUS_PATH}`);

  if (!addedPLI.length && !addedRenewable.length) {
    console.log("No new beneficiaries to add to macro-context.json — exiting cleanly.");
    return;
  }

  // Append new entries, preserving existing order.
  if (addedPLI.length) stat.pli_companies = [...(stat.pli_companies || []), ...addedPLI];
  if (addedRenewable.length) stat.renewable_companies = [...(stat.renewable_companies || []), ...addedRenewable];
  stat._last_auto_refresh = new Date().toISOString();
  stat._last_auto_refresh_summary = {
    added_pli: addedPLI,
    added_renewable: addedRenewable,
  };

  writeFileSync(STATIC_PATH, JSON.stringify(stat, null, 2) + "\n");
  console.log(`Wrote updated macro-context.json with ${addedPLI.length} new PLI + ${addedRenewable.length} new Renewable entries.`);
}

main().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  process.exit(0);
});
