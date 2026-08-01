#!/usr/bin/env node
// Auditor opinions scraper. Hits the Muns "Auditor Opinion" agent
// (devde.muns.io/agents/run) once per company in the Screener-derived
// universe, extracts the auditor opinion classification from the agent's
// <conclusion> block, and writes a per-ticker map.
//
// The Fundamentals "Auditor Remarks" rule reads this file to score the
// company's most recent annual audit opinion.
//
// Caching: entries younger than 30 days are kept as-is. Errors are
// always re-tried next run. "Not disclosed" / "Not provided" responses
// are cached (so the API isn't burned daily on companies the agent can't
// resolve) but render as N/A on the dashboard, and refresh after 30 days.
//
// Concurrency: 50 calls in flight at a time.

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SCREENER_PATH = resolve(__dirname, "../public/data/screener-companies.json");
const OUT_PATH      = resolve(__dirname, "../public/data/auditor-opinions.json");

const API_URL           = "https://devde.muns.io/agents/run";
const USER_INDEX        = 124;
const AGENT_LIBRARY_ID  = "0789b5f8-48f2-49db-9490-399dcdc294ba";
const STOCK_COUNTRY     = "IN";
const TIMEZONE          = "Asia/Kolkata";
const CACHE_MAX_AGE_DAYS = 30;
const BATCH_SIZE        = 50;
const PER_CALL_TIMEOUT_MS = 360_000;

const TOKEN = process.env.MUNS_ACCESS_TOKEN;
const FORCE_REFRESH = String(process.env.FORCE_REFRESH || "").toLowerCase() === "true";
const MAX_COMPANIES = Number(process.env.MAX_COMPANIES || 0);

if (!TOKEN) {
  console.error("MUNS_ACCESS_TOKEN env var not set — aborting.");
  process.exit(1);
}

if (!existsSync(SCREENER_PATH)) {
  console.error(`Screener companies file missing: ${SCREENER_PATH}`);
  process.exit(1);
}

const companies = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));

let cache = { generated_at: null, companies: {} };
if (existsSync(OUT_PATH)) {
  try {
    cache = JSON.parse(readFileSync(OUT_PATH, "utf8"));
    if (!cache.companies) cache.companies = {};
  } catch (e) {
    console.warn("Existing auditor-opinions.json unparseable — starting fresh.", e.message);
    cache = { generated_at: null, companies: {} };
  }
}
const cacheMap = cache.companies;

const cacheMaxAgeMs = CACHE_MAX_AGE_DAYS * 24 * 3600 * 1000;
const now = Date.now();

function tickerOf(row) {
  const m = String(row["Screener URL"] || "").match(/\/company\/([^/]+)/);
  return m ? m[1].toUpperCase() : null;
}

function needsRefresh(ticker) {
  if (FORCE_REFRESH) return true;
  const entry = cacheMap[ticker];
  if (!entry) return true;
  if (entry.error) return true;
  if (!entry.fetched_at) return true;
  const age = now - new Date(entry.fetched_at).getTime();
  return age > cacheMaxAgeMs;
}

// Find every company that needs a fresh API call.
const allWithTicker = companies
  .map((row) => ({ ticker: tickerOf(row), name: row.Company }))
  .filter((x) => x.ticker && x.name);

let toFetch = allWithTicker.filter((x) => needsRefresh(x.ticker));
if (MAX_COMPANIES > 0) toFetch = toFetch.slice(0, MAX_COMPANIES);

const skippedCount = allWithTicker.length - toFetch.length;
console.log(`Universe: ${allWithTicker.length} companies with tickers.`);
console.log(`Cached (≤${CACHE_MAX_AGE_DAYS}d, no error): ${skippedCount} — skipping.`);
console.log(`Calling API for: ${toFetch.length}${FORCE_REFRESH ? " (force-refresh)" : ""}.`);

// Drop cache entries for tickers no longer in the universe (delisted /
// renamed) so the file doesn't grow unboundedly across CSV refreshes.
const validTickers = new Set(allWithTicker.map((x) => x.ticker));
for (const t of Object.keys(cacheMap)) {
  if (!validTickers.has(t)) delete cacheMap[t];
}

const today = new Date().toISOString().slice(0, 10);

// Pull the <conclusion>...</conclusion> block out of the agent response and
// extract the middle column ("Auditor's Opinion") from its markdown table.
// Shape we expect (from a real sample):
//   <conclusion>| Company Name | Auditor's Opinion | Source |
//   |---|---|---|
//   | Not provided | Unqualified Opinion | Most recent annual report ... |</conclusion>
function parseAuditorOpinion(responseText) {
  const m = String(responseText).match(/<conclusion>([\s\S]*?)<\/conclusion>/i);
  if (!m) return { opinion: null, source: null, conclusion_block: null };
  const block = m[1].trim();
  const rows = block.split("\n").map((l) => l.trim()).filter((l) => l.startsWith("|"));
  // Drop the separator row (|---|---|...|)
  const cleaned = rows.filter((l) => !/^\|[\s\-:|]+\|?$/.test(l));
  if (cleaned.length < 2) return { opinion: null, source: null, conclusion_block: block };

  const splitCells = (line) =>
    line.split("|").slice(1, -1).map((s) => s.trim());
  const header = splitCells(cleaned[0]);
  const data   = splitCells(cleaned[1]);

  let opIdx = header.findIndex((h) => /auditor/i.test(h) && /opinion/i.test(h));
  if (opIdx < 0) opIdx = 1;
  let srcIdx = header.findIndex((h) => /source/i.test(h));
  if (srcIdx < 0) srcIdx = 2;

  const opinion = (data[opIdx] || "").trim() || null;
  const source  = (data[srcIdx] || "").trim() || null;
  return { opinion, source, conclusion_block: block };
}

async function fetchOne({ ticker, name }) {
  const body = {
    user_index: USER_INDEX,
    agent_library_id: AGENT_LIBRARY_ID,
    metadata: {
      stock_ticker: ticker,
      stock_company_name: name,
      context_company_name: name,
      stock_country: STOCK_COUNTRY,
      to_date: today,
      timezone: TIMEZONE,
    },
  };
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), PER_CALL_TIMEOUT_MS);
  try {
    const res = await fetch(API_URL, {
      method: "POST",
      headers: {
        "accept": "*/*",
        "Authorization": `Bearer ${TOKEN}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    const text = await res.text();
    if (!res.ok) {
      return { ticker, name, error: `HTTP ${res.status}`, http_status: res.status,
               error_snippet: text.slice(0, 200), fetched_at: new Date().toISOString() };
    }
    const parsed = parseAuditorOpinion(text);
    return {
      ticker, name,
      opinion: parsed.opinion,
      source: parsed.source,
      conclusion_block: parsed.conclusion_block,
      fetched_at: new Date().toISOString(),
    };
  } catch (err) {
    return { ticker, name, error: err.name === "AbortError" ? "timeout" : err.message,
             fetched_at: new Date().toISOString() };
  } finally {
    clearTimeout(timer);
  }
}

let processed = 0;
for (let i = 0; i < toFetch.length; i += BATCH_SIZE) {
  const batch = toFetch.slice(i, i + BATCH_SIZE);
  const batchNum = Math.floor(i / BATCH_SIZE) + 1;
  const totalBatches = Math.ceil(toFetch.length / BATCH_SIZE);
  console.log(`Batch ${batchNum}/${totalBatches}: ${batch.length} companies (${batch.slice(0, 5).map((b) => b.ticker).join(", ")}${batch.length > 5 ? ", ..." : ""})`);
  const results = await Promise.all(batch.map(fetchOne));
  for (const r of results) {
    cacheMap[r.ticker] = {
      name: r.name,
      opinion: r.opinion || null,
      source: r.source || null,
      conclusion_block: r.conclusion_block || null,
      error: r.error || null,
      http_status: r.http_status || null,
      error_snippet: r.error_snippet || null,
      fetched_at: r.fetched_at,
    };
    processed++;
    if (r.error) {
      console.log(`  ✗ ${r.ticker}: ${r.error}`);
    } else if (!r.opinion) {
      console.log(`  ~ ${r.ticker}: no opinion in response`);
    }
  }
  console.log(`  → ${processed}/${toFetch.length} processed.`);
}

// Stats for the output header.
const entries = Object.values(cacheMap);
const liveCount = entries.filter((e) => !e.error && e.opinion).length;
const ndCount   = entries.filter((e) => !e.error && !e.opinion).length;
const errCount  = entries.filter((e) => e.error).length;

const out = {
  generated_at: new Date().toISOString(),
  source: {
    url: API_URL,
    agent_library_id: AGENT_LIBRARY_ID,
    user_index: USER_INDEX,
    label: "Muns Auditor Opinion agent",
  },
  cache_max_age_days: CACHE_MAX_AGE_DAYS,
  total_universe: allWithTicker.length,
  total_cached: Object.keys(cacheMap).length,
  total_live_opinion: liveCount,
  total_not_disclosed: ndCount,
  total_errors: errCount,
  companies: cacheMap,
};

writeFileSync(OUT_PATH, JSON.stringify(out, null, 2));
console.log(`\nWrote ${OUT_PATH}`);
console.log(`Live opinion: ${liveCount}, not disclosed: ${ndCount}, errors: ${errCount}, total cached: ${Object.keys(cacheMap).length}`);
