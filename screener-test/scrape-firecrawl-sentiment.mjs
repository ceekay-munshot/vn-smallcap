#!/usr/bin/env node
// Sentiment extras scraper via Firecrawl. NSE blocks GitHub Actions runner
// IPs directly, but Firecrawl's residential-proxy + headless-browser stack
// bypasses that. This file handles two NSE-blocked sentiment rules:
//
//   1. Put-Call Ratio  — fetched from NSE's NIFTY option-chain JSON API,
//      with a Trendlyne HTML fallback if NSE refuses.
//   2. Impact Cost     — probed against NSE's monthly archive CSVs
//      (URL pattern shifts month to month, we try ~8 variants × 4 months).
//
// Output: public/data/sentiment-extras.json
//   { pcr: { value, source, fetched_at },
//     impact_cost: { period, source, companies: { TICKER: pct } } }
//
// If Firecrawl returns nothing usable for either, we write the slot as null
// and the rules stay deferred. No regression vs the pre-Firecrawl state.

import { writeFileSync, mkdirSync, readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_PATH = resolve(__dirname, "../public/data/sentiment-extras.json");

const API_KEY = process.env.FIRECRAWL_API_KEY;
if (!API_KEY) {
  console.error("FIRECRAWL_API_KEY env var not set. Writing empty payload.");
  writeStub("FIRECRAWL_API_KEY not set");
  process.exit(0);
}

const FIRECRAWL_URL = "https://api.firecrawl.dev/v1/scrape";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Diagnostics that survive the workflow run — appended to the committed
// sentiment-extras.json under `_debug` so we can inspect them without
// needing access to the GHA log viewer. Trimmed to ~20 entries to keep
// file size sane.
const DEBUG_CALLS = [];
function pushDebug(entry) {
  DEBUG_CALLS.push({ at: new Date().toISOString(), ...entry });
  if (DEBUG_CALLS.length > 30) DEBUG_CALLS.shift();
}

// Wraps the Firecrawl /v1/scrape call. Returns { markdown, html, json, status }
// or null on failure. `formats` is one of ["markdown", "html", "rawHtml"].
// On failure we log the response status + body sample so we can diagnose
// from the workflow log without re-running the API.
async function firecrawl(url, { formats = ["html"], timeoutMs = 45000 } = {}) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(FIRECRAWL_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${API_KEY}`,
        "Content-Type": "application/json",
      },
      signal: ctrl.signal,
      body: JSON.stringify({ url, formats, onlyMainContent: false }),
    });
    const bodyText = await r.text();
    if (!r.ok) {
      console.log(`  Firecrawl HTTP ${r.status} for ${url} — body: ${bodyText.slice(0, 200)}`);
      pushDebug({ url, fc_status: r.status, fc_body_sample: bodyText.slice(0, 250), outcome: "fc_http_error" });
      return null;
    }
    let j;
    try { j = JSON.parse(bodyText); }
    catch (e) {
      console.log(`  Firecrawl non-JSON response for ${url}: ${bodyText.slice(0, 200)}`);
      pushDebug({ url, fc_status: r.status, fc_body_sample: bodyText.slice(0, 250), outcome: "fc_non_json" });
      return null;
    }
    if (!j?.success) {
      console.log(`  Firecrawl success=false for ${url}: ${JSON.stringify(j).slice(0, 300)}`);
      pushDebug({ url, fc_status: r.status, fc_response: j, outcome: "fc_success_false" });
      return null;
    }
    const data = j.data || null;
    if (!data) {
      console.log(`  Firecrawl success but data is null for ${url}: ${JSON.stringify(j).slice(0, 300)}`);
      pushDebug({ url, fc_status: r.status, outcome: "fc_data_null" });
    } else {
      const htmlLen = (data.html || "").length;
      const rawLen = (data.rawHtml || "").length;
      const mdLen = (data.markdown || "").length;
      const statusCode = data.metadata?.statusCode || data.metadata?.["sourceURL-status"] || "?";
      console.log(`  Firecrawl OK for ${url} — upstream status=${statusCode}, html=${htmlLen}B rawHtml=${rawLen}B md=${mdLen}B`);
      pushDebug({
        url,
        fc_status: r.status,
        upstream_status: statusCode,
        html_len: htmlLen,
        raw_len: rawLen,
        md_len: mdLen,
        body_sample: ((data.rawHtml || data.html || data.markdown) || "").slice(0, 400),
        outcome: "ok",
      });
    }
    return data;
  } catch (err) {
    console.log(`  Firecrawl error for ${url}: ${err.message}`);
    return null;
  } finally {
    clearTimeout(timer);
  }
}

// ---------- 1. Put-Call Ratio ----------

async function fetchPCR() {
  // The 28 May diagnostic run showed NSE's option-chain HTML page loads
  // perfectly through Firecrawl (482 KB HTML, 138 KB markdown), but plain
  // regex extraction is fragile. So Path A uses Firecrawl's LLM-based
  // JSON extract on that page — explicit, costs ~5 credits/day, robust
  // to layout changes. Path B is a free-text regex fallback for the
  // same page in case extract fails. No other sources are tried; the
  // diagnostic run proved Trendlyne / MoneyControl / Sensibull / Upstox
  // / NiftyTrader are all 404 or captcha-blocked.
  const nseUrl = "https://www.nseindia.com/option-chain";

  // Path A — LLM extract.
  console.log("PCR path A: NSE option-chain via Firecrawl LLM extract...");
  const extracted = await firecrawlExtract(nseUrl, {
    schema: {
      type: "object",
      properties: {
        nifty_pcr: {
          type: "number",
          description: "The NIFTY 50 index Put-Call Ratio (PCR), typically a number between 0.3 and 3.0. Look for a stat labeled 'PCR', 'Put-Call Ratio', or 'Put Call Ratio'.",
        },
      },
      required: ["nifty_pcr"],
    },
    timeoutMs: 90000,
  });
  if (extracted && typeof extracted.nifty_pcr === "number") {
    const pcr = Math.round(extracted.nifty_pcr * 100) / 100;
    if (pcr > 0 && pcr < 5) {
      return { value: pcr, source: "NSE option-chain (via Firecrawl LLM extract)", basis: "LLM-extracted" };
    }
    console.log(`  Extract returned out-of-range value: ${extracted.nifty_pcr}`);
  }

  // Path B — same page, regex fallback in case LLM extract is unavailable.
  console.log("PCR path B: NSE option-chain via Firecrawl regex fallback...");
  const r2 = await firecrawl(nseUrl, { formats: ["markdown", "html"], timeoutMs: 60000 });
  if (r2) {
    const text = (r2.markdown || "") + "\n" + (r2.html || "");
    const patterns = [
      /(?:Put[-\s]Call\s*Ratio|PCR)\s*[:\-]?\s*(\d+\.\d{1,3})/i,
      /(?:Put\s*Call\s*Ratio)[^0-9]{0,80}(\d+\.\d{1,3})/i,
      /PCR\s*[^0-9]{0,12}(\d+\.\d{1,3})/i,
    ];
    for (const pat of patterns) {
      const m = text.match(pat);
      if (m) {
        const pcr = Number(m[1]);
        if (Number.isFinite(pcr) && pcr > 0 && pcr < 5) {
          return { value: pcr, source: "NSE option-chain (via Firecrawl regex)", basis: m[0].slice(0, 60) };
        }
      }
    }
  }

  return null;
}

// Firecrawl LLM-based structured extraction (~5 credits/call). Returns the
// extracted object or null.
async function firecrawlExtract(url, { schema, timeoutMs = 60000 } = {}) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(FIRECRAWL_URL, {
      method: "POST",
      headers: { "Authorization": `Bearer ${API_KEY}`, "Content-Type": "application/json" },
      signal: ctrl.signal,
      body: JSON.stringify({
        url,
        formats: ["json"],
        jsonOptions: { schema },
        onlyMainContent: false,
      }),
    });
    const bodyText = await r.text();
    if (!r.ok) {
      console.log(`  Firecrawl extract HTTP ${r.status}: ${bodyText.slice(0, 200)}`);
      pushDebug({ url, fc_status: r.status, fc_body_sample: bodyText.slice(0, 250), outcome: "fc_extract_http_error" });
      return null;
    }
    let j;
    try { j = JSON.parse(bodyText); }
    catch {
      pushDebug({ url, fc_status: r.status, fc_body_sample: bodyText.slice(0, 250), outcome: "fc_extract_non_json" });
      return null;
    }
    const data = j?.data || null;
    const extract = data?.json || data?.llm_extraction || data?.extract || null;
    pushDebug({
      url,
      fc_status: r.status,
      outcome: "extract_ok",
      extract_keys: extract ? Object.keys(extract) : null,
      extract_value: extract ? JSON.stringify(extract).slice(0, 300) : null,
    });
    return extract;
  } catch (err) {
    pushDebug({ url, outcome: "fc_extract_error", error: err.message });
    return null;
  } finally {
    clearTimeout(timer);
  }
}

// ---------- 2. Impact Cost ----------

const MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function recentMonths(count) {
  const out = [];
  const now = new Date();
  let y = now.getUTCFullYear();
  let m = now.getUTCMonth();
  for (let i = 0; i < count; i++) {
    out.push({ year: y, monthIdx: m, abbr: MONTH_ABBR[m] });
    m--;
    if (m < 0) { m = 11; y--; }
  }
  return out;
}

function urlCandidates({ year, abbr }) {
  const A = abbr;
  const a = abbr.toLowerCase();
  return [
    `https://archives.nseindia.com/content/equities/impact_cost/imc_${A}_${year}.csv`,
    `https://nsearchives.nseindia.com/content/equities/impact_cost/imc_${A}_${year}.csv`,
    `https://archives.nseindia.com/content/equities/impact_cost/imc_${A}${year}.csv`,
    `https://nsearchives.nseindia.com/content/equities/impact_cost/imc_${A}${year}.csv`,
    `https://archives.nseindia.com/content/equities/impact_cost/imc${a}${year}.csv`,
    `https://nsearchives.nseindia.com/content/equities/impact_cost/imc${a}${year}.csv`,
  ];
}

function parseImpactCostCSV(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) return null;
  const headerCols = lines[0].split(",").map((h) => h.trim().replace(/^"|"$/g, ""));
  const symIdx = headerCols.findIndex((h) => /^symbol$/i.test(h));
  if (symIdx === -1) return null;
  let icIdx = headerCols.findIndex((h) => /^impact[_\s]?cost$/i.test(h));
  if (icIdx === -1) icIdx = headerCols.findIndex((h) => /impact.*cost/i.test(h));
  if (icIdx === -1) return null;
  const out = {};
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",").map((c) => c.trim().replace(/^"|"$/g, ""));
    const sym = cols[symIdx]?.toUpperCase();
    const raw = cols[icIdx];
    if (!sym || !raw) continue;
    const v = parseFloat(raw.replace(/[^0-9.\-]/g, ""));
    if (!Number.isFinite(v)) continue;
    out[sym] = v;
  }
  return Object.keys(out).length ? out : null;
}

async function fetchImpactCost() {
  // After two rounds of probing — once direct, once via Firecrawl's
  // residential-proxy headless browser — NSE's public Impact Cost
  // distribution is genuinely gone. The product page itself
  // (/products-services/equity-market-impact-cost) returns 404
  // server-side, no aggregator republishes the per-ticker numbers, and
  // we've burnt ~25 Firecrawl credits proving it. Treat the rule as
  // deferred — see rule-meta entry for the rationale. We do not call
  // Firecrawl from this code path anymore.
  console.log("Impact cost: skipped — NSE removed the public CSV (confirmed via Firecrawl).");
  return null;
}

// ---------- main ----------

function writeStub(reason) {
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    pcr: null,
    impact_cost: null,
    error: reason,
  }, null, 2) + "\n");
}

async function main() {
  // Read existing file to preserve any previous successful values if
  // today's fetch happens to fail (PCR especially — yesterday's value
  // is far more useful than null).
  let prev = {};
  try { prev = JSON.parse(readFileSync(OUT_PATH, "utf8")); } catch { /* first run */ }

  const out = {
    generated_at: new Date().toISOString(),
    pcr: null,
    impact_cost: null,
  };

  const pcr = await fetchPCR();
  if (pcr) {
    out.pcr = { ...pcr, fetched_at: new Date().toISOString() };
    console.log(`PCR: ${pcr.value} (${pcr.source})`);
  } else if (prev.pcr) {
    out.pcr = { ...prev.pcr, stale: true };
    console.log(`PCR fetch failed — retained yesterday's value ${prev.pcr.value} (marked stale).`);
  } else {
    console.log("PCR: both paths failed and no cached value to fall back on.");
  }

  const impact = await fetchImpactCost();
  if (impact) {
    out.impact_cost = { ...impact, fetched_at: new Date().toISOString() };
    console.log(`Impact cost: ${Object.keys(impact.companies).length} tickers (period ${impact.period})`);
  } else if (prev.impact_cost) {
    out.impact_cost = { ...prev.impact_cost, stale: true };
    console.log(`Impact cost fetch failed — retained previous (${prev.impact_cost.period}, ${Object.keys(prev.impact_cost.companies || {}).length} tickers, marked stale).`);
  } else {
    console.log("Impact cost: all probes missed and no cached value.");
  }

  // Stamp diagnostics into the output so we can inspect from the
  // committed JSON file without needing GHA log access.
  out._debug = {
    api_key_set: !!API_KEY,
    api_key_prefix: API_KEY ? API_KEY.slice(0, 6) + "..." : null,
    calls: DEBUG_CALLS,
  };

  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(out, null, 2) + "\n");
  console.log(`\nWrote: ${OUT_PATH}`);
}

main().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  writeStub(err.message);
  process.exit(0);
});
