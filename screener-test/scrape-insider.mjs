#!/usr/bin/env node
// NSE PIT (Prohibition of Insider Trading) scraper. Pulls insider-trade
// disclosures filed in the last ~180 calendar days, aggregates per ticker
// (buy shares, sell shares, buy value, sell value, transaction count, last
// trade date), and writes public/data/insider-trades.json which the
// Fundamentals scoring rule reads.

import { writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_PATH = resolve(__dirname, "../public/data/insider-trades.json");

const LOOKBACK_DAYS = 180;
const CHUNK_DAYS = 7;       // NSE PIT silently caps recent chunks at ~316
                            // rows even though older 60-day chunks return
                            // 2000+. 7-day chunks stay under the cap on
                            // every period we've measured. Total ~26 calls.
const CAP_WARNING_THRESHOLD = 300;  // log a canary if any chunk approaches the cap

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------- main ----------
async function run() {
  console.log("Initialising NSE session...");
  await initNSESession();

  const today = new Date();
  const trades = [];
  let chunks = 0, chunkOk = 0;
  for (let offset = 0; offset < LOOKBACK_DAYS; offset += CHUNK_DAYS) {
    const toD   = new Date(today.getTime() - offset * 86400000);
    const fromD = new Date(toD.getTime() - CHUNK_DAYS * 86400000);
    chunks++;
    process.stdout.write(`Chunk ${chunks} (${fmtDate(fromD)} → ${fmtDate(toD)})... `);
    try {
      const data = await fetchPIT(fmtDate(fromD), fmtDate(toD));
      const rows = data?.data || [];
      trades.push(...rows);
      chunkOk++;
      const cap = rows.length >= CAP_WARNING_THRESHOLD ? `  ⚠ approaching NSE cap (${rows.length} ≥ ${CAP_WARNING_THRESHOLD}) — recent disclosures may be silently truncated; shrink CHUNK_DAYS` : "";
      console.log(`${rows.length} trades${cap}`);
    } catch (err) {
      console.log(`FAIL: ${err.message}`);
    }
    await sleep(400);
  }

  if (chunkOk === 0) {
    console.log("\nNo PIT chunks succeeded — writing empty payload.");
    writeFallback("NSE PIT API unreachable (likely IP-blocked or schema changed). Will retry on next refresh.");
    return;
  }

  console.log(`\nGot ${trades.length} raw trade entries from ${chunkOk}/${chunks} chunks`);
  // Show first row schema so we can debug if NSE changes field names later.
  if (trades[0]) {
    console.log("Sample raw entry keys:", Object.keys(trades[0]).join(", "));
  }

  // -------- DIAGNOSTIC LOGGING --------
  // (1) Count tdpTransactionType distribution across all chunks. If NSE
  //     introduces a new label (e.g. "Disposal" instead of "Sell") that
  //     we don't classify, it'll show up here as a sizable "OTHER"
  //     bucket instead of disappearing into pledgeSkipped silently.
  const typeCounts = {};
  let missingSymbol = 0;
  for (const t of trades) {
    const tt = String(t.tdpTransactionType || "(empty)").trim();
    typeCounts[tt] = (typeCounts[tt] || 0) + 1;
    if (!String(t.symbol || "").trim()) missingSymbol++;
  }
  console.log("\ntdpTransactionType distribution:");
  Object.entries(typeCounts).sort((a, b) => b[1] - a[1]).forEach(([k, n]) => {
    console.log(`  ${k.padEnd(20)} ${n}`);
  });
  if (missingSymbol > 0) console.log(`Entries with missing/blank symbol: ${missingSymbol}`);

  // (2) If WATCH_TICKERS env is set (comma-separated), print every raw
  //     entry for those tickers — schema + values — so we can see what
  //     NSE is actually returning. Used to debug "the form shows a sell
  //     but our aggregate doesn't" mismatches.
  const watch = (process.env.WATCH_TICKERS || "").split(",").map((s) => s.trim().toUpperCase()).filter(Boolean);
  if (watch.length) {
    console.log(`\n=== Watching tickers: ${watch.join(", ")} ===`);
    const debugRows = [];
    for (const t of trades) {
      const sym = String(t.symbol || "").trim().toUpperCase();
      if (!watch.includes(sym)) continue;
      const row = {
        symbol: t.symbol,
        type: t.tdpTransactionType,
        secAcq: t.secAcq,
        secVal: t.secVal,
        acqMode: t.acqMode,
        personCategory: t.personCategory,
        date: t.date,
        intimDt: t.intimDt,
        acqfromDt: t.acqfromDt,
        acqtoDt: t.acqtoDt,
      };
      debugRows.push(row);
      console.log("  ROW:", JSON.stringify(row));
    }
    console.log(`Total rows seen for watched tickers: ${debugRows.length}`);
    // Also persist the debug dump to a committed file so it's auditable
    // after the workflow run finishes, without needing to scroll GitHub
    // Actions logs (which require auth to read remotely).
    // EXPERIMENT: also try the per-symbol variant of the same endpoint
    // for the LAST 30 DAYS — this is the window where the date-range
    // endpoint returned 0 rows (Chunks 1-4 in the previous run). If
    // NSE's API returns INOX's May Sell disclosures when filtered by
    // symbol, we know the fix path: switch to per-symbol queries for
    // the recent window. If even this returns nothing, escalate to
    // Firecrawl on the NSE web UI.
    const today = new Date();
    const thirtyDaysAgo = new Date(today.getTime() - 30 * 86400000);
    const symbolResults = {};
    for (const sym of watch) {
      try {
        console.log(`\n--- Per-symbol experiment for ${sym} (${fmtDate(thirtyDaysAgo)} → ${fmtDate(today)}) ---`);
        const data = await fetchPIT(fmtDate(thirtyDaysAgo), fmtDate(today), sym);
        const rows = data?.data || [];
        console.log(`  Returned ${rows.length} rows via ?symbol=${sym}`);
        symbolResults[sym] = {
          window: `${fmtDate(thirtyDaysAgo)} → ${fmtDate(today)}`,
          count: rows.length,
          rows: rows.slice(0, 50).map((t) => ({
            symbol: t.symbol,
            type: t.tdpTransactionType,
            secAcq: t.secAcq,
            secVal: t.secVal,
            acqMode: t.acqMode,
            personCategory: t.personCategory,
            date: t.date,
            intimDt: t.intimDt,
            acqfromDt: t.acqfromDt,
            acqtoDt: t.acqtoDt,
          })),
        };
        for (const r of rows.slice(0, 5)) {
          console.log("    ROW:", JSON.stringify({ type: r.tdpTransactionType, secAcq: r.secAcq, date: r.date, acqtoDt: r.acqtoDt }));
        }
        await sleep(800);
      } catch (err) {
        console.log(`  FAIL: ${err.message}`);
        symbolResults[sym] = { error: err.message };
      }
    }

    const DEBUG_PATH = resolve(__dirname, "../public/data/insider-debug.json");
    writeFileSync(DEBUG_PATH, JSON.stringify({
      generated_at: new Date().toISOString(),
      watch_tickers: watch,
      total_trades_in_pull: trades.length,
      type_distribution: typeCounts,
      date_range_rows_seen: debugRows.length,
      date_range_rows: debugRows,
      per_symbol_experiment: symbolResults,
      firecrawl_response: LAST_FIRECRAWL_RESPONSE,
    }, null, 2) + "\n");
    console.log(`\nWrote debug dump to ${DEBUG_PATH}`);
  }

  // ---- Firecrawl supplement for the last ~30 days ----
  // The NSE date-range API has a confirmed ~28-day blackout on recent
  // disclosures (proven in PRs #93-96 — INOX India had 3 May 2026
  // Sell disclosures visible on the NSE web UI but the API returned 0
  // entries for the May 4 → June 1 window). The web UI uses a different
  // backend we don't have public API access to, but Firecrawl can render
  // the page and LLM-extract the table.
  // Record the env-check BEFORE we try Firecrawl, so the debug file
  // shows the truth even if Firecrawl never runs / silently throws.
  LAST_FIRECRAWL_RESPONSE = {
    env_check: {
      has_api_key: !!process.env.FIRECRAWL_API_KEY,
      api_key_length: (process.env.FIRECRAWL_API_KEY || "").length,
      api_key_prefix: (process.env.FIRECRAWL_API_KEY || "").slice(0, 3),
    },
    called_firecrawl: false,
  };
  console.log(`\nFirecrawl env-check: has_key=${LAST_FIRECRAWL_RESPONSE.env_check.has_api_key} length=${LAST_FIRECRAWL_RESPONSE.env_check.api_key_length}`);

  let firecrawlAdded = 0;
  if (process.env.FIRECRAWL_API_KEY) {
    LAST_FIRECRAWL_RESPONSE.called_firecrawl = true;
    try {
      const fcRows = await scrapeRecentViaFirecrawl();
      firecrawlAdded = fcRows.length;
      // Dedupe vs the API-pulled rows (same disclosure could appear in
      // both for the ~April overlap window). Key on symbol + acqfromDt
      // + secAcq + tdpTransactionType; API wins on collision because
      // its values are reliably structured.
      const seen = new Set();
      for (const t of trades) {
        seen.add([
          String(t.symbol || "").toUpperCase(),
          String(t.acqfromDt || ""),
          String(t.secAcq || ""),
          String(t.tdpTransactionType || ""),
        ].join("|"));
      }
      for (const f of fcRows) {
        const k = [
          String(f.symbol || "").toUpperCase(),
          String(f.acqfromDt || ""),
          String(f.secAcq || ""),
          String(f.tdpTransactionType || ""),
        ].join("|");
        if (!seen.has(k)) { trades.push(f); seen.add(k); }
      }
      console.log(`Firecrawl added ${fcRows.length} recent disclosures (${trades.length} total after dedupe)`);
    } catch (err) {
      console.log(`Firecrawl supplement failed (non-fatal): ${err.message}`);
      // Capture error in debug too so we can see it without log access.
      LAST_FIRECRAWL_RESPONSE.scrape_error = err.message;
    }
  } else {
    console.log("FIRECRAWL_API_KEY not set — skipping recent-window supplement.");
  }

  const perTicker = aggregate(trades);
  const tickerCount = Object.keys(perTicker).length;
  console.log(`Aggregated into ${tickerCount} tickers.`);

  const payload = {
    generated_at: new Date().toISOString(),
    source: "NSE corporates-pit API",
    lookback_days: LOOKBACK_DAYS,
    total_trades: trades.length,
    chunks_ok: chunkOk,
    chunks_total: chunks,
    companies: perTicker,
  };
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload) + "\n");
  console.log(`\nWrote: ${OUT_PATH}`);
}

// Firecrawl LLM-extract of NSE's web UI insider-trading page. Renders
// the JS-heavy page with a headless browser, then an LLM pulls out the
// structured disclosure rows. We map each row back to the API's field
// shape (symbol/tdpTransactionType/secAcq/secVal/personCategory/dates)
// so downstream aggregate() can process them identically.
const FIRECRAWL_INSIDER_SCHEMA = {
  type: "object",
  properties: {
    disclosures: {
      type: "array",
      description: "Insider trading disclosures (SEBI Reg 7(2)) on this page. INCLUDE every visible row — the table may show multiple weeks of activity per company. Each row is one disclosure: a Buy, Sell, Pledge, Pledge Release, or Pledge Invocation. Extract values exactly as written (don't compute or estimate).",
      items: {
        type: "object",
        properties: {
          symbol: { type: "string", description: "NSE ticker symbol — uppercase letters only (e.g. INOXINDIA, RELIANCE). Required." },
          company_name: { type: "string", description: "Full company name as shown." },
          person_name: { type: "string", description: "Name of person/entity filing the disclosure." },
          person_category: { type: "string", description: "e.g. Promoter, Promoter Group, Director, KMP." },
          transaction_type: { type: "string", description: "Exactly one of: Buy, Sell, Pledge, Pledge Revoke, Pledge Invoke. Use the page's wording, just match case." },
          shares: { type: "number", description: "Number of securities acquired or disposed (the middle 'No. of security' column under 'Securities acquired / disposed')." },
          value_rs: { type: "number", description: "Value of the transaction in INR. May be 0 if unreported." },
          from_date: { type: "string", description: "Acquisition / disposal FROM date in DD-MM-YYYY format." },
          to_date: { type: "string", description: "Acquisition / disposal TO date in DD-MM-YYYY format." },
          filing_date: { type: "string", description: "Disclosure filing date in DD-MM-YYYY format." },
        },
        required: ["symbol", "transaction_type", "shares"],
      },
    },
  },
  required: ["disclosures"],
};

// Captured on every call so the debug dump can include WHAT Firecrawl
// saw (upstream status, raw extracted JSON, response shape) — useful
// when extraction returns 0 rows and we need to know why without
// hitting auth-gated GitHub Actions logs.
let LAST_FIRECRAWL_RESPONSE = null;

async function scrapeRecentViaFirecrawl() {
  const url = "https://www.nseindia.com/companies-listing/corporate-filings-insider-trading";
  console.log("\nFirecrawl: fetching recent disclosures from NSE web UI...");
  const r = await fetch("https://api.firecrawl.dev/v1/scrape", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.FIRECRAWL_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      url,
      formats: ["json", "markdown"],   // also pull markdown so we can SEE what page Firecrawl rendered
      jsonOptions: { schema: FIRECRAWL_INSIDER_SCHEMA },
      onlyMainContent: false,
    }),
  });
  if (!r.ok) {
    const body = await r.text();
    Object.assign(LAST_FIRECRAWL_RESPONSE || {}, { http_status: r.status, error_body: body.slice(0, 1000) });
    throw new Error(`Firecrawl HTTP ${r.status}: ${body.slice(0, 200)}`);
  }
  const j = await r.json();
  const extract = j?.data?.json || j?.data?.llm_extraction || j?.data?.extract || {};
  const upstream = j?.data?.metadata?.statusCode || null;
  const md = j?.data?.markdown || "";
  Object.assign(LAST_FIRECRAWL_RESPONSE || {}, {
    http_status: r.status,
    upstream_status: upstream,
    metadata: j?.data?.metadata || null,
    extract_top_level_keys: Object.keys(extract),
    disclosures_count_raw: Array.isArray(extract.disclosures) ? extract.disclosures.length : -1,
    sample_extracted_rows: Array.isArray(extract.disclosures) ? extract.disclosures.slice(0, 10) : [],
    markdown_length: md.length,
    markdown_head: md.slice(0, 2000),
    markdown_contains_inoxindia: /INOXINDIA/i.test(md),
  });
  if (upstream && upstream !== 200) {
    console.log(`  Firecrawl: upstream NSE page returned ${upstream} — refusing extracted rows to avoid hallucination`);
    return [];
  }
  const disclosures = Array.isArray(extract.disclosures) ? extract.disclosures : [];
  console.log(`  Firecrawl: extracted ${disclosures.length} disclosure rows from NSE web UI`);
  console.log(`  Firecrawl markdown length: ${md.length}, contains INOXINDIA: ${LAST_FIRECRAWL_RESPONSE.markdown_contains_inoxindia}`);
  // Map each row to the same shape as NSE's API response so aggregate()
  // can process it unchanged.
  return disclosures
    .filter((d) => d && d.symbol && d.transaction_type && d.shares != null)
    .map((d) => ({
      symbol: String(d.symbol).trim().toUpperCase(),
      tdpTransactionType: String(d.transaction_type).trim(),
      secAcq: String(d.shares),
      secVal: String(d.value_rs || 0),
      personCategory: d.person_category || "",
      acqfromDt: d.from_date || "",
      acqtoDt: d.to_date || "",
      intimDt: d.filing_date || "",
      date: d.filing_date || "",
      _source: "firecrawl",
    }));
}

function writeFallback(reason) {
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    source: "NSE corporates-pit API",
    lookback_days: LOOKBACK_DAYS,
    total_trades: 0,
    chunks_ok: 0,
    error: reason,
    companies: {},
  }) + "\n");
  console.log(`Wrote empty payload: ${OUT_PATH}`);
}

// ---------- NSE session + fetch ----------
const NSE_HEADERS = {
  "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
  "Accept": "*/*",
  "Accept-Language": "en-US,en;q=0.9",
  "Referer": "https://www.nseindia.com/companies-listing/corporate-filings-insider-trading",
  "Connection": "keep-alive",
};

let cookieJar = "";

function saveCookies(response) {
  const setCookies = response.headers.getSetCookie?.() || [];
  for (const sc of setCookies) {
    const [kv] = sc.split(";");
    const eq = kv.indexOf("=");
    if (eq < 0) continue;
    const k = kv.slice(0, eq).trim();
    const v = kv.slice(eq + 1).trim();
    const existing = new RegExp(`(^|; )${k}=[^;]*`);
    if (existing.test(cookieJar)) cookieJar = cookieJar.replace(existing, (m, p) => `${p}${k}=${v}`);
    else cookieJar += (cookieJar ? "; " : "") + `${k}=${v}`;
  }
}

async function initNSESession() {
  cookieJar = "";
  const warmupUrls = [
    "https://www.nseindia.com/",
    "https://www.nseindia.com/companies-listing/corporate-filings-insider-trading",
  ];
  for (const url of warmupUrls) {
    const r = await fetch(url, { headers: NSE_HEADERS });
    saveCookies(r);
    await sleep(600);
  }
}

async function fetchPIT(fromDate, toDate, symbol = null) {
  const symbolParam = symbol ? `&symbol=${encodeURIComponent(symbol)}` : "";
  const url = `https://www.nseindia.com/api/corporates-pit?index=equities&from_date=${fromDate}&to_date=${toDate}${symbolParam}`;
  let attempt = 0;
  while (attempt < 3) {
    try {
      const r = await fetch(url, {
        headers: { ...NSE_HEADERS, Cookie: cookieJar, Accept: "application/json" },
      });
      if (r.status === 401 || r.status === 403) {
        // Re-initialise session and retry
        await initNSESession();
        attempt++;
        continue;
      }
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      saveCookies(r);
      return await r.json();
    } catch (err) {
      attempt++;
      if (attempt >= 3) throw err;
      await sleep(1500 * attempt);
    }
  }
}

function fmtDate(d) {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}-${mm}-${d.getUTCFullYear()}`;
}

// ---------- aggregation ----------
function num(v) { const n = Number(String(v ?? "").replace(/,/g, "")); return Number.isFinite(n) ? n : 0; }

function aggregate(trades) {
  const out = {};
  let pledgeSkipped = 0;
  for (const t of trades) {
    const sym = String(t.symbol || "").trim().toUpperCase();
    if (!sym) continue;

    // Actual NSE PIT values (probed via PR #15):
    //   buyValue / sellValue / buyQuantity / sellquantity are always "0"
    //   placeholders. Real values live in:
    //     secAcq             – share count (regardless of direction)
    //     secVal             – transaction value in ₹
    //     tdpTransactionType – "Buy" | "Sell" | "Pledge" | "Pledge Release"
    //     acqMode            – "Market Purchase" | "Market Sale" | "ESOP" | ...
    //     personCategory     – "Promoters" | "Director" | "Promoter Group" | ...
    const shares = num(t.secAcq);
    const value  = num(t.secVal);
    const type   = String(t.tdpTransactionType || "").trim();

    if (!out[sym]) out[sym] = {
      buy_shares: 0, sell_shares: 0, buy_value: 0, sell_value: 0,
      transactions: 0, pledges_excluded: 0, last_date: null,
    };

    if (type === "Buy") {
      out[sym].buy_shares += shares;
      out[sym].buy_value  += value;
    } else if (type === "Sell") {
      out[sym].sell_shares += shares;
      out[sym].sell_value  += value;
    } else {
      // Pledge / Pledge Release / Invocation — not a buy or sell of beneficial
      // ownership. Skip from the scoring count (the client rule is about
      // net BUYING vs SELLING, not collateral pledges). We still TRACK these
      // per-ticker so the rule cell can transparently say "X trades + Y
      // pledges excluded" — otherwise a user comparing our number to the
      // NSE PIT page (which lists everything) will be confused.
      pledgeSkipped++;
      out[sym].pledges_excluded = (out[sym].pledges_excluded || 0) + 1;
      continue;
    }
    out[sym].transactions++;

    const date = t.date || t.intimDt || t.acqtoDt;
    if (date && (!out[sym].last_date || dateGt(date, out[sym].last_date))) {
      out[sym].last_date = date;
    }
  }
  console.log(`  Skipped ${pledgeSkipped} pledge/non-buy-sell disclosures.`);
  // Derive net + round
  for (const sym in out) {
    const o = out[sym];
    o.net_shares = o.buy_shares - o.sell_shares;
    o.net_value  = Math.round((o.buy_value - o.sell_value) * 100) / 100;
    o.buy_value  = Math.round(o.buy_value  * 100) / 100;
    o.sell_value = Math.round(o.sell_value * 100) / 100;
  }
  return out;
}

function dateGt(a, b) {
  // Both NSE strings (e.g. "15-Jan-2026"); naive string compare is enough.
  return new Date(a) > new Date(b);
}

// ---------- entry ----------
// Invoked at the bottom so module-level let/const (cookieJar, NSE_HEADERS)
// reach their declarations before initNSESession() touches them.
run().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  writeFallback(err.message);
  process.exit(0);
});
