#!/usr/bin/env node
// Insider-trade scraper backed by the muns.io filings API
//   POST https://devde.muns.io/filings/data/insider_trades
//   body: { "ticker": "RELIANCE", "country": "India" }
//   auth: Bearer <MUNS_TOKEN>
//   response: a MARKDOWN PIPE TABLE (not JSON, despite the Accept header) —
//     columns: Company, Insider, Category, Security Type, Transaction, Trade
//     Shares, Trade %, Trade Value, Post Holding Shares, Post Holding %,
//     Mode, From Date, To Date, Broadcast Date, Source.
//
// Replaces the old NSE PIT direct-scrape + Firecrawl supplement. muns.io
// already aggregates BSE + Trendlyne disclosures with full history and no
// ~28-day blackout, so the cookie dance, 7-day chunking, and Firecrawl
// LLM-extract are all gone.
//
// We iterate every ticker in the Nifty-500 universe (public/data/
// screener-companies.json), call the API once per ticker, keep disclosures
// from the last LOOKBACK_DAYS, and aggregate per ticker into the SAME
// schema the dashboard already reads:
//
//   companies: { TICKER: { buy_shares, sell_shares, buy_value, sell_value,
//                          transactions, pledges_excluded, last_date,
//                          net_shares, net_value } }
//
// so nothing on the front-end or in scoring.js has to change.
//
// Auth token comes from process.env.MUNS_TOKEN (a GitHub Actions secret of
// the same name). It is NEVER logged — we only print presence + length.
//
// Two offline helpers for development / validation (no network, no token):
//   MUNS_FIXTURE=/path/to/sample.json   run a saved API response through the
//                                        parser+aggregator and print the result
//   MUNS_PROBE_TICKER=RELIANCE          fetch ONE ticker live and dump the raw
//                                        shape to public/data/insider-debug.json

import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const UNIVERSE_PATH = resolve(__dirname, "../public/data/screener-companies.json");
const OUT_PATH      = resolve(__dirname, "../public/data/insider-trades.json");
const DEBUG_PATH    = resolve(__dirname, "../public/data/insider-debug.json");

const API_URL       = "https://devde.muns.io/filings/data/insider_trades";
const COUNTRY       = process.env.MUNS_COUNTRY || "India";
const LOOKBACK_DAYS = Number(process.env.MUNS_LOOKBACK_DAYS || 180);
const CONCURRENCY   = Math.max(1, Number(process.env.MUNS_CONCURRENCY || 5));
const DELAY_MS      = Number(process.env.MUNS_DELAY_MS || 120);   // gap between request launches
const MAX_RETRIES   = Number(process.env.MUNS_MAX_RETRIES || 3);
const PROBE_TICKER  = (process.env.MUNS_PROBE_TICKER || "").trim().toUpperCase();
const FIXTURE_PATH  = (process.env.MUNS_FIXTURE || "").trim();
const TICKERS_ENV   = (process.env.MUNS_TICKERS || "").split(",").map((s) => s.trim().toUpperCase()).filter(Boolean);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------- response parsing ----------
// The muns.io endpoint returns its result as a MARKDOWN PIPE TABLE (despite
// the application/json Accept header) — the body literally starts with
// "| Company | Insider | ...". parseRows() handles that, and still accepts
// real JSON in case the API changes later. fieldGetter matches columns by a
// normalised (lowercase, alphanumeric-only) key, so the table headers
// ("Trade Shares", "Broadcast Date", …) resolve without hard-coding spelling.

function extractRows(json) {
  const PREFERRED = ["data", "insider_trades", "insiderTrades", "result", "results",
                     "rows", "filings", "trades", "records", "items"];
  const isRowArray = (v) =>
    Array.isArray(v) && (v.length === 0 || (v[0] && typeof v[0] === "object" && !Array.isArray(v[0])));
  function dig(o, depth) {
    if (Array.isArray(o)) return o;
    if (!o || typeof o !== "object" || depth > 2) return null;
    for (const k of PREFERRED) if (Array.isArray(o[k])) return o[k];
    for (const k of PREFERRED) if (o[k] && typeof o[k] === "object") {
      const r = dig(o[k], depth + 1); if (r) return r;
    }
    for (const v of Object.values(o)) if (isRowArray(v)) return v;
    return null;
  }
  return dig(json, 0) || [];
}

// Sniff the response: a markdown pipe-table is parsed column-by-column into
// row objects keyed by the header labels; anything starting with [ or { is
// tried as JSON (future-proofing). Returns an array of row objects either way.
function parseRows(body) {
  const text = String(body || "").trim();
  if (!text) return [];
  if (text[0] === "[" || text[0] === "{") {
    try { return extractRows(JSON.parse(text)); } catch { /* not JSON — fall through */ }
  }
  if (text.includes("|")) return parseMarkdownTable(text);
  return [];
}

function parseMarkdownTable(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.includes("|"));
  if (lines.length < 2) return [];
  const cells = (l) => {
    let s = l;
    if (s.startsWith("|")) s = s.slice(1);
    if (s.endsWith("|")) s = s.slice(0, -1);
    return s.split("|").map((c) => c.trim());
  };
  const header = cells(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (/^[\s|:\-]+$/.test(lines[i])) continue;   // markdown separator row (---|---)
    const c = cells(lines[i]);
    const obj = {};
    header.forEach((h, j) => { obj[h] = c[j] !== undefined ? c[j] : ""; });
    rows.push(obj);
  }
  return rows;
}

const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");

function fieldGetter(row) {
  const map = {};
  for (const [k, v] of Object.entries(row)) map[norm(k)] = v;
  return (aliases) => {
    for (const a of aliases) if (a in map && map[a] != null && map[a] !== "") return map[a];
    return undefined;
  };
}

// Column → candidate normalised key names. Order doesn't matter; first hit wins.
const F = {
  shares:    ["tradeshares", "shares", "quantity", "qty", "noofsecurities",
              "numberofsecurities", "securities", "secacq", "securitiestraded"],
  value:     ["tradevalue", "value", "transactionvalue", "amount", "secval", "tradeamount"],
  type:      ["transaction", "transactiontype", "type", "tradetype", "action",
              "dealtype", "tdptransactiontype", "acquisitiondisposal"],
  mode:      ["mode", "acqmode", "transactionmode", "manner", "modeofacquisition"],
  broadcast: ["broadcastdate", "broadcast", "disseminationdate", "intimationdate",
              "intimdt", "filingdate", "disclosuredate", "date"],
  toDate:    ["todate", "acqtodt", "todt", "to", "transactiontodate"],
};

function num(v) {
  if (v == null) return 0;
  const n = Number(String(v).replace(/[,\s₹]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

// Acquisition→buy, Disposal→sell, Pledge/Revoke→excluded (collateral moves,
// not buy/sell of beneficial ownership — parity with the old NSE pipeline).
// `mode` is a secondary signal: anything mentioning a pledge is excluded.
function classify(typeRaw, modeRaw) {
  const t = String(typeRaw || "").toLowerCase().trim();
  const m = String(modeRaw || "").toLowerCase();
  if (/pledge|revoke|revocation|invocat/.test(t) || /pledge/.test(m)) return "pledge";
  if (/acquisition|acquire|acquired|buy|purchase|allot|subscrib/.test(t)) return "buy";
  if (/disposal|dispose|disposed|sell|sale|sold/.test(t)) return "sell";
  return "other";
}

function parseDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  const dmy = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);          // DD-MM-YYYY fallback
  const d = dmy ? new Date(`${dmy[3]}-${dmy[2]}-${dmy[1]}`) : new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}
const isoOf = (d) => (d ? d.toISOString().slice(0, 10) : null);
const round2 = (n) => Math.round(n * 100) / 100;

// Aggregate one ticker's disclosure rows (already the array from extractRows)
// into the dashboard's per-company shape, keeping only the last LOOKBACK_DAYS.
function aggregateTicker(rows, cutoff) {
  const o = {
    buy_shares: 0, sell_shares: 0, buy_value: 0, sell_value: 0,
    transactions: 0, pledges_excluded: 0, last_date: null,
  };
  let inWindow = 0, undated = 0;
  for (const row of rows) {
    const get = fieldGetter(row);
    const d = parseDate(get(F.broadcast)) || parseDate(get(F.toDate));
    if (!d) { undated++; continue; }              // can't confirm recency → skip (conservative)
    if (cutoff && d < cutoff) continue;            // older than the window
    inWindow++;

    const shares = num(get(F.shares));
    const value  = num(get(F.value));
    switch (classify(get(F.type), get(F.mode))) {
      case "buy":    o.buy_shares  += shares; o.buy_value  += value; o.transactions++; break;
      case "sell":   o.sell_shares += shares; o.sell_value += value; o.transactions++; break;
      case "pledge": o.pledges_excluded++; break;
      default:       break;                        // unknown type — tracked via inWindow only
    }
    const ds = isoOf(d);
    if (ds && (!o.last_date || ds > o.last_date)) o.last_date = ds;
  }
  o.net_shares = o.buy_shares - o.sell_shares;
  o.net_value  = round2(o.buy_value - o.sell_value);
  o.buy_value  = round2(o.buy_value);
  o.sell_value = round2(o.sell_value);
  return { entry: o, inWindow, undated };
}

// ---------- live API ----------
async function fetchInsider(ticker) {
  const body = JSON.stringify({ ticker, country: COUNTRY });
  let attempt = 0;
  while (true) {
    attempt++;
    let r;
    try {
      r = await fetch(API_URL, {
        method: "POST",
        headers: {
          accept: "application/json",
          "content-type": "application/json",
          authorization: `Bearer ${process.env.MUNS_TOKEN}`,
        },
        body,
      });
    } catch (err) {
      if (attempt > MAX_RETRIES) throw new Error(`network: ${err.message}`);
      await sleep(500 * attempt);
      continue;
    }
    if (r.status === 401 || r.status === 403) {
      const e = new Error(`auth ${r.status}`); e.auth = true; throw e;
    }
    if (r.status === 429 || r.status >= 500) {
      if (attempt > MAX_RETRIES) throw new Error(`HTTP ${r.status} after ${attempt} attempts`);
      const ra = Number(r.headers.get("retry-after"));
      await sleep(Number.isFinite(ra) && ra > 0 ? ra * 1000 : 600 * attempt);
      continue;
    }
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return await r.text();   // body is a markdown table, not JSON — parse downstream
  }
}

// Bounded-concurrency worker pool over the ticker list.
async function runPool(tickers) {
  const results = {};
  const stats = { ok: 0, empty: 0, fail: 0, auth: 0, errors: [] };
  let i = 0, done = 0;
  async function worker() {
    while (i < tickers.length) {
      const t = tickers[i++];
      try {
        const rows = parseRows(await fetchInsider(t));
        results[t] = rows;
        stats.ok++;
        if (!rows.length) stats.empty++;
      } catch (err) {
        if (err.auth) stats.auth++;
        stats.fail++;
        if (stats.errors.length < 10) stats.errors.push(`${t}: ${err.message}`);
        if (stats.auth >= 5) throw new Error("Too many auth failures — MUNS_TOKEN likely missing/invalid.");
      }
      if (++done % 50 === 0) console.log(`  ...${done}/${tickers.length}`);
      await sleep(DELAY_MS);
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, tickers.length) }, worker));
  return { results, stats };
}

function loadTickers() {
  const raw = JSON.parse(readFileSync(UNIVERSE_PATH, "utf8"));
  const rows = Array.isArray(raw) ? raw : (raw.rows || raw.companies || raw.data || []);
  const seen = new Set(), tickers = [];
  for (const r of rows) {
    const m = String(r["Screener URL"] || "").match(/\/company\/([^/]+)/);
    const t = m ? m[1].toUpperCase() : null;
    if (t && !seen.has(t)) { seen.add(t); tickers.push(t); }
  }
  return tickers;
}

// ---------- modes ----------
function runFixture() {
  // Validate the parser/aggregator against a saved response, offline.
  console.log(`FIXTURE mode — reading ${FIXTURE_PATH}`);
  const body = readFileSync(FIXTURE_PATH, "utf8");
  const rows = parseRows(body);
  console.log(`Parsed ${rows.length} rows (${body.trim()[0] === "|" ? "markdown table" : "json"}).`);
  if (rows[0]) console.log(`Row keys: ${Object.keys(rows[0]).join(", ")}`);
  const breakdown = {};
  for (const row of rows) {
    const get = fieldGetter(row);
    const k = classify(get(F.type), get(F.mode));
    breakdown[k] = (breakdown[k] || 0) + 1;
  }
  console.log(`Classification: ${JSON.stringify(breakdown)}`);
  const cutoff = new Date(Date.now() - LOOKBACK_DAYS * 86400000);
  const { entry, inWindow, undated } = aggregateTicker(rows, cutoff);
  console.log(`In-window rows: ${inWindow}, undated/skipped: ${undated}`);
  console.log(`Aggregated entry:\n${JSON.stringify(entry, null, 2)}`);
}

async function runProbe() {
  console.log(`PROBE mode — fetching one ticker live: ${PROBE_TICKER}`);
  const body = await fetchInsider(PROBE_TICKER);
  const rows = parseRows(body);
  const head = body.trim()[0];
  const parsedAs = head === "|" ? "markdown-table" : (head === "[" || head === "{") ? "json" : "unknown";
  console.log(`Response parsed as: ${parsedAs}. Rows: ${rows.length}.`);
  console.log(`Row keys: ${rows[0] ? Object.keys(rows[0]).join(", ") : "(none)"}`);
  const breakdown = {};
  for (const row of rows) {
    const g = fieldGetter(row);
    const k = classify(g(F.type), g(F.mode));
    breakdown[k] = (breakdown[k] || 0) + 1;
  }
  console.log(`Classification across all rows: ${JSON.stringify(breakdown)}`);
  const cutoff = new Date(Date.now() - LOOKBACK_DAYS * 86400000);
  const { entry, inWindow, undated } = aggregateTicker(rows, cutoff);
  mkdirSync(dirname(DEBUG_PATH), { recursive: true });
  writeFileSync(DEBUG_PATH, JSON.stringify({
    probe: PROBE_TICKER, generated_at: new Date().toISOString(),
    parsed_as: parsedAs, raw_head: body.slice(0, 1200),
    header_columns: rows[0] ? Object.keys(rows[0]) : [],
    row_count: rows.length, classification: breakdown, in_window: inWindow, undated,
    sample_rows: rows.slice(0, 5), aggregated: entry,
  }, null, 2) + "\n");
  console.log(`In-window: ${inWindow}, undated: ${undated}. Aggregated:\n${JSON.stringify(entry, null, 2)}`);
  console.log(`Wrote ${DEBUG_PATH}`);
}

// Merge mode: refresh only an explicit set of tickers into the existing
// insider-trades.json, preserving every other company. Powers the manual,
// incremental 5-at-a-time rollout via the workflow's `tickers` input. Each
// named ticker's entry is REPLACED with its fresh muns aggregate (or removed
// if it has no in-window disclosures); re-running the same ticker is
// idempotent — no additive double-counting.
async function runBatch(tickers) {
  console.log(`BATCH (merge) mode — refreshing ${tickers.length}: ${tickers.join(", ")}`);
  let companies = {};
  try {
    const existing = JSON.parse(readFileSync(OUT_PATH, "utf8"));
    companies = existing.companies || {};
    console.log(`Loaded existing file: ${Object.keys(companies).length} companies preserved.`);
  } catch {
    console.log("No existing insider-trades.json — starting a fresh companies map.");
  }
  const cutoff = new Date(Date.now() - LOOKBACK_DAYS * 86400000);
  const updated = [], emptied = [], failed = [];
  for (const t of tickers) {
    try {
      const { entry } = aggregateTicker(parseRows(await fetchInsider(t)), cutoff);
      if (entry.transactions > 0 || entry.pledges_excluded > 0) {
        companies[t] = entry;
        updated.push(`${t}: ${entry.transactions} trades / ${entry.pledges_excluded} pledges / net ${entry.net_shares >= 0 ? "+" : ""}${entry.net_shares} sh`);
      } else if (companies[t]) {
        delete companies[t];
        emptied.push(`${t} (stale entry removed)`);
      } else {
        emptied.push(t);
      }
    } catch (err) {
      failed.push(`${t}: ${err.message}`);
    }
    await sleep(DELAY_MS);
  }
  if (!updated.length && !emptied.length) {
    console.error(`All ${tickers.length} requested ticker(s) failed — not writing. Errors: ${failed.join(" | ")}`);
    process.exit(1);
  }
  const totalTrades = Object.values(companies).reduce((a, c) => a + (c.transactions || 0) + (c.pledges_excluded || 0), 0);
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    source: "muns.io filings/insider_trades (BSE + Trendlyne) — rolling merge",
    lookback_days: LOOKBACK_DAYS,
    total_trades: totalTrades,
    companies,
  }) + "\n");
  console.log(`\nUpdated (${updated.length}): ${updated.join("  |  ") || "none"}`);
  if (emptied.length) console.log(`No in-window disclosures (${emptied.length}): ${emptied.join(", ")}`);
  if (failed.length) console.log(`Failed (${failed.length}): ${failed.join("  |  ")}`);
  console.log(`Wrote ${OUT_PATH} — ${Object.keys(companies).length} companies total, ${totalTrades} trades.`);
}

async function run() {
  if (FIXTURE_PATH) { runFixture(); return; }

  const hasToken = !!process.env.MUNS_TOKEN;
  console.log(`MUNS_TOKEN present: ${hasToken} (length ${(process.env.MUNS_TOKEN || "").length})`);
  if (!hasToken) {
    // Don't clobber the last-good file — just fail the step loudly.
    console.error("No MUNS_TOKEN in environment — cannot call the insider API. Leaving insider-trades.json untouched.");
    process.exit(1);
  }

  if (PROBE_TICKER) { await runProbe(); return; }
  if (TICKERS_ENV.length) { await runBatch(TICKERS_ENV); return; }

  const tickers = loadTickers();
  console.log(`Universe: ${tickers.length} tickers · country=${COUNTRY} · lookback=${LOOKBACK_DAYS}d · concurrency=${CONCURRENCY}`);

  const { results, stats } = await runPool(tickers);

  const cutoff = new Date(Date.now() - LOOKBACK_DAYS * 86400000);
  const companies = {};
  let totalTrades = 0, withData = 0, undatedTotal = 0, inWindowTotal = 0;
  for (const [t, rows] of Object.entries(results)) {
    if (!rows.length) continue;
    const { entry, inWindow, undated } = aggregateTicker(rows, cutoff);
    undatedTotal += undated;
    inWindowTotal += inWindow;
    // Keep a ticker only if it has a real classified disclosure. If a schema
    // drift made classify() return "other" for every row, transactions and
    // pledges_excluded stay 0 for every ticker → withData ends up 0 → the
    // guard below preserves the last-good file instead of overwriting the
    // whole universe with zero-trade entries.
    if (entry.transactions > 0 || entry.pledges_excluded > 0) {
      companies[t] = entry;
      withData++;
      totalTrades += entry.transactions + entry.pledges_excluded;
    }
  }

  console.log(`\nTickers queried: ${tickers.length} — ok ${stats.ok}, empty ${stats.empty}, failed ${stats.fail} (auth ${stats.auth})`);
  if (stats.errors.length) { console.log("Sample errors:"); stats.errors.forEach((e) => console.log(`  - ${e}`)); }
  if (undatedTotal) console.log(`Skipped ${undatedTotal} undated rows across all tickers.`);
  console.log(`Companies with in-window disclosures: ${withData}`);

  if (withData === 0) {
    // Whole run produced nothing usable (transient outage / schema drift).
    // Preserve the last-good file rather than overwrite it with zeros, which
    // would flip the Insider rule to N/A across the entire universe.
    const hint = inWindowTotal > 0
      ? `Saw ${inWindowTotal} in-window rows but classified none as Buy/Sell/Pledge — the API's transaction labels probably changed. Run with MUNS_PROBE_TICKER=RELIANCE and inspect insider-debug.json.`
      : "No in-window rows returned at all — check the response envelope, date parsing, or auth.";
    console.error(`No usable insider data — preserving last-good insider-trades.json and failing the step.\n  ${hint}`);
    process.exit(1);
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: "muns.io filings/insider_trades (BSE + Trendlyne)",
    lookback_days: LOOKBACK_DAYS,
    total_trades: totalTrades,
    tickers_ok: stats.ok,
    tickers_total: tickers.length,
    companies,
  };
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload) + "\n");
  console.log(`\nWrote ${OUT_PATH} — ${withData} companies, ${totalTrades} trades.`);
}

run().catch((err) => {
  // Never overwrite the committed file on an unexpected error — last-good wins.
  console.error("Fatal:", err.stack || err.message);
  process.exit(1);
});
