#!/usr/bin/env node
// Macro scraper. Fetches the few daily-changing market values (crude oil,
// USD/INR, India VIX) from Yahoo Finance, computes a short trend window for
// each, then merges with the slow-changing inputs in screener-test/static/
// macro-context.json and writes public/data/macro.json. The dashboard's
// Macro tab reads that single file.

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const STATIC_PATH       = resolve(__dirname, "static/macro-context.json");
const OUT_PATH          = resolve(__dirname, "../public/data/macro.json");
const TECH_PATH         = resolve(__dirname, "../public/data/technicals.json");
const FII_HISTORY_PATH  = resolve(__dirname, "../public/data/fii-dii-history.json");

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

run().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  // Write a best-effort payload using just the static data so the
  // dashboard still has something to score with.
  writeStaticOnly(err.message);
  process.exit(0);
});

async function run() {
  console.log("Loading static macro context...");
  const stat = JSON.parse(readFileSync(STATIC_PATH, "utf8"));

  console.log("Fetching daily market values from Yahoo Finance...");
  const [crude, inrusd, vix] = await Promise.all([
    fetchYahoo("BZ=F", 35),         // Brent crude
    fetchYahoo("USDINR=X", 35),     // USD/INR
    fetchYahoo("^INDIAVIX", 35),    // India VIX (used by Sentiment tab)
  ]);

  const live = {
    crude_brent: trendSummary(crude, "$/bbl"),
    usdinr:      trendSummary(inrusd, "₹/$"),
    india_vix:   trendSummary(vix, ""),
  };
  console.log("  Crude Brent:", live.crude_brent.latest, "(30d change:", live.crude_brent.pct_change_30d + "%)");
  console.log("  USD/INR:    ", live.usdinr.latest, "(30d change:", live.usdinr.pct_change_30d + "%)");
  console.log("  India VIX:  ", live.india_vix.latest);

  // 10Y G-Sec yield — try a small chain of public sources. If all fail we
  // silently fall through to the static `gsec10y` value baked into
  // macro-context.json (so the dashboard never goes blank).
  console.log("\nFetching India 10Y G-Sec yield...");
  const gsec = await fetchGSec10Y();
  if (gsec) {
    live.gsec10y = gsec;
    console.log(`  10Y yield: ${gsec.latest}% (${gsec.trend}) via ${gsec.source}`);
  } else {
    console.log("  10Y yield: live fetch failed — falling back to static macro-context.json value.");
  }

  // Live NSE sentiment fetches — FII/DII flow, PCR, A/D ratio. Each is
  // best-effort: if NSE blocks our IP we fall back to the static value
  // (or null) and the Sentiment tab rule shows the right N/A note.
  console.log("\nInitialising NSE session for sentiment fetches...");
  let liveSentiment = {};
  try {
    await initNSESession();
    liveSentiment = await fetchNSESentimentAll();
    console.log("  FII flow signal:", liveSentiment.fii_net_positive_last_20d, `(${liveSentiment.fii_positive_days}/${liveSentiment.fii_total_days} days)`);
    console.log("  Put/Call Ratio:", liveSentiment.put_call_ratio);
    console.log("  A/D ratio:     ", liveSentiment.market_breadth_ad_ratio, "(", liveSentiment.market_breadth_note, ")");
  } catch (err) {
    console.log("  NSE sentiment fetch failed:", err.message, "— falling back to static values.");
  }

  // Merge: prefer live values where present, fall back to static
  const sentiment = { ...(stat.sentiment || {}), ...liveSentiment };

  // Economic block — overlay live 10Y yield on top of static values if we got it.
  const economic = { ...stat.economic };
  if (live.gsec10y) {
    economic.gsec10y = live.gsec10y.latest;
    economic.gsec10y_trend = live.gsec10y.trend;
    economic.gsec10y_source = live.gsec10y.source;
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source: "Yahoo Finance (live) + NSE (live sentiment) + macro-context.json (slow-changing)",
    live,
    economic,
    regime: stat.regime,
    sentiment,
    sector_themes: stat.sector_themes,
    pli_companies: stat.pli_companies,
    renewable_companies: stat.renewable_companies,
    china_plus_one_companies: stat.china_plus_one_companies || [],
    static_last_updated: stat._last_manual_update || null,
  };

  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\nWrote: ${OUT_PATH}`);
  console.log(`PLI companies: ${stat.pli_companies.length}, Renewable: ${stat.renewable_companies.length}`);
}

function writeStaticOnly(reason) {
  try {
    const stat = JSON.parse(readFileSync(STATIC_PATH, "utf8"));
    const payload = {
      generated_at: new Date().toISOString(),
      source: "macro-context.json (live fetch failed)",
      live: null,
      live_error: reason,
      economic: stat.economic,
      regime: stat.regime,
      sentiment: stat.sentiment || null,
      sector_themes: stat.sector_themes,
      pli_companies: stat.pli_companies,
      renewable_companies: stat.renewable_companies,
      china_plus_one_companies: stat.china_plus_one_companies || [],
      static_last_updated: stat._last_manual_update || null,
    };
    mkdirSync(dirname(OUT_PATH), { recursive: true });
    writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
    console.log(`Wrote static-only payload: ${OUT_PATH}`);
  } catch (e) {
    console.error("Couldn't even write static-only:", e.message);
  }
}

// India 10Y G-Sec yield — public sources are unreliable from CI IPs, so we
// try a small chain and accept whichever responds first. If all fail we
// return null and the static value in macro-context.json wins.
async function fetchGSec10Y() {
  // Path 1: Yahoo Finance under any of the symbols the community has used
  // for India 10Y bond yield. None of these is officially documented; we
  // probe and use whichever returns a non-empty close series.
  const yahooSymbols = ["INDGB10Y=X", "^IN10Y", "^IN10YT", "IN10YT-RR", "^IND10Y"];
  for (const sym of yahooSymbols) {
    try {
      const closes = await fetchYahoo(sym, 35);
      if (closes && closes.length >= 5) {
        const summary = trendSummary(closes, "%");
        return { latest: summary.latest, trend: summary.trend, history_bars: closes.length, source: `Yahoo (${sym})` };
      }
    } catch { /* try next */ }
  }

  // Path 2: scrape worldgovernmentbonds.com — simple static HTML, no JS.
  // The 10Y row sits in a table; we look for the yield percentage adjacent
  // to the "10 years" anchor.
  try {
    const r = await fetch("https://www.worldgovernmentbonds.com/country/india/", {
      headers: {
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9",
      },
    });
    if (r.ok) {
      const html = await r.text();
      // The page renders a table row like:
      //   <tr><td>10 years</td><td>6.95%</td><td>...</td><td>+0.01</td>...</tr>
      // We grab the first percentage in the row containing "10 years".
      const rowMatch = html.match(/10\s*years?[\s\S]{0,400}?<\/tr>/i);
      if (rowMatch) {
        const pctMatch = rowMatch[0].match(/(\d+\.\d{1,3})\s*%/);
        const changeMatch = rowMatch[0].match(/([+-]?\d+\.\d{1,4})\s*<\/td>/);
        if (pctMatch) {
          const value = Number(pctMatch[1]);
          // The change column on WGB is a 1-day move, so we don't have a
          // robust 30d trend from a single fetch. Approximate trend from
          // sign of recent change column if present, else "stable".
          let trend = "stable";
          if (changeMatch) {
            const change = Number(changeMatch[1]);
            if (change > 0.05) trend = "rising";
            else if (change < -0.05) trend = "declining";
          }
          return { latest: value, trend, history_bars: 1, source: "worldgovernmentbonds.com" };
        }
      }
    }
  } catch { /* fall through */ }

  return null;
}

async function fetchYahoo(symbol, days) {
  const end = Math.floor(Date.now() / 1000);
  const start = end - days * 86400;
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?period1=${start}&period2=${end}&interval=1d&events=history`;
  const headers = { "User-Agent": "Mozilla/5.0 (compatible; KLPDashboardBot/1.0)" };
  let attempt = 0, lastErr;
  while (attempt < 3) {
    try {
      const r = await fetch(url, { headers });
      if (!r.ok) throw new Error(`${symbol} HTTP ${r.status}`);
      const j = await r.json();
      const result = j?.chart?.result?.[0];
      if (!result?.timestamp) throw new Error(`${symbol} no data`);
      const closes = result.indicators?.quote?.[0]?.close || [];
      return closes.filter((x) => x != null);
    } catch (err) {
      lastErr = err;
      attempt++;
      if (attempt < 3) await sleep(800 * attempt);
    }
  }
  throw lastErr;
}

function trendSummary(closes, unit) {
  if (!closes.length) return { latest: null };
  const latest = round(closes.at(-1), 2);
  const back30 = closes.length >= 22 ? closes.at(-22) : closes[0];   // ~30 calendar days ≈ 22 trading days
  const pct = back30 ? round(((closes.at(-1) - back30) / back30) * 100, 2) : 0;
  let trend;
  if (pct > 2) trend = "rising";
  else if (pct < -2) trend = "falling";
  else trend = "stable";
  return { latest, latest_unit: unit, pct_change_30d: pct, trend, history_bars: closes.length };
}

function round(n, d = 2) {
  if (n == null || !Number.isFinite(n)) return null;
  const m = Math.pow(10, d);
  return Math.round(n * m) / m;
}

// ---------- NSE session + sentiment fetches ----------

const NSE_HEADERS = {
  "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
  "Accept": "application/json, text/plain, */*",
  "Accept-Language": "en-US,en;q=0.9",
  "Referer": "https://www.nseindia.com/",
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
  for (const url of ["https://www.nseindia.com/", "https://www.nseindia.com/market-data/live-equity-market"]) {
    const r = await fetch(url, { headers: NSE_HEADERS });
    saveCookies(r);
    await sleep(600);
  }
}

async function fetchNSEJson(url, attempts = 3) {
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try {
      const r = await fetch(url, { headers: { ...NSE_HEADERS, Cookie: cookieJar } });
      if (r.status === 401 || r.status === 403) {
        await initNSESession();
        continue;
      }
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      saveCookies(r);
      return await r.json();
    } catch (err) {
      lastErr = err;
      await sleep(800 * (i + 1));
    }
  }
  throw lastErr;
}

async function fetchNSESentimentAll() {
  const out = {};

  // 1) FII/DII: accumulator approach — NSE's /api/fiidiiTradeReact returns
  //    today's snapshot only. We append it to a committed history file and
  //    compute the rolling signal from that history.
  try {
    const todayRows = await fetchFIIDIITodayRows();
    const history = appendToFIIHistory(todayRows);
    const signal = computeFIISignal(history);
    // Always assign — computeFIISignal now returns an explicit
    // "accumulating" state (null signal) when there's too little
    // history, which correctly overrides any optimistic seed value.
    Object.assign(out, signal);
    if (signal.fii_net_positive_last_20d != null) {
      console.log(`  FII flow: ${signal.fii_net_positive_last_20d} (${signal.fii_positive_days}/${signal.fii_total_days} days) — history file has ${history.length} entries`);
    } else {
      console.log(`  FII flow: accumulating — ${signal.fii_total_days}/20 days collected (${signal.fii_positive_days} net-positive); signal goes live at 5 days.`);
    }
  } catch (err) {
    console.log("  FII fetch error:", err.message);
  }

  // 2) Market breadth: read from the just-committed technicals.json. Our own
  //    Nifty 500 universe — no NSE breadth API needed.
  try {
    const tech = JSON.parse(readFileSync(TECH_PATH, "utf8"));
    if (tech.market_breadth && tech.market_breadth.ad_ratio != null) {
      out.market_breadth_ad_ratio = tech.market_breadth.ad_ratio;
      out.market_breadth_note = `${tech.market_breadth.advances} advances vs ${tech.market_breadth.declines} declines across ${tech.market_breadth.universe} Nifty 500 stocks (computed from Yahoo OHLCV).`;
      console.log(`  A/D ratio: ${tech.market_breadth.ad_ratio} (${tech.market_breadth.advances}adv / ${tech.market_breadth.declines}dec)`);
    } else {
      console.log("  A/D ratio: technicals.json has no market_breadth block yet — rerun Technicals scrape.");
    }
  } catch (err) {
    console.log("  A/D fetch error: couldn't read technicals.json —", err.message);
  }

  // 3) PCR: permanently deferred. We attempted 5 sources (NSE option-chain
  //    legacy, NSE option-chain v3, Yahoo options, MoneyControl scrape, NSE
  //    F&O bhavcopy in 4 URL formats) and none returned usable data from a
  //    GitHub Actions runner IP. NSE is the only definitive source and they
  //    actively block automated runners. Worth 1 pt — accepting deferred is
  //    the right call. Rule shows clear N/A note in the dashboard.
  return out;
}

async function fetchFIIDIITodayRows() {
  const data = await fetchNSEJson("https://www.nseindia.com/api/fiidiiTradeReact");
  if (!Array.isArray(data) || !data.length) throw new Error("empty response");
  return data;
}

function appendToFIIHistory(todayRows) {
  let history = [];
  try {
    history = JSON.parse(readFileSync(FII_HISTORY_PATH, "utf8"));
    if (!Array.isArray(history)) history = [];
  } catch { /* file doesn't exist yet — start fresh */ }

  // Append today's rows if not already in history (key by date+category)
  const seen = new Set(history.map((r) => `${r.date}::${r.category}`));
  for (const r of todayRows) {
    const key = `${r.date}::${r.category}`;
    if (!seen.has(key)) {
      history.push(r);
      seen.add(key);
    }
  }
  // Sort newest-first by date and trim to ~60 days to keep the file bounded
  history.sort((a, b) => new Date(b.date) - new Date(a.date));
  history = history.slice(0, 120);   // up to ~60 days × 2 (FII + DII)

  writeFileSync(FII_HISTORY_PATH, JSON.stringify(history, null, 2) + "\n");
  return history;
}

function computeFIISignal(history) {
  const fiiRows = history.filter((r) => {
    const cat = String(r.category || "").toLowerCase();
    return cat.includes("fii") || cat.includes("fpi") || cat.includes("foreign");
  });
  // Group by date (one row per day) and aggregate netValue
  const byDate = new Map();
  for (const r of fiiRows) {
    const d = r.date;
    const v = Number(r.netValue ?? r.net ?? 0);
    byDate.set(d, (byDate.get(d) || 0) + v);
  }
  // Take latest up to 20 days
  const dates = [...byDate.keys()].sort((a, b) => new Date(b) - new Date(a)).slice(0, 20);
  const days = dates.length;
  const positive = dates.filter((d) => byDate.get(d) > 0).length;
  // The rolling signal needs at least 5 trading days before it means
  // anything. Below that, return an explicit "accumulating" state with
  // a null signal (scored N/A by the rule) rather than leaving a stale
  // seed in place — never assert a flow direction we can't yet compute.
  if (days < 5) {
    return {
      fii_net_positive_last_20d: null,
      fii_positive_days: positive,
      fii_total_days: days,
      fii_signal_note: `Building rolling window — ${days} of 20 trading days collected so far (${positive} net-positive). NSE only publishes a daily FII/DII snapshot, so the signal accumulates day-by-day and goes live once ≥ 5 trading days are gathered, then fills toward a true 20-day window.`,
    };
  }
  const pct = positive / days;
  let signal;
  if (pct >= 0.5) signal = "yes";
  else if (pct >= 0.3) signal = "mixed";
  else signal = "no";
  return {
    fii_net_positive_last_20d: signal,
    fii_positive_days: positive,
    fii_total_days: days,
    fii_signal_note: `FII net positive in ${positive} of last ${days} reported trading days (auto-computed from rolling daily snapshots).`,
  };
}

