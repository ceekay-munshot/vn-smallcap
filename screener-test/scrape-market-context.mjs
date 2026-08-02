#!/usr/bin/env node
// Pre-open market context. Answers one question before the bell: is the
// market likely to open sharply up or down, and how jumpy is it right now?
//
// GIFT Nifty would be the direct read, but there is no free, stable feed
// for it -- Yahoo has no symbol, Moneycontrol's price API has no key for
// it, and their page renders the number in JavaScript. Rather than build
// a cron on a scraper that breaks silently, this uses the inputs that
// actually drive GIFT Nifty overnight and ARE reliably fetchable:
//
//   S&P 500 futures   trade nearly 24h; the biggest overnight driver
//   Nikkei 225        Asian session, live while our pre-open runs
//   Hang Seng         same
//   India VIX         how jumpy the market itself is
//
// The output is deliberately advisory. It widens or tightens the entry
// tolerance on the Trade Plan and warns that limits may not fill on a
// gap-up morning. It does not pick stocks.
//
// If a GIFT Nifty source ever becomes available, add it to FEEDS and the
// dashboard picks it up -- the UI already prefers it when present.
//
// Writes: public/data/market-context.json

import { writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_PATH = resolve(__dirname, "../public/data/market-context.json");

const FEEDS = [
  { key: "spFutures", symbol: "ES=F",       label: "S&P 500 futures", role: "overnight" },
  { key: "nikkei",    symbol: "^N225",      label: "Nikkei 225",      role: "asia" },
  { key: "hangSeng",  symbol: "^HSI",       label: "Hang Seng",       role: "asia" },
  { key: "vix",       symbol: "^INDIAVIX",  label: "India VIX",       role: "vol" },
  { key: "nifty",     symbol: "^NSEI",      label: "Nifty 50",        role: "home" },
  { key: "smallcap",  symbol: "NIFTYSMLCAP250.NS", label: "Nifty Smallcap 250", role: "home" },
];

run().catch((e) => { console.error("Fatal:", e.stack || e.message); process.exit(1); });

async function run() {
  const feeds = {};
  for (const f of FEEDS) {
    process.stdout.write(`${f.label}… `);
    const q = await fetchQuote(f.symbol);
    if (q) { feeds[f.key] = { ...f, ...q }; console.log(`${q.last?.toFixed(2)} (${fmtPct(q.changePct)})`); }
    else console.log("FAILED");
    await new Promise((r) => setTimeout(r, 250));
  }

  const signal = buildSignal(feeds);
  const payload = {
    generated_at: new Date().toISOString(),
    source: "Yahoo Finance chart endpoint",
    note: "Pre-open context. GIFT Nifty has no free stable feed; S&P futures plus the Asian session carry the same overnight information.",
    giftNifty: null,          // slot reserved — fill if a source appears
    feeds,
    signal,
  };
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\n${signal.headline}\nWrote ${OUT_PATH}`);
}

// Overnight lean = S&P futures weighted double the Asian session, because
// that is roughly how our open actually tracks them. Deliberately coarse:
// this decides a tolerance band, not a trade.
function buildSignal(feeds) {
  const parts = [];
  const sp = feeds.spFutures?.changePct;
  const asia = [feeds.nikkei?.changePct, feeds.hangSeng?.changePct].filter((v) => v != null);
  if (sp != null) parts.push({ v: sp, w: 2 });
  for (const a of asia) parts.push({ v: a, w: 1 });
  const wsum = parts.reduce((a, p) => a + p.w, 0);
  const lean = wsum ? parts.reduce((a, p) => a + p.v * p.w, 0) / wsum : null;

  const vix = feeds.vix?.last ?? null;
  const vixChange = feeds.vix?.changePct ?? null;
  const vixBand = vix == null ? "unknown" : vix < 14 ? "calm" : vix < 20 ? "normal" : "elevated";

  // Entry tolerance: how far above yesterday's close we will still pay.
  // On a strongly positive overnight lean the whole market gaps, so a
  // limit pinned at yesterday's close throws away good names for no
  // reason. On a jumpy day we tighten instead.
  let tolerancePct = 0.5;
  if (lean != null) tolerancePct = Math.max(0.25, Math.min(1.5, 0.5 + lean * 0.6));
  if (vixBand === "elevated") tolerancePct = Math.min(tolerancePct, 0.5);

  const dir = lean == null ? "flat" : lean > 0.5 ? "up" : lean < -0.5 ? "down" : "flat";
  const headline =
    dir === "up"   ? "Overnight positive — expect a higher open, limits may not fill" :
    dir === "down" ? "Overnight negative — expect a lower open, limits should fill easily" :
                     "Overnight quiet — a normal open";

  return { lean, dir, vix, vixChange, vixBand, tolerancePct: +tolerancePct.toFixed(2), headline };
}

const fmtPct = (v) => (v == null ? "—" : (v >= 0 ? "+" : "") + v.toFixed(2) + "%");

async function fetchQuote(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?range=5d&interval=1d`;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const r = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (compatible; VNSmallcapBot/1.0)" } });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const res = (await r.json())?.chart?.result?.[0];
      const m = res?.meta;
      if (!m) throw new Error("no meta");
      const last = m.regularMarketPrice ?? null;
      const prev = m.chartPreviousClose ?? m.previousClose ?? null;
      return {
        last,
        prevClose: prev,
        changePct: last != null && prev ? (last / prev - 1) * 100 : null,
        asOf: m.regularMarketTime ? new Date(m.regularMarketTime * 1000).toISOString() : null,
      };
    } catch (e) {
      if (attempt === 2) return null;
      await new Promise((r) => setTimeout(r, 600 * (attempt + 1)));
    }
  }
  return null;
}
