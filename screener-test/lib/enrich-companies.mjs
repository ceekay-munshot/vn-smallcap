// Shared company-enrichment step. The dashboard's live SPIP Basket tab
// (public/js/app.js) fuses several auxiliary data sources onto each
// company BEFORE running the composite scorer — auditor opinions,
// governance flags, insider trades, per-company revenue mix, macro
// sector overlays (PLI / renewable / China+1) and ATR history. The
// snapshot writers used to skip all of that and score on raw screener
// data, so a snapshot's composite / rating / hard-fail could differ from
// what the live tab showed for the same day.
//
// This module reproduces app.js's enrichment 1:1 for Node, so
// write-snapshot.mjs / backfill-snapshots.mjs bake the SAME numbers the
// live tab renders. Keep this in lockstep with app.js loadTab() (the
// composite branch) — if the live enrichment changes, mirror it here.

import { readFileSync } from "node:fs";
import { resolve } from "node:path";

function readJsonSafe(dir, name) {
  try { return JSON.parse(readFileSync(resolve(dir, name), "utf8")); }
  catch { return null; }
}

const tickerFromUrl = (url) => {
  const m = String(url || "").match(/\/company\/([^/]+)/);
  return m ? m[1].toUpperCase() : null;
};

// Mirror of app.js loadInsiderMerged(): base insider-trades.json, then
// fold in the recent supplement (summing aggregates, keeping latest date).
function loadInsiderMerged(dir) {
  const out = {};
  const base = readJsonSafe(dir, "insider-trades.json");
  if (base?.companies) Object.assign(out, base.companies);
  const supp = readJsonSafe(dir, "insider-recent-supplement.json");
  if (!supp?.companies) return out;
  for (const [ticker, s] of Object.entries(supp.companies)) {
    const cur = out[ticker] || {
      buy_shares: 0, sell_shares: 0, buy_value: 0, sell_value: 0,
      transactions: 0, pledges_excluded: 0, last_date: null,
      net_shares: 0, net_value: 0,
    };
    cur.buy_shares       = (cur.buy_shares       || 0) + (s.buy_shares       || 0);
    cur.sell_shares      = (cur.sell_shares      || 0) + (s.sell_shares      || 0);
    cur.buy_value        = (cur.buy_value        || 0) + (s.buy_value        || 0);
    cur.sell_value       = (cur.sell_value       || 0) + (s.sell_value       || 0);
    cur.transactions     = (cur.transactions     || 0) + (s.transactions     || 0);
    cur.pledges_excluded = (cur.pledges_excluded || 0) + (s.pledges_excluded || 0);
    cur.net_shares       = (cur.buy_shares || 0) - (cur.sell_shares || 0);
    cur.net_value        = (cur.buy_value  || 0) - (cur.sell_value  || 0);
    if (s.last_date && (!cur.last_date || new Date(s.last_date) > new Date(cur.last_date))) {
      cur.last_date = s.last_date;
    }
    out[ticker] = cur;
  }
  return out;
}

// Mutates fundCompanies / techCompanies in place, attaching the same
// enrichment fields the live composite tab sets. `dataDir` is the folder
// holding the auxiliary JSON (normally public/data). `macro` is the
// already-parsed macro.json (its overlay lists drive in_pli / in_renewable
// / in_china_plus_one).
export function enrichCompanies(fundCompanies, techCompanies, macro, dataDir) {
  const insiderByTicker = loadInsiderMerged(dataDir);
  const insiderLoaded   = Object.keys(insiderByTicker).length > 0;

  const revenueMixByTicker = readJsonSafe(dataDir, "company-revenue-mix.json")?.companies || {};

  const govJson = readJsonSafe(dataDir, "governance-flags.json");
  const governanceByTicker = govJson?.flagged_companies || {};
  const governanceLoaded   = !!govJson && !govJson.error;

  const audJson = readJsonSafe(dataDir, "auditor-opinions.json");
  const auditorByTicker = audJson?.companies || {};
  const auditorLoaded   = !!audJson && Object.keys(auditorByTicker).length > 0;

  const pli   = new Set((macro?.pli_companies || []).map((s) => String(s).toUpperCase()));
  const renew = new Set((macro?.renewable_companies || []).map((s) => String(s).toUpperCase()));
  const cp1   = new Set((macro?.china_plus_one_companies || []).map((s) => String(s).toUpperCase()));

  for (const row of fundCompanies) {
    const ticker = tickerFromUrl(row["Screener URL"]);
    // Insider (Fundamentals enrichment)
    const ins = ticker ? insiderByTicker[ticker] : null;
    row.insider_loaded = insiderLoaded;
    if (ins) {
      row.insider_net_shares = ins.net_shares; row.insider_net_value = ins.net_value;
      row.insider_buy_shares = ins.buy_shares; row.insider_sell_shares = ins.sell_shares;
      row.insider_transactions = ins.transactions; row.insider_last_date = ins.last_date;
      row.insider_pledges_excluded = ins.pledges_excluded || 0;
    } else { row.insider_transactions = 0; }
    // Governance
    row.governance_loaded = governanceLoaded;
    row.governance_flag = ticker && governanceByTicker[ticker] ? governanceByTicker[ticker] : null;
    // Auditor opinion
    row.auditor_opinions_loaded = auditorLoaded;
    const aud = ticker ? auditorByTicker[ticker] : null;
    row.auditor_opinion = aud?.opinion || null;
    row.auditor_opinion_source = aud?.source || null;
    row.auditor_firm = aud?.auditor_firm || null;
    row.auditor_opinion_date = aud?.auditor_opinion_date || null;
    row.auditor_report_year = aud?.auditor_report_year || null;
    row.auditor_emphasis_of_matter = aud?.auditor_emphasis_of_matter || null;
    row.auditor_key_concerns = aud?.auditor_key_concerns || null;
    row.auditor_confidence = aud?.auditor_confidence || null;
    // Macro sector overlays
    row.in_pli = ticker ? pli.has(ticker) : false;
    row.in_renewable = ticker ? renew.has(ticker) : false;
    row.in_china_plus_one = ticker ? cp1.has(ticker) : false;
    // Per-company revenue-mix extraction (truth from annual report)
    row._revenue_mix = ticker ? (revenueMixByTicker[ticker] || null) : null;
  }

  // ATR history for technicals (keyed the same way app.js keys it).
  const atrHistory = readJsonSafe(dataDir, "atr-history.json");
  if (atrHistory) {
    for (const row of techCompanies) {
      if (row.ticker && atrHistory[row.ticker]) row.atr_history = atrHistory[row.ticker];
    }
  }

  return {
    insiderLoaded, governanceLoaded, auditorLoaded,
    counts: {
      insider: Object.keys(insiderByTicker).length,
      governance: Object.keys(governanceByTicker).length,
      auditor: Object.keys(auditorByTicker).length,
      revenueMix: Object.keys(revenueMixByTicker).length,
    },
  };
}
