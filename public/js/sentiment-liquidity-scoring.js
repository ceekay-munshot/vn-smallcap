// Sentiment & Liquidity scoring — rules verbatim from the client's checklist.
// Sentiment rules are MARKET-WIDE (every company gets the same score that day).
// Liquidity rules are PER-COMPANY (ADTV / F&O eligibility).

const NA = { points: 0, max: 0, status: "na", value: null, note: "Data not available" };

// ---- Sentiment (market-wide) ----

function ruleIndiaVIX(c) {
  const vix = c?._macro?.live?.india_vix?.latest;
  if (vix == null) return { ...NA, max: 2, note: "Live India VIX not yet fetched — run the Macro workflow." };
  const val = `India VIX ${vix}`;
  if (vix < 15) return { points: 2, max: 2, status: "pass", value: val, note: "VIX < 15 — low fear, equity accumulation mode." };
  if (vix <= 20) return { points: 1, max: 2, status: "partial", value: val, note: "VIX 15–20 — caution zone (neither risk-on nor panic)." };
  return { points: 0, max: 2, status: "fail", value: val, note: "VIX > 20 — reduce exposure." };
}

function ruleFIIDIIFlow(c) {
  const s = c?._macro?.sentiment;
  if (!s) return { ...NA, max: 2 };
  if (s.fii_net_positive_last_20d == null) return { ...NA, max: 2, value: s.fii_total_days != null ? `Accumulating — ${s.fii_total_days}/20 days` : null, note: s.fii_signal_note || "FII/DII flow signal not yet available — accumulating daily snapshots (goes live at 5 trading days)." };
  const signal = String(s.fii_net_positive_last_20d).toLowerCase();
  const note = s.fii_signal_note || "FII/DII flow signal from market activity report.";
  if (signal === "yes") return { points: 2, max: 2, status: "pass", value: "FII net positive in 10+ of last 20 days", note };
  if (signal === "mixed") return { points: 1, max: 2, status: "partial", value: "FII flow mixed", note };
  return { points: 0, max: 2, status: "fail", value: "Consistent FII selling", note };
}

function rulePutCallRatio(c) {
  const s = c?._macro?.sentiment;
  if (!s || s.put_call_ratio == null) {
    return { ...NA, max: 1,
      note: "Put-Call Ratio not available today. Sourced via Firecrawl LLM extract from NSE option-chain. Rule will auto-recover on the next successful sentiment-extras scrape.",
    };
  }
  const pcr = s.put_call_ratio;
  const stale = s.put_call_ratio_stale ? " (yesterday's value — today's fetch failed)" : "";
  const val = `PCR ${pcr}${stale}`;
  // Strict client logic: pass only inside 0.9-1.3 band. Everything else
  // is 0 pts; the note distinguishes "overly bullish caution" (< 0.8),
  // "panic" (> 1.5), and the transitional zones that the spec doesn't
  // name explicitly.
  if (pcr >= 0.9 && pcr <= 1.3) {
    return { points: 1, max: 1, status: "pass", value: val,
      note: "PCR within 0.9–1.3 (balanced to mildly bullish) — full credit per client framework." };
  }
  if (pcr < 0.8) {
    return { points: 0, max: 1, status: "fail", value: val,
      note: "PCR < 0.8 — overly bullish positioning, caution flag per client framework." };
  }
  if (pcr > 1.5) {
    return { points: 0, max: 1, status: "fail", value: val,
      note: "PCR > 1.5 — panic levels per client framework." };
  }
  // 0.8 ≤ PCR < 0.9 or 1.3 < PCR ≤ 1.5 — transitional zones the spec
  // doesn't define separately. Default to fail (0 pts).
  const zone = pcr < 0.9 ? "0.8–0.9 (between caution and balanced)" : "1.3–1.5 (between balanced and panic)";
  return { points: 0, max: 1, status: "fail", value: val,
    note: `PCR in transitional zone ${zone} — outside the 0.9–1.3 pass band, no points awarded.` };
}

function ruleMarketBreadth(c) {
  const s = c?._macro?.sentiment;
  if (!s || s.market_breadth_ad_ratio == null) return { ...NA, max: 1, note: s?.market_breadth_note || "Live A/D breadth scrape not yet wired." };
  const ad = s.market_breadth_ad_ratio;
  const val = `A/D ratio ${ad}`;
  if (ad > 1.5) return { points: 1, max: 1, status: "pass", value: val, note: "A/D > 1.5 — broad participation, healthy breadth." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Narrow breadth — only large caps participating." };
}

// ---- Liquidity (per-company) ----

function ruleADTV(c) {
  const adtv = c?.adtv_20d_cr;
  if (adtv == null) return { ...NA, max: 2, note: "20-day ADTV not yet computed — needs OHLCV from the Technicals scrape." };
  const val = `ADTV ₹${adtv} Cr (20-day average)`;
  if (adtv >= 10) return { points: 2, max: 2, status: "pass", value: val, note: "ADTV ≥ ₹10 Cr — comfortable liquidity for institutional positions." };
  if (adtv >= 5) return { points: 1, max: 2, status: "partial", value: val, note: "ADTV ₹5–10 Cr — moderate liquidity, fine for retail/mid-AUM portfolios." };
  return { points: 0, max: 2, status: "hard_fail", value: val, note: "ADTV < ₹5 Cr — illiquid, hard fail per client framework. Stock excluded from SPIP basket." };
}

function ruleImpactCost(c) {
  // Prefer a real, vendor-sourced impact cost if/when it ever populates
  // (e.g. a future Upstox / TrueData integration); fall back to the
  // Amihud-derived estimate we compute from daily OHLCV at ₹5 Cr.
  const real = c?.impact_cost_pct;
  const est  = c?.impact_cost_pct_est_5cr;
  const ic   = (typeof real === "number" && real >= 0) ? real : est;
  if (ic == null) return { ...NA, max: 2,
    note: "Impact cost not yet computable — Technicals scrape needs ≥ 30 days of OHLCV bars + non-zero volume." };
  const isEst = !(typeof real === "number" && real >= 0);
  const label = isEst ? `Impact cost ${ic}% (est., ₹5 Cr order)` : `Impact cost ${ic}%`;
  const tail  = isEst
    ? " Estimated via Amihud illiquidity ratio over last 30 trading days, modelled at a ₹5 Cr institutional position size (NSE's official monthly Impact Cost CSV was discontinued July 2024)."
    : "";
  if (ic <= 0.3) return { points: 2, max: 2, status: "pass",    value: label, note: "Impact cost ≤ 0.3% — tight execution for mid-cap orders." + tail };
  if (ic <= 0.5) return { points: 1, max: 2, status: "partial", value: label, note: "Impact cost 0.3–0.5% — acceptable slippage." + tail };
  return            { points: 0, max: 2, status: "fail",    value: label, note: "Impact cost > 0.5% — high slippage risk." + tail };
}

function ruleBidAskSpread(c) {
  // Prefer a real bid/ask spread if any source ever provides one;
  // fall back to the Abdi-Ranaldo CHL estimate from daily OHLCV.
  const real = c?.bid_ask_spread_pct;
  const est  = c?.bid_ask_spread_pct_est;
  const sp   = (typeof real === "number" && real >= 0) ? real : est;
  if (sp == null) return { ...NA, max: 1,
    note: "Bid-ask spread not yet computable — Technicals scrape needs ≥ 30 days of high/low/close bars." };
  const isEst = !(typeof real === "number" && real >= 0);
  const label = isEst ? `Bid-ask spread ${sp}% of CMP (est.)` : `Bid-ask spread ${sp}% of CMP`;
  const tail  = isEst
    ? " Estimated via Abdi-Ranaldo (2017) close/high/low proxy — daily OHLCV stand-in for NSE Level-1 quotes, which Yahoo doesn't carry."
    : "";
  if (sp < 0.1) return { points: 1, max: 1, status: "pass", value: label, note: "Spread < 0.1% — tight order book, full credit per client framework." + tail };
  if (sp > 0.3) return { points: 0, max: 1, status: "fail", value: label, note: "Spread > 0.3% — thin order book per client framework." + tail };
  return            { points: 0, max: 1, status: "fail", value: label, note: `Spread ${sp}% in transitional zone (0.1–0.3%) — outside the < 0.1% pass band.` + tail };
}

function ruleFnOAvailability(c) {
  if (c?.fno_eligible == null) return { ...NA, max: 1 };
  if (c.fno_eligible) return { points: 1, max: 1, status: "pass", value: "On NSE F&O list", note: "Active futures + options contracts — hedging flexibility, institutional interest." };
  return { points: 0, max: 1, status: "fail", value: "Not on NSE F&O list", note: "Cash-only stock. Preferred but not mandatory per client framework." };
}

// ---- master ----

// All 8 client-framework rules are active. PCR sourced via Firecrawl
// LLM extract from NSE option-chain. Bid-Ask Spread + Impact Cost are
// computed from daily OHLCV using validated academic estimators
// (Abdi-Ranaldo for spread, Amihud for impact) — NSE Level-1 quotes
// and the official Impact Cost CSV are both behind paywalls / were
// discontinued, and Yahoo doesn't carry NSE bid/ask.
const ACTIVE_RULES = [
  { key: "vix",      label: "India VIX",         category: "Sentiment", criteria: "< 15 (risk-on)",          fn: ruleIndiaVIX },
  { key: "fiidii",   label: "FII / DII Flow",    category: "Sentiment", criteria: "Net positive 10/20 days", fn: ruleFIIDIIFlow },
  { key: "pcr",      label: "Put Call Ratio",    category: "Sentiment", criteria: "0.9 – 1.3",               fn: rulePutCallRatio },
  { key: "breadth",  label: "Market Breadth",    category: "Sentiment", criteria: "A/D > 1.5",               fn: ruleMarketBreadth },
  { key: "adtv",     label: "Avg Daily Traded Value", category: "Liquidity", criteria: "≥ ₹10 Cr (20-day)", fn: ruleADTV },
  { key: "impact",   label: "Impact Cost",       category: "Liquidity", criteria: "≤ 0.3% for ₹5 Cr",        fn: ruleImpactCost },
  { key: "spread",   label: "Bid-Ask Spread",    category: "Liquidity", criteria: "< 0.1% of CMP",           fn: ruleBidAskSpread },
  { key: "fno",      label: "F&O Availability",  category: "Liquidity", criteria: "On NSE F&O list",         fn: ruleFnOAvailability },
];

const DEFERRED = [];

export function scoreCompany(c) {
  if (c?.error) {
    return {
      company: c, breakdown: [], deferred: DEFERRED,
      totalPoints: 0, totalMax: 0, scorePct: 0,
      hardFails: [], naCount: ACTIVE_RULES.length,
      tickerError: c.error,
    };
  }
  const breakdown = ACTIVE_RULES.map((r) => ({ ...r, ...r.fn(c) }));
  const totalPoints = breakdown.reduce((s, b) => s + b.points, 0);
  const totalMax = breakdown.reduce((s, b) => s + b.max, 0);
  const naCount = breakdown.filter((b) => b.status === "na").length;
  const hardFails = breakdown.filter((b) => b.status === "hard_fail").map((b) => b.label);
  return {
    company: c, breakdown, deferred: DEFERRED,
    totalPoints, totalMax,
    scorePct: totalMax ? Math.round((totalPoints / totalMax) * 100) : 0,
    hardFails,
    naCount,
  };
}

export { ACTIVE_RULES, DEFERRED };
