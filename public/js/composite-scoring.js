// SPIP Composite Scoring — implements the client's weighted 5-pillar
// model from SPIP_Stock_Selection_Model_v1.xlsx (Scoring_Model sheet).
//
//   Composite = (Fund/29)*40 + (Tech/24)*35 + (Macro/17)*15
//             + (Sent/6)*5   + (Liq/6)*5
//   → 0-100 score
//
// Rating bands (Scoring_Model · Section C):
//   ≥ 75  STRONG BUY   — initiate full position
//   60-74 BUY          — initiate 50-75% position
//   45-59 WATCH / HOLD — watchlist, do not initiate
//   < 45  AVOID / EXIT — exclude from basket
//
// Hard-fail rule (Scoring_Model · Section A): if ANY of the client's
// hard fails trigger, the stock exits the pipeline entirely — its
// composite is shown but it is filtered out of the SPIP basket.
// The 5 hard fails per the spec:
//   - Fundamentals: D/E > 2, Auditor adverse, Active SEBI investigation
//                   (also CFO negative, ICR < 1.5, Pledge > 20%)
//   - Technicals:   Price < 200 DMA
//   - Liquidity:    ADTV < ₹5 Cr

import * as fund from "./scoring.js";
import * as tech from "./tech-scoring.js";
import * as macro from "./macro-scoring.js";
import * as senliq from "./sentiment-liquidity-scoring.js";

// Hard-fail rule labels that are TREND / TRADABILITY filters, not company
// red flags. A stock trips these because of where the market is right now
// (below its 200-day average) or how thinly it trades — both flip back as
// conditions change, so the UI shows them apart from fundamental red flags.
const FILTER_HARD_FAILS = ["Price Above 200 DMA", "Avg Daily Traded Value"];

export const PILLAR_WEIGHTS = {
  fundamentals: 40,
  technicals: 35,
  macro: 15,
  sentiment: 5,
  liquidity: 5,
};

export const PILLAR_MAX_RAW = {
  fundamentals: 29,
  technicals: 24,
  macro: 17,
  sentiment: 6,
  liquidity: 6,
};

export function ratingFromComposite(composite, hardFailed) {
  if (hardFailed) return "FILTERED";
  if (composite == null) return "UNRATED";
  if (composite >= 75) return "STRONG BUY";
  if (composite >= 60) return "BUY";
  if (composite >= 45) return "WATCH";
  return "AVOID";
}

export function ratingClass(rating) {
  switch (rating) {
    case "STRONG BUY": return "bg-emerald-100 text-emerald-800 ring-emerald-200";
    case "BUY":        return "bg-blue-100 text-blue-800 ring-blue-200";
    case "WATCH":      return "bg-amber-100 text-amber-800 ring-amber-200";
    case "AVOID":      return "bg-rose-100 text-rose-800 ring-rose-200";
    case "FILTERED":   return "bg-rose-50 text-rose-700 ring-rose-200";
    default:           return "bg-slate-100 text-slate-700 ring-slate-200";
  }
}

// Gradient + accent palette used by the SPIP Basket drill / hero cards.
// Each entry: { from, to, accent (text), ring, dot (solid color for chips) }
export function ratingTheme(rating) {
  switch (rating) {
    case "STRONG BUY": return { from: "from-emerald-500", to: "to-teal-500", soft: "from-emerald-50 to-teal-50", accent: "text-emerald-700", textOn: "text-white", ring: "ring-emerald-200", dot: "bg-emerald-500" };
    case "BUY":        return { from: "from-blue-500",    to: "to-indigo-500", soft: "from-blue-50 to-indigo-50", accent: "text-blue-700",    textOn: "text-white", ring: "ring-blue-200",    dot: "bg-blue-500" };
    case "WATCH":      return { from: "from-amber-500",   to: "to-orange-500", soft: "from-amber-50 to-orange-50", accent: "text-amber-700",   textOn: "text-white", ring: "ring-amber-200",   dot: "bg-amber-500" };
    case "AVOID":      return { from: "from-rose-500",    to: "to-pink-500", soft: "from-rose-50 to-pink-50", accent: "text-rose-700",    textOn: "text-white", ring: "ring-rose-200",    dot: "bg-rose-500" };
    case "FILTERED":   return { from: "from-slate-500",   to: "to-slate-600", soft: "from-slate-50 to-slate-100", accent: "text-slate-700",   textOn: "text-white", ring: "ring-slate-300",   dot: "bg-slate-500" };
    default:           return { from: "from-slate-400",   to: "to-slate-500", soft: "from-slate-50 to-slate-100", accent: "text-slate-700",   textOn: "text-white", ring: "ring-slate-200",   dot: "bg-slate-400" };
  }
}

// Decision-framework text taken verbatim from SPIP_Stock_Selection_Model_v1.xlsx
// Scoring_Model · Section C. The analyst-facing recommendation that pairs with
// each rating band — what to do, how big, how often to review, when to exit.
export function decisionFor(rating) {
  switch (rating) {
    case "STRONG BUY": return {
      verdict: "Include in SPIP Basket",
      action: "Initiate FULL position",
      size: "8–10% of basket",
      review: "Monthly",
      exit: "Score drops below 60 OR hard fail triggered",
      profile: "Passes all hard filters · strong fundamentals + technical momentum + macro tailwind",
    };
    case "BUY": return {
      verdict: "Include with monitoring",
      action: "Initiate 50–75% of target position",
      size: "5–7% of basket",
      review: "Bi-weekly",
      exit: "Score drops below 50 OR technical breakdown",
      profile: "Strong on 3 of 5 pillars · minor gaps in 1–2 parameters",
    };
    case "WATCH": return {
      verdict: "Watchlist · do not initiate",
      action: "No new position · hold if already invested",
      size: "3–5% max",
      review: "Weekly",
      exit: "Score does not improve within 2 months",
      profile: "Mixed signals · fundamentals OK but technicals or macro lagging",
    };
    case "AVOID": return {
      verdict: "Exclude from SPIP Basket",
      action: "Exit existing position · no fresh entry",
      size: "Zero allocation",
      review: "Immediate",
      exit: "Already triggered — exit now",
      profile: "Fails multiple pillar thresholds",
    };
    case "FILTERED": return {
      verdict: "Filtered out of pipeline",
      action: "Stock excluded · address red flag before reconsidering",
      size: "Zero allocation",
      review: "On red-flag resolution",
      exit: "Stock exits pipeline regardless of composite score",
      profile: "Hard-fail rule triggered (per client framework, Section A)",
    };
    default: return {
      verdict: "Unrated",
      action: "Data incomplete — composite cannot be computed",
      size: "—",
      review: "—",
      exit: "—",
      profile: "Technicals / sentiment / liquidity data missing for this ticker",
    };
  }
}

// Split the senliq breakdown into Sentiment vs Liquidity sub-pillars
// (they live in the same module but have separate weights per spec).
function splitSenliq(senliqResult) {
  const sent = senliqResult.breakdown.filter((b) => b.category === "Sentiment");
  const liq  = senliqResult.breakdown.filter((b) => b.category === "Liquidity");
  const sumP = (rs) => rs.reduce((s, r) => s + (r.points || 0), 0);
  const sumM = (rs) => rs.reduce((s, r) => s + (r.max    || 0), 0);
  return {
    sentiment: { points: sumP(sent), max: sumM(sent), breakdown: sent },
    liquidity: { points: sumP(liq),  max: sumM(liq),  breakdown: liq },
  };
}

// Weighted contribution of one pillar (handles 0/0 → 0 safely).
function weighted(points, max, weightPct) {
  if (!max) return 0;
  return Math.round((points / max) * weightPct * 100) / 100;
}

// Score a single company across all five pillars. Inputs:
//   fundCo  - row from screener-companies.json (Fundamentals + Macro)
//   techCo  - matching row from technicals.json (Technicals + Senliq)
//             may be null if Yahoo OHLCV was missing for this ticker
//   macroCtx - parsed macro.json (live VIX, sentiment, sector themes)
//   weights  - optional override of PILLAR_WEIGHTS (lets the SPIP basket
//              tab re-score with client-adjusted weights without changing
//              the framework constants)
export function scoreCompositeOne(fundCo, techCo, macroCtx, weights = PILLAR_WEIGHTS) {
  const fundResult  = fund.scoreCompany(fundCo);
  const macroResult = macro.scoreCompany({ ...fundCo, _macro: macroCtx });

  // Technicals + Senliq need techCo. If we don't have OHLCV, return
  // an explicit "data missing" state rather than zero-scoring those
  // pillars (which would unfairly penalise the stock).
  const hasTechData = techCo && techCo.cmp != null;
  const techResult   = hasTechData ? tech.scoreCompany(techCo) : null;
  const senliqResult = hasTechData ? senliq.scoreCompany({ ...techCo, _macro: macroCtx }) : null;
  const split        = senliqResult ? splitSenliq(senliqResult) : null;

  // Collect all hard fails from every pillar, then split them into two very
  // different kinds so the UI can stop lumping them together:
  //  - FUNDAMENTAL red flags = something wrong with the COMPANY (bad auditor,
  //    extreme debt, negative cash flow, pledge, SEBI) → genuine "avoid".
  //  - FILTER flags = a temporary market condition, not a company problem:
  //    below the 200-day trend, or too thinly traded. These flip back as the
  //    stock/market moves, so they shouldn't read as scary red flags.
  const allHardFails = [
    ...(fundResult.hardFails || []),
    ...(techResult?.hardFails || []),
    ...(senliqResult?.hardFails || []),
  ];
  const filterFlags = allHardFails.filter((l) => FILTER_HARD_FAILS.includes(l));
  const fundamentalFlags = allHardFails.filter((l) => !FILTER_HARD_FAILS.includes(l));
  const isRedFlag = fundamentalFlags.length > 0;

  // Weighted composite — only sum pillars that have data.
  const pillars = {
    fundamentals: {
      raw: fundResult.totalPoints, max: fundResult.totalMax,
      pct: fundResult.scorePct,
      weighted: weighted(fundResult.totalPoints, fundResult.totalMax, weights.fundamentals),
    },
    technicals: techResult ? {
      raw: techResult.totalPoints, max: techResult.totalMax,
      pct: techResult.scorePct,
      weighted: weighted(techResult.totalPoints, techResult.totalMax, weights.technicals),
    } : { raw: null, max: PILLAR_MAX_RAW.technicals, pct: null, weighted: null },
    macro: {
      raw: macroResult.totalPoints, max: macroResult.totalMax,
      pct: macroResult.scorePct,
      weighted: weighted(macroResult.totalPoints, macroResult.totalMax, weights.macro),
    },
    sentiment: split ? {
      raw: split.sentiment.points, max: split.sentiment.max,
      pct: split.sentiment.max ? Math.round((split.sentiment.points / split.sentiment.max) * 100) : 0,
      weighted: weighted(split.sentiment.points, split.sentiment.max, weights.sentiment),
    } : { raw: null, max: PILLAR_MAX_RAW.sentiment, pct: null, weighted: null },
    liquidity: split ? {
      raw: split.liquidity.points, max: split.liquidity.max,
      pct: split.liquidity.max ? Math.round((split.liquidity.points / split.liquidity.max) * 100) : 0,
      weighted: weighted(split.liquidity.points, split.liquidity.max, weights.liquidity),
    } : { raw: null, max: PILLAR_MAX_RAW.liquidity, pct: null, weighted: null },
  };

  // If technicals/senliq are missing we cannot compute a faithful
  // composite — return null to indicate "unrated" rather than a
  // misleading partial figure.
  const composite = hasTechData
    ? Math.round((pillars.fundamentals.weighted +
                  pillars.technicals.weighted +
                  pillars.macro.weighted +
                  pillars.sentiment.weighted +
                  pillars.liquidity.weighted) * 10) / 10
    : null;

  const hardFailed = allHardFails.length > 0;
  const rating = ratingFromComposite(composite, hardFailed);

  return {
    company: fundCo,
    composite,
    rating,
    hardFails: allHardFails,
    hardFailed,
    fundamentalFlags,
    filterFlags,
    isRedFlag,
    pillars,
    pillarResults: { fund: fundResult, tech: techResult, macro: macroResult, senliq: senliqResult, split },
    dataComplete: hasTechData,
  };
}

// Score every company in the universe. Returns array sorted by
// composite descending; hard-failed and unrated stocks are kept in
// the array but sorted to the bottom so the UI can split them out.
export function scoreCompositeBatch(fundCompanies, techCompanies, macroCtx, weights = PILLAR_WEIGHTS) {
  // Build ticker → techCo lookup. extractTicker mirrors what
  // scrape-technicals.mjs does.
  const techByTicker = {};
  for (const t of techCompanies || []) {
    if (t && t.ticker) techByTicker[String(t.ticker).toUpperCase()] = t;
  }
  const extractTickerFromUrl = (url) => {
    const m = String(url || "").match(/\/company\/([^/]+)/);
    return m ? m[1].toUpperCase() : null;
  };

  const results = (fundCompanies || []).map((fc) => {
    const ticker = extractTickerFromUrl(fc["Screener URL"]);
    const tc = ticker ? techByTicker[ticker] : null;
    return scoreCompositeOne(fc, tc, macroCtx, weights);
  });

  // Sort: rated stocks descending by composite, then unrated, then filtered last.
  results.sort((a, b) => {
    if (a.hardFailed !== b.hardFailed) return a.hardFailed ? 1 : -1;
    if ((a.composite == null) !== (b.composite == null)) return a.composite == null ? 1 : -1;
    return (b.composite ?? -1) - (a.composite ?? -1);
  });

  return results;
}
