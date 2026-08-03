import * as fund from "./scoring.js";
import * as tech from "./tech-scoring.js";
import * as composite from "./composite-scoring.js";
import { META as RULE_META } from "./rule-meta.js";
import { exportToExcel as exportToExcelNew } from "./excel-export.js";
import { CHARGE_DEFAULTS, CHARGE_FIELDS, makeCharger, ZERO_CHARGER, roundTrip, chargeSummary, buyRate, sellRate } from "./charges.js";
import * as strat5 from "./strategies.js";

// ---------------- Tab configuration ----------------
const CONFIGS = {
  fundamentals: {
    label: "Fundamentals",
    dataUrl: "data/screener-companies.json",
    metaUrl: "data/metadata.json",
    parseData: (raw) => raw,
    rules: fund.ACTIVE_RULES,
    deferred: fund.DEFERRED,
    score: fund.scoreCompany,
    // accessors
    name: (c) => c.Company,
    marketCap: (c) => c["Market Cap"] || "",
    screenerUrl: (c) => c["Screener URL"],
    sector: (c) => c["Sector"] || null,
    industry: (c) => c["Broad Industry"] || null,
    // table columns (in addition to #, Company, Score, Signals, Link)
    columns: [
      { label: "ROE",    get: (c) => c.ROE || "—" },
      { label: "ROCE",   get: (c) => c.ROCE || "—" },
      { label: "Rev 3Y", get: (c) => c["Sales growth 3Years"] || "—" },
      { label: "PAT 3Y", get: (c) => c["Profit Var 3Yrs"] || "—" },
      { label: "D/E",    get: (c) => c["Debt to equity"] || "—" },
      { label: "P/E",    get: (c) => c["Stock P/E"] || "—" },
    ],
    // 3 stat-card values for the header strip
    stats: {
      rules: `${fund.ACTIVE_RULES.length} / ${fund.ACTIVE_RULES.length}`, rulesNote: "Active rules",
      maxScore: `${fund.ACTIVE_RULES.reduce((n, r) => n + (r.max || 0), 0) || 27} pts`, maxNote: "All rules active",
    },
    drillHeaderStats: (c) => [
      { label: "Market Cap", main: c["Market Cap"] || "—", sub: `CMP ${c["Current Price"] || "—"}` },
      { label: "Valuation & Returns",
        metrics: [
          { name: "P/E",  value: c["Stock P/E"]       || "—" },
          { name: "D/E",  value: c["Debt to equity"]  || "—" },
          { name: "ROCE", value: c["ROCE"]            || "—" },
        ],
        sub: `Capital efficiency snapshot` },
    ],
  },
  technicals: {
    label: "Technicals",
    dataUrl: "data/technicals.json",
    metaUrl: null,
    parseData: (raw) => ({ rows: raw.companies || [], meta: raw }),
    rules: tech.ACTIVE_RULES,
    deferred: tech.DEFERRED,
    score: tech.scoreCompany,
    name: (c) => c.name,
    marketCap: (c) => c.marketCap || "",
    screenerUrl: (c) => c.screenerUrl,
    sector: (c) => c.sector || null,
    industry: (c) => c.industry || null,
    columns: [
      { label: "CMP",  get: (c) => c.cmp ? "₹" + Math.round(c.cmp).toLocaleString("en-IN") : "—" },
      { label: "RSI",  get: (c) => c.rsi14 ?? "—" },
      { label: "ADX",  get: (c) => c.adx14 ?? "—" },
      { label: "6M RS", get: (c) => c.relative_strength_6m == null ? "—" : (c.relative_strength_6m > 0 ? "+" : "") + (c.relative_strength_6m * 100).toFixed(1) + "%" },
      { label: "Beta", get: (c) => c.beta_1y ?? "—" },
      { label: "ATR%", get: (c) => c.atr14_pct == null ? "—" : c.atr14_pct + "%" },
    ],
    stats: {
      rules: "16 / 16",   rulesNote: "Active rules",
      maxScore: "24 pts", maxNote: "All rules active",
    },
    drillHeaderStats: (c) => [
      { label: "CMP · 52W High", main: c.cmp ? "₹" + Math.round(c.cmp).toLocaleString("en-IN") : "—",
        sub: c.high_52w ? `52W ₹${Math.round(c.high_52w).toLocaleString("en-IN")}` : "" },
      { label: "Momentum & Risk",
        metrics: [
          { name: "RSI",  value: c.rsi14 ?? "—" },
          { name: "ADX",  value: c.adx14 ?? "—" },
          { name: "β",    value: c.beta_1y ?? "—" },
        ],
        sub: c.relative_strength_6m == null ? "" : `6M relative strength ${(c.relative_strength_6m * 100).toFixed(1)}% vs Smallcap 250` },
    ],
  },
  composite: {
    label: "AI Basket",
    composite: true,                  // marker — special loading + drill
    // Loaded specially in loadTab (fuses fundamentals + technicals + macro).
    dataUrl: "data/screener-companies.json",
    metaUrl: "data/macro.json",
    parseData: (raw) => raw,
    rules: [],                        // no per-rule breakdown — pillars instead
    deferred: [],
    score: (x) => x,                  // identity — scoring is done in loadTab
    // Accessors receive s.company (the fund row), matching all other tabs.
    name: (co) => co?.Company || "",
    marketCap: (co) => co?.["Market Cap"] || "",
    screenerUrl: (co) => co?.["Screener URL"],
    sector: (co) => co?.["Sector"] || null,
    industry: (co) => co?.["Broad Industry"] || null,
    // Column getters read from co._composite which loadTab stashes on each row.
    // Composite-as-number is already shown in the standard Score pill column —
    // no duplicate column here.
    columns: [
      { label: "Rating", html: true, get: (co) => {
        const r = co._composite?.rating;
        if (!r) return `<span class="text-slate-400">—</span>`;
        const cls = composite.ratingClass(r);
        return `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap ${cls}">${escapeHtml(r)}</span>`;
      } },
      { label: "Fund", get: (co) => `${co._composite?.pillars?.fundamentals?.raw ?? "—"}/29` },
      { label: "Tech", get: (co) => co._composite?.pillars?.technicals?.raw == null ? "—" : `${co._composite.pillars.technicals.raw}/${co._composite.pillars.technicals.max}` },
    ],
    stats: {
      get rules() { const l = weightsLabel(); return `${l.names.split(" · ").length} pillars`; },
      get rulesNote() { return `Weighted: ${weightsLabel().nums}`; },
      maxScore: "100 pts", maxNote: "≥75 STRONG BUY · 60 BUY · 45 WATCH",
    },
    drillHeaderStats: (co) => [
      { label: "Market Cap", main: co["Market Cap"] || "—", sub: `CMP ${co["Current Price"] || "—"}` },
      { label: "Rating",     main: co._composite?.rating || "—",
        sub: co._composite?.composite != null ? `Composite ${co._composite.composite.toFixed(1)} / 100` : "Unrated — data missing" },
    ],
  },
  // Premium hero page — surfaces stocks scoring above 75 composite
  // (STRONG BUY band) with hard-fails excluded. Renders bespoke cards,
  // not the standard table layout — handled by renderTopPicks().
  topPicks: {
    label: "Top Picks",
    hero: true,
  },
  // Strategy — merged History + Active. Top-level toggle picks Active
  // (re-locking basket at Daily / Weekly / Monthly cadence) vs Passive
  // (basket frozen at upload). Both anchor at lkp.generated_at and
  // share the Manual basket + Nifty comparison + per-pick accuracy.
  // The "active: true" flag is the legacy switchTab marker — the route
  // stays data-tab="active" so persisted tab state survives.
  active: {
    label: "Strategy",
    active: true,
  },

  // Custom Strategy Lab — user-defined backtests (playbook). Its own
  // bespoke section + renderer, sharing the strategy engine + viz from
  // the Strategy tab. The selection/exit knobs live only here; the
  // fixed tabs above stay untouched (founder's "keep existing fixed").
  custom: {
    label: "Custom",
    custom: true,
  },
};

// ---------------- State ----------------
// Watchlist persisted in localStorage. Stored as an array of company
// identifiers (Screener URL slug — stable across sessions). Toggled via
// the ☆/★ button in each row and the "Watchlist" filter pill in the toolbar.
// Client-adjustable pillar weights — stored in localStorage so the
// per-client tweak survives reloads. Defaults to the framework's
// PILLAR_WEIGHTS (40/35/15/5/5). AI basket + Top Picks tabs re-score
// composites whenever the user changes these.
// One place that renders the active pillar mix. Several UI spots used to
// print a frozen "40 · 35 · 15 · 5 · 5" that ignored the real weights, so the
// dashboard advertised a model it was not running. Zero-weighted pillars are
// omitted rather than shown as 0.
// Pillars carrying non-zero weight. Everything user-facing iterates this
// rather than the full five, so a zero-weighted pillar disappears from the
// sliders, the mini bars and the radar instead of rendering as a dead 0.
function livePillars(w) {
  const ALL = Object.keys(composite.PILLAR_WEIGHTS);
  const src = w || (typeof state !== "undefined" && state.pillarWeights) || composite.PILLAR_WEIGHTS;
  const live = ALL.filter((k) => (src[k] || 0) > 0);
  return live.length ? live : ALL;      // never render an empty control
}

function weightsLabel(w) {
  const LBL = { fundamentals: "Fund", technicals: "Tech", macro: "Macro", sentiment: "Sent", liquidity: "Liq" };
  const src = w || (typeof state !== "undefined" && state.pillarWeights) || composite.PILLAR_WEIGHTS;
  const live = Object.keys(LBL).filter((k) => (src[k] || 0) > 0);
  return { nums: live.map((k) => src[k]).join(" · "), names: live.map((k) => LBL[k]).join(" · ") };
}

const PILLAR_WEIGHTS_KEY = "klpdash-pillar-weights-v1";
function loadPillarWeights() {
  try {
    const stored = JSON.parse(localStorage.getItem(PILLAR_WEIGHTS_KEY) || "null");
    if (stored && typeof stored === "object") {
      // Sanity-clamp + sum normalisation guard so a corrupted value
      // can't silently break composite scoring.
      const w = {};
      for (const k of Object.keys(composite.PILLAR_WEIGHTS)) {
        w[k] = Math.max(0, Math.min(100, Number(stored[k]) || 0));
      }
      const sum = Object.values(w).reduce((a, b) => a + b, 0);
      if (sum >= 95 && sum <= 105) return w; // accept ±5 tolerance
    }
  } catch {}
  return { ...composite.PILLAR_WEIGHTS };
}
function savePillarWeights(w) {
  try { localStorage.setItem(PILLAR_WEIGHTS_KEY, JSON.stringify(w)); } catch {}
}

// Weight Lab — a sandbox pillar-weight mix, independent of the live SPIP
// weights, so the desk can experiment (and let AI search) without touching
// the published basket until they hit "Apply to SPIP".
const LAB_WEIGHTS_KEY = "klpdash-lab-weights-v1";
function loadLabWeights() {
  try {
    const s = JSON.parse(localStorage.getItem(LAB_WEIGHTS_KEY) || "null");
    if (s && Object.keys(composite.PILLAR_WEIGHTS).every((k) => typeof s[k] === "number")) {
      // Migrate mixes saved before the sum-to-100 cap (old sliders let each
      // pillar reach 100 independently) so a restored total never exceeds 100.
      const total = Object.keys(composite.PILLAR_WEIGHTS).reduce((a, k) => a + (s[k] || 0), 0);
      return total > 100 ? normalizeWeights(s) : s;
    }
  } catch {}
  return { ...composite.PILLAR_WEIGHTS };
}
function saveLabWeights(w) {
  try { localStorage.setItem(LAB_WEIGHTS_KEY, JSON.stringify(w)); } catch {}
}

// 1-year technical back-test settings (defined here so state init can read
// them — the rest of the back-test lives down by the Custom tab).
const TECH_PARAMS_KEY = "klpdash-tech-params-v1";
const TECH_DEFAULTS = { basketSize: 7, rebalanceDays: 10, targetPct: 12, slPct: 8, threshold: 55, windowDays: 252 };
function loadTechParams() { try { const r = JSON.parse(localStorage.getItem(TECH_PARAMS_KEY)); if (r) return { ...TECH_DEFAULTS, ...r }; } catch {} return { ...TECH_DEFAULTS }; }
function saveTechParams(p) { try { localStorage.setItem(TECH_PARAMS_KEY, JSON.stringify(p)); } catch {} }

// Client-configurable allocation per pick for the History portfolio
// backtest. Defaults to ₹1L. Stored as a plain integer in localStorage.
// Trade Plan capital — the total rupees to deploy across this month's basket.
// Editable in the Strategy > Plan tab; the whole sizing table recomputes from it.
const PLAN_CAPITAL_KEY = "vn-plan-capital-v1";
const PLAN_STOP_ATR    = 2.5;    // stop = entry - 2.5 x ATR%  (see the manual)
const PLAN_RSI_HOT      = 75;    // above this the stock has already run
const PLAN_MAX_PER_SECTOR = 3;   // more than this in one industry is a sector bet
function loadPlanCapital() {
  try { const v = Number(localStorage.getItem(PLAN_CAPITAL_KEY)); if (Number.isFinite(v) && v >= 1000) return v; } catch {}
  return 50000;
}
function savePlanCapital(n) { try { localStorage.setItem(PLAN_CAPITAL_KEY, String(n)); } catch {} }

const ALLOC_KEY = "klpdash-alloc-per-pick-v1";
function loadAllocPerPick() {
  try {
    const v = Number(localStorage.getItem(ALLOC_KEY));
    if (Number.isFinite(v) && v >= 1000) return v;
  } catch {}
  return 100000;
}
function saveAllocPerPick(n) {
  try { localStorage.setItem(ALLOC_KEY, String(n)); } catch {}
}

// Client's locked-in cohort, one per month. Founder asked for parity
// with the way the firm actually runs: at each month-end, the client
// picks their own 7 stocks (CSV upload here) and we track them along-
// side our SPIP top 7. Keyed by month "YYYY-MM" (the month the cohort
// is HELD, not the month-end entry date — June 2026 = "2026-06"; the
// entry close comes from the last May snapshot).
const CLIENT_COHORT_KEY = "klpdash-client-cohort-v1";
function loadClientCohorts() {
  try { return JSON.parse(localStorage.getItem(CLIENT_COHORT_KEY) || "{}"); }
  catch { return {}; }
}
function saveClientCohorts(d) {
  try { localStorage.setItem(CLIENT_COHORT_KEY, JSON.stringify(d)); } catch {}
}

// Performance Tracker view ("static" | "monthly" | "weekly"). Persisted
// in localStorage so the analyst's choice survives reloads.
//   - static  : prior month-end snapshot is the anchor; AI top 7 frozen
//               for the whole new month (the SPIP production model)
//   - monthly : anchored at the client's upload date; AI top 7 frozen
//               from THAT date — handles mid-month uploads
//   - weekly  : anchored at upload date; AI re-locks each Monday
const COHORT_VIEW_KEY = "klpdash-cohort-view-v1";
// "monthly" removed — it anchored at the client's upload date, which does
// not exist here. "static" IS the monthly model: picked at month start,
// held for the month.
const COHORT_VIEWS = ["static", "weekly"];
function loadCohortView() {
  try {
    const v = localStorage.getItem(COHORT_VIEW_KEY);
    return COHORT_VIEWS.includes(v) ? v : "static";
  } catch { return "static"; }
}
function saveCohortView(v) {
  try { localStorage.setItem(COHORT_VIEW_KEY, v); } catch {}
}

// Strategy tab top-level mode ("active" | "passive"). Active = AI
// basket rebalances at the chosen cadence (Daily / Weekly / Monthly);
// Passive = AI top 7 picked once at upload and frozen forever (single
// segment). Both share the Manual basket comparison + Nifty benchmark
// + per-pick accuracy.
const STRATEGY_MODE_KEY = "klpdash-strategy-mode-v1";
const STRATEGY_MODES = ["active", "passive"];
function loadStrategyMode() {
  try {
    const v = localStorage.getItem(STRATEGY_MODE_KEY);
    return STRATEGY_MODES.includes(v) ? v : "passive";
  } catch { return "passive"; }
}
function saveStrategyMode(v) {
  try { localStorage.setItem(STRATEGY_MODE_KEY, v); } catch {}
}

// Active sub-cadence ("daily" | "weekly" | "monthly"). Drives the
// rebalance frequency / segment length when strategyMode === "active".
// All three anchor at the client upload date (lkp.generated_at).
const ACTIVE_CADENCE_KEY = "klpdash-active-cadence-v1";
// Daily rebalancing retired — it re-picks 7 fresh names every day, the
// opposite of the client's "pick 7 and hold" product, and only added
// noise. Weekly / Monthly remain for rebalancing-frequency experiments.
const ACTIVE_CADENCES = ["weekly", "monthly"];
function loadActiveCadence() {
  try {
    const v = localStorage.getItem(ACTIVE_CADENCE_KEY);
    return ACTIVE_CADENCES.includes(v) ? v : "weekly";
  } catch { return "weekly"; }
}
function saveActiveCadence(v) {
  try { localStorage.setItem(ACTIVE_CADENCE_KEY, v); } catch {}
}

// Basket return mode ("booked" | "held") — applies to BOTH baskets (AI and
// manual) so the comparison is like-to-like. Booked (default) = each pick
// freezes the day it first hits its first target (cap the gain) or SL (cap
// the loss) — the desk's real exit, so one runaway winner can't flatter the
// basket. Held = mark-to-market to today, ignoring the exits ("what if we
// never sold"). The Strategy-tab toggle flips it and every figure — both
// headlines, both chart lines, basket rows — follows.
const MANUAL_RETURN_KEY = "klpdash-manual-return-mode-v1";
const MANUAL_RETURN_MODES = ["booked", "held"];
function loadManualReturnMode() {
  try {
    const v = localStorage.getItem(MANUAL_RETURN_KEY);
    return MANUAL_RETURN_MODES.includes(v) ? v : "booked";
  } catch { return "booked"; }
}
function saveManualReturnMode(v) {
  try { localStorage.setItem(MANUAL_RETURN_KEY, v); } catch {}
}

// Strategy-tab sub-tab ("overview" | "accuracy" | "sector" | "industry" |
// "capital"). Keeps the tab on one screen — the main view (chart + picks +
// alpha) is the default; everything else is one click away, no long scroll.
// v2: the default moved from "plan" to "overview". Bumping the key rather
// than just changing the fallback, so anyone carrying a stored "plan" from
// the old default lands on Overview too instead of never seeing the change.
const STRATEGY_SUBTAB_KEY = "klpdash-strategy-subtab-v2";
const STRATEGY_SUBTABS = ["plan", "compare", "overview", "accuracy", "balancing"]; // Sector + Industry merged into "balancing" (client ask); "capital" parked — renderSimPanel + its case kept below for easy restore
function loadStrategySubTab() {
  try {
    let v = localStorage.getItem(STRATEGY_SUBTAB_KEY);
    if (v === "sector" || v === "industry") v = "balancing";   // migrate the old split sub-tabs
    return STRATEGY_SUBTABS.includes(v) ? v : "overview";
  }
  catch { return "overview"; }
}
function saveStrategySubTab(v) {
  try { localStorage.setItem(STRATEGY_SUBTAB_KEY, v); } catch {}
}

// Focus mode. Hides the rotation / rebalancing strategy views — the
// Strategy-tab "Active" mode and the Custom-tab strategy cards — so the
// dashboard shows only the client's core product: the held basket vs the
// manual basket, plus the Weight Lab. Nothing is deleted; flip this to
// true to bring the rotation strategies back exactly as they were.
const SHOW_ROTATION_STRATEGIES = false;

// ── Simulation inputs: capital, cash buffer, transaction charges ──────
// Everything here is adjustable from the Strategy tab's "Capital &
// charges" panel and persisted locally — the "real world" layer on top
// of the model. Defaults are realistic Indian-equity delivery rates;
// they are deliberately NOT the client's exact figures (he hasn't sent
// them) — they're knobs he overwrites in the dashboard, and the curve
// recomputes. Capital is the total pot; bufferPct is held back as idle
// cash (his "buffer for charges / maintenance" idea), the rest is
// deployed equal-weight across the basket.
// Charge rates themselves live in charges.js — verified against the
// Zerodha calculator. The old model here was a single symmetric "% per
// side", which got three things wrong: it charged delivery brokerage
// (there is none), it missed buy-side stamp duty, and — the one that
// actually matters — it ignored the flat ₹15.34 DP fee on every sell.
// Being flat, that fee is 0.43% of a ₹3,500 position and 0.11% of a
// ₹14,000 one, so the old model understated the cost of small tickets
// and of churn alike.
const SIM_PREFS_KEY = "klpdash-sim-prefs-v2";   // v2: charge model changed shape
const SIM_DEFAULTS = {
  capital: 50000,      // ₹ total pot — the account this is actually run at
  bufferPct: 0,        // % of capital kept as idle cash reserve
  ...CHARGE_DEFAULTS,
};
const SIM_FIELDS = [
  { key: "capital",   label: "Capital (₹)",     step: 5000, money: true },
  { key: "bufferPct", label: "Cash buffer (%)", step: 1 },
  ...CHARGE_FIELDS,
];
function loadSimPrefs() {
  try {
    const raw = JSON.parse(localStorage.getItem(SIM_PREFS_KEY));
    if (raw && typeof raw === "object") return { ...SIM_DEFAULTS, ...raw };
  } catch {}
  return { ...SIM_DEFAULTS };
}
function saveSimPrefs(p) {
  try { localStorage.setItem(SIM_PREFS_KEY, JSON.stringify(p)); } catch {}
}
let simPrefs = loadSimPrefs();

// The engines take a charger rather than a rate, because a sell charge is
// not proportional to notional — it carries the flat DP fee.
function charger(p = simPrefs) { return makeCharger(p); }
function zeroCharger() { return ZERO_CHARGER; }

// History tab sub-view ("history" | "accuracy"). Toggled by an inline
// switch below the Performance Tracker; persisted in localStorage so
// the analyst's last choice survives reloads.
const HISTORY_VIEW_KEY = "klpdash-history-view-v1";
function loadHistoryView() {
  try {
    const v = localStorage.getItem(HISTORY_VIEW_KEY);
    return v === "accuracy" ? "accuracy" : "history";
  } catch { return "history"; }
}
function saveHistoryView(v) {
  try { localStorage.setItem(HISTORY_VIEW_KEY, v); } catch {}
}

// "Recompute history with current pillar weights" opt-in. Snapshots are
// stored with v1 framework weights — when the analyst changes weights
// in the AI basket pillar editor, the History tab still shows the
// v1 composites by default. Opting in here re-derives each snapshot's
// composite + rating from the stored per-pillar pct × current weights.
const RECOMPUTE_HISTORY_KEY = "klpdash-recompute-history-v1";
function loadRecomputeHistory() {
  try { return localStorage.getItem(RECOMPUTE_HISTORY_KEY) === "1"; } catch { return false; }
}
function saveRecomputeHistory(b) {
  try {
    if (b) localStorage.setItem(RECOMPUTE_HISTORY_KEY, "1");
    else localStorage.removeItem(RECOMPUTE_HISTORY_KEY);
  } catch {}
}

// Apply current weights to every cached snapshot, mutating composite +
// rating in place. Cheap: snapshots store pillars.pct already, so the
// new composite is just a weighted sum + a rating-band lookup. Doesn't
// touch hardFailed (rule-based, not weight-based).
function recomputeSnapshotsWithWeights(snapshots, weights) {
  for (const snap of snapshots) {
    for (const s of snap.stocks) {
      const p = s.pillars;
      if (!p) continue;
      const fundPct = p.fundamentals?.pct;
      const techPct = p.technicals?.pct;
      const macroPct = p.macro?.pct;
      const sentPct = p.sentiment?.pct;
      const liqPct = p.liquidity?.pct;
      if ([fundPct, techPct, macroPct, sentPct, liqPct].some((v) => v == null)) continue;
      const newComposite = (
        fundPct * weights.fundamentals +
        techPct * weights.technicals +
        macroPct * weights.macro +
        sentPct * weights.sentiment +
        liqPct * weights.liquidity
      ) / 100;
      s.composite = Number(newComposite.toFixed(2));
      s.rating = composite.ratingFromComposite(s.composite, s.hardFailed);
    }
  }
}

function weightsMatchDefault(w) {
  if (!w) return true;
  const d = composite.PILLAR_WEIGHTS;
  return ["fundamentals", "technicals", "macro", "sentiment", "liquidity"].every((k) => w[k] === d[k]);
}

const WATCHLIST_KEY = "klpdash-watchlist-v1";
function loadWatchlist() {
  try { return new Set(JSON.parse(localStorage.getItem(WATCHLIST_KEY) || "[]")); }
  catch { return new Set(); }
}
function saveWatchlist(set) {
  try { localStorage.setItem(WATCHLIST_KEY, JSON.stringify([...set])); }
  catch {}
}
function companyKey(co) {
  // Stable identifier: URL slug from Screener URL (NSE ticker or BSE code)
  const slug = String(co?.["Screener URL"] || co?.["Composite URL"] || "").match(/\/company\/([^/]+)/)?.[1];
  return slug ? slug.toUpperCase() : (co?.Company || co?.name || "").toUpperCase();
}

const state = {
  activeTab: "fundamentals",
  cache: {},                  // tab → { scored, raw, meta, filtered }
  search: "",
  scoreFilter: "all",
  sortBy: "score",
  sortDir: "desc",
  watchlist: loadWatchlist(),
  watchOnly: false,
  pillarWeights: loadPillarWeights(),
  planCapital: loadPlanCapital(),   // Trade Plan sizing input
  paperMonth: null,                 // which month Paper Results is showing
  allocPerPick: loadAllocPerPick(),
  cohortMonth: null,            // legacy — kept for compat; unused after move to anchor-date model
  cohortView: loadCohortView(), // "static" | "monthly" | "weekly"
  cohortSegmentIdx: null,       // which week pill is selected (null = latest)
  historyView: loadHistoryView(), // "history" | "accuracy"
  strategyMode: loadStrategyMode(),  // "active" | "passive"
  activeCadence: loadActiveCadence(), // "daily" | "weekly" | "monthly" (used when strategyMode === "active")
  manualReturnMode: loadManualReturnMode(), // "booked" | "held" — manual-basket return convention
  strategySubTab: loadStrategySubTab(),     // Strategy-tab sub-tab (plan / compare / overview / accuracy / balancing)
  compareView: (() => { try { return localStorage.getItem("vn-compare-view-v1") || "summary"; } catch { return "summary"; } })(),
  compareSource: (() => { try { return localStorage.getItem("vn-compare-source-v1") || "live"; } catch { return "live"; } })(),
  planStrategyId: (() => { try { return localStorage.getItem("vn-plan-strategy-v1") || null; } catch { return null; } })(),
  manualMonth: null,                  // selected client-basket month "YYYY-MM" (null = latest)
  labWeights: loadLabWeights(),       // Weight Lab sandbox pillar mix
  labAiBest: null,                    // last AI weight-search result (in-memory)
  techParams: loadTechParams(),       // 1-year technical back-test settings
  techAiBest: null,                   // last AI tech-strategy search (in-memory)
  strategySegmentIdx: null,           // which segment pill is selected (null = latest)
  recomputeHistory: loadRecomputeHistory(),
  // Lazy composite cache — populated on first drill-down or when composite
  // tab loads. Maps slug → composite result { pillars, composite, rating, ... }.
  compositeBySlug: new Map(),
  lazyTechBySlug: null,        // Map slug → tech row (loaded on demand)
  lazyMacroCtx: null,          // raw macro.json (loaded on demand)
  lazyFundByTicker: null,      // Map ticker → fundamentals row (loaded on demand)
};

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ---------------- visual helpers ----------------
const PALETTE = [
  "from-purple-500 to-indigo-500", "from-pink-500 to-rose-500", "from-amber-500 to-orange-500",
  "from-emerald-500 to-teal-500", "from-blue-500 to-cyan-500", "from-fuchsia-500 to-pink-500",
  "from-violet-500 to-purple-500", "from-lime-500 to-emerald-500", "from-sky-500 to-indigo-500",
  "from-red-500 to-pink-500", "from-yellow-500 to-amber-500", "from-teal-500 to-cyan-500",
];
function avatarFor(name) {
  let h = 0; for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  const color = PALETTE[h % PALETTE.length];
  const initials = name.split(/\s+/).slice(0, 2).map((w) => w[0] || "").join("").toUpperCase().slice(0, 2) || "?";
  return { color, initials };
}
function scoreBadgeClass(pct) {
  if (pct >= 80) return "bg-emerald-100 text-emerald-700 ring-1 ring-emerald-200";
  if (pct >= 60) return "bg-blue-100 text-blue-700 ring-1 ring-blue-200";
  if (pct >= 40) return "bg-amber-100 text-amber-700 ring-1 ring-amber-200";
  return "bg-rose-100 text-rose-700 ring-1 ring-rose-200";
}
function scoreTier(pct) {
  if (pct >= 80) return "excellent";
  if (pct >= 60) return "good";
  if (pct >= 40) return "average";
  return "weak";
}
function tierLabel(t) { return ({ excellent: "Excellent", good: "Good", average: "Average", weak: "Weak", hardfail: "Hard Fail" })[t]; }
function tierColor(t) { return ({ excellent: "text-emerald-600", good: "text-blue-600", average: "text-amber-600", weak: "text-rose-600", hardfail: "text-rose-700" })[t]; }
// Gradient theme keyed off the per-tab tier — feeds renderScoreGauge / renderPillarRadar
// (both expect the { from, to } shape that ratingTheme returns for composite tab).
function themeFromTier(t) {
  const map = {
    excellent: { from: "from-emerald-500", to: "to-teal-500" },
    good:      { from: "from-blue-500",    to: "to-indigo-500" },
    average:   { from: "from-amber-500",   to: "to-orange-500" },
    weak:      { from: "from-rose-500",    to: "to-pink-500" },
    hardfail:  { from: "from-slate-500",   to: "to-slate-600" },
  };
  return map[t] || { from: "from-slate-400", to: "to-slate-500" };
}
function statusPill(status) {
  switch (status) {
    case "pass":      return `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200">✓ Pass</span>`;
    case "partial":   return `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-amber-50 text-amber-700 ring-1 ring-amber-200">~ Partial</span>`;
    case "fail":      return `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-rose-50 text-rose-700 ring-1 ring-rose-200">✕ Fail</span>`;
    case "hard_fail": return `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-rose-100 text-rose-800 ring-1 ring-rose-300">⚠ Hard Fail</span>`;
    case "na":        return `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-slate-100 text-slate-600 ring-1 ring-slate-200">— N/A</span>`;
    default: return "";
  }
}
function escapeHtml(s) { return String(s ?? "").replace(/[&<>"']/g, (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }

// SVG icons (heroicons-style, monoline)
const ICON_LINK = `<svg class="w-3 h-3 flex-shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M10 6H6a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4M14 4h6m0 0v6m0-6L10 14"/></svg>`;
const ICON_CALC = `<svg class="w-3 h-3 flex-shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 7h6m-6 4h6m-3 4h3M7 21h10a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2z"/></svg>`;
const ICON_LOGIC = `<svg class="w-3 h-3 flex-shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5.586a1 1 0 0 1 .707.293l5.414 5.414a1 1 0 0 1 .293.707V19a2 2 0 0 1-2 2z"/></svg>`;
const ICON_CHEVRON = `<svg class="w-2.5 h-2.5 flex-shrink-0 text-slate-400 transition-transform group-open:rotate-90" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 5l7 7-7 7"/></svg>`;

// Renders 3 small chips on each drill-down rule card:
//   Source — opens the actual data source in a new tab
//   Calculation — expands to show the formula (only when computed)
//   Implementation — expands; if we deviate from client (source, calc,
//   or scoring), chip is amber + shows client logic vs our implementation
function renderRuleMetaButtons(ruleKey, company) {
  const tab = state.activeTab;
  const meta = RULE_META[tab]?.[ruleKey];
  if (!meta) return "";
  const src = typeof meta.source === "function" ? meta.source(company || {}) : meta.source;

  const baseChip = "group inline-flex items-center gap-1.5 px-2.5 py-1.5 text-[11px] font-medium rounded-md border transition-colors select-none";
  const neutralChip = "bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 border-slate-200 hover:border-slate-300";
  const amberChip   = "bg-amber-50 hover:bg-amber-100 text-amber-800 hover:text-amber-900 border-amber-200 hover:border-amber-300";

  const sourceBtn = src ? `
    <a href="${escapeHtml(src.url)}" target="_blank" rel="noopener"
       class="${baseChip} ${neutralChip} no-underline"
       title="${escapeHtml(src.label + (src.section ? " — " + src.section : ""))}">
      ${ICON_LINK}<span>${escapeHtml(src.label)}</span>
    </a>` : "";

  const calcBtn = meta.calculation ? `
    <details class="meta-details">
      <summary class="${baseChip} ${neutralChip} cursor-pointer">
        ${ICON_CALC}<span>Calculation</span>${ICON_CHEVRON}
      </summary>
      <div class="mt-2 p-3 bg-slate-50 border border-slate-200 rounded-md text-xs text-slate-700 leading-relaxed">
        ${escapeHtml(meta.calculation)}
      </div>
    </details>` : "";

  const logicBody = meta.ourLogic
    ? `<div>
         <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1">Client's scoring logic</div>
         <div class="text-slate-700 mb-3">${escapeHtml(meta.clientLogic)}</div>
         <div class="text-[10px] font-bold uppercase tracking-wider text-amber-700 mb-1">Our implementation</div>
         <div class="text-amber-800">${escapeHtml(meta.ourLogic)}</div>
       </div>`
    : `<div>
         <div class="text-slate-700">${escapeHtml(meta.clientLogic)}</div>
         <div class="text-[10px] text-emerald-600 mt-2 font-bold uppercase tracking-wider">✓ Matches our implementation exactly</div>
       </div>`;

  const logicBtn = meta.clientLogic ? `
    <details class="meta-details">
      <summary class="${baseChip} ${meta.ourLogic ? amberChip : neutralChip} cursor-pointer">
        ${ICON_LOGIC}<span>Implementation${meta.ourLogic ? " · diff" : ""}</span>${ICON_CHEVRON}
      </summary>
      <div class="mt-2 p-3 bg-slate-50 border border-slate-200 rounded-md text-xs leading-relaxed">${logicBody}</div>
    </details>` : "";

  if (!sourceBtn && !calcBtn && !logicBtn) return "";
  return `
    <div class="flex flex-wrap items-center gap-2 mt-3 pt-3 border-t border-slate-100">
      ${sourceBtn}${calcBtn}${logicBtn}
    </div>`;
}
function cfg() { return CONFIGS[state.activeTab]; }
function tabState() { return state.cache[state.activeTab]; }

// ---------------- load / switch tab ----------------
async function switchTab(tabId) {
  if (!CONFIGS[tabId]) return;
  // Toggle nav. All tabs carry a permanent transparent `border-b-2`
  // (for layout consistency — keeps the row height constant regardless
  // of active state). The visible underline is painted exclusively by
  // the .tab-btn::after pseudo-element driven by .is-active, so we
  // never touch border-colour classes here (toggling those alongside
  // ::after produces the double-underline reported by user).
  $$(".tab-btn").forEach((b) => {
    // Top Picks has no nav tab of its own — it's launched from the SPIP
    // (composite) tab, so keep SPIP highlighted while viewing it.
    const active = b.dataset.tab === tabId || (tabId === "topPicks" && b.dataset.tab === "composite");
    b.classList.toggle("text-indigo-600", active);
    b.classList.toggle("text-slate-500", !active);
    b.classList.toggle("hover:text-slate-700", !active);
    b.classList.toggle("is-active", active);
  });
  // "★ Top Picks" launcher button lives in the controls row, SPIP tab only.
  $("#top-picks-btn")?.classList.toggle("hidden", tabId !== "composite");
  state.activeTab = tabId;
  state.search = "";
  state.scoreFilter = "all";
  state.sortBy = "score";
  state.sortDir = "desc";
  $("#search").value = "";
  $("#score-filter").value = "all";

  // Section toggles per tab type:
  // - Hero    (Top Picks): only #top-picks-section visible
  // - History            : only #history-section visible
  // - Active             : only #active-section visible
  // - Normal             : all other sections
  const c = CONFIGS[tabId];
  const hero = !!c?.hero;
  const history = !!c?.history;
  const active = !!c?.active;
  const custom = !!c?.custom;
  const bespoke = hero || history || active || custom;
  document.querySelectorAll("main > section").forEach((sec) => {
    if (sec.id === "top-picks-section") sec.classList.toggle("hidden", !hero);
    else if (sec.id === "history-section") sec.classList.toggle("hidden", !history);
    else if (sec.id === "active-section") sec.classList.toggle("hidden", !active);
    else if (sec.id === "custom-section") sec.classList.toggle("hidden", !custom);
    else sec.classList.toggle("hidden", bespoke);
  });
  if (hero) {
    // Composite scoring drives Top Picks. Lazy-build the composite cache.
    if (!state.cache.composite) await loadTab("composite");
    renderTopPicks();
    return;
  }
  if (history) {
    await renderHistory();
    return;
  }
  if (active) {
    await renderActive();
    return;
  }
  if (custom) {
    await renderCustom();
    return;
  }

  if (!state.cache[tabId]) await loadTab(tabId);
  renderAll();
}

async function loadTab(tabId) {
  const c = CONFIGS[tabId];

  // Composite (AI Basket) tab: fuse fundamentals + technicals + macro
  // (the same enrichments the per-pillar tabs do) and run the weighted
  // composite scorer. The output rows are composite-result objects, not
  // plain companies — drill/table renderers branch on c.composite.
  if (c.composite) {
    const [fundData, techData, macroData] = await Promise.all([
      fetch("data/screener-companies.json").then((r) => r.json()),
      fetch("data/technicals.json").then((r) => r.json()),
      // Optional: macro.json is not produced by this dashboard. Inside a
      // Promise.all a 404 would reject and blank the entire page, so swallow it.
      fetch("data/macro.json").then((r) => (r.ok ? r.json() : null)).catch(() => null),
    ]);
    const fundCompanies = Array.isArray(fundData) ? fundData : (fundData.companies || []);
    const techCompanies = techData.companies || [];

    // Apply the same per-tab enrichments so each pillar scores identically
    // to the standalone tab. Best-effort fetches.
    let insiderByTicker = {}, governanceByTicker = {}, auditorByTicker = {};
    let insiderLoaded = false, governanceLoaded = false, auditorLoaded = false;
    try { insiderByTicker = await loadInsiderMerged(); insiderLoaded = Object.keys(insiderByTicker).length > 0; } catch {}
    // Per-company annual-report extraction (replaces sector proxies in macro rules
    // once extraction lands for a ticker — routine processes ~10/day in SPIP order).
    let revenueMixByTicker = {};
    try { revenueMixByTicker = (await fetch("data/company-revenue-mix.json").then((r) => r.json()))?.companies || {}; } catch {}
    try { const j = await fetch("data/governance-flags.json").then((r) => r.json()); governanceByTicker = j?.flagged_companies || {}; governanceLoaded = !!j && !j.error; } catch {}
    try { const j = await fetch("data/auditor-opinions.json").then((r) => r.json()); auditorByTicker = j?.companies || {}; auditorLoaded = !!j && Object.keys(auditorByTicker).length > 0; } catch {}
    const pli   = new Set((macroData?.pli_companies || []).map((s) => String(s).toUpperCase()));
    const renew = new Set((macroData?.renewable_companies || []).map((s) => String(s).toUpperCase()));
    const cp1   = new Set((macroData?.china_plus_one_companies || []).map((s) => String(s).toUpperCase()));
    for (const row of fundCompanies) {
      const m = String(row["Screener URL"] || "").match(/\/company\/([^/]+)/);
      const ticker = m ? m[1].toUpperCase() : null;
      // Insider / governance / auditor (Fundamentals enrichments)
      const ins = ticker ? insiderByTicker[ticker] : null;
      row.insider_loaded = insiderLoaded;
      if (ins) {
        row.insider_net_shares = ins.net_shares; row.insider_net_value = ins.net_value;
        row.insider_buy_shares = ins.buy_shares; row.insider_sell_shares = ins.sell_shares;
        row.insider_transactions = ins.transactions; row.insider_last_date = ins.last_date;
        row.insider_pledges_excluded = ins.pledges_excluded || 0;
      } else { row.insider_transactions = 0; }
      row.governance_loaded = governanceLoaded;
      row.governance_flag = ticker && governanceByTicker[ticker] ? governanceByTicker[ticker] : null;
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
      // Per-company revenue-mix extraction (truth from annual report,
      // replaces sector proxy in macro rules once extraction lands).
      row._revenue_mix = ticker ? (revenueMixByTicker[ticker] || null) : null;
    }
    // ATR history for technicals
    try {
      const atrHistory = await fetch("data/atr-history.json").then((r) => r.json());
      for (const row of techCompanies) if (row.ticker && atrHistory[row.ticker]) row.atr_history = atrHistory[row.ticker];
    } catch {}

    const compositeResults = composite.scoreCompositeBatch(fundCompanies, techCompanies, macroData, state.pillarWeights);
    // Stash composite result on the company so column getters can reach it,
    // and map to the score-shape existing renderers expect.
    const scored = compositeResults.map((r) => {
      r.company._composite = r;
      return {
        company: r.company,
        composite: r.composite,
        rating: r.rating,
        pillars: r.pillars,
        pillarResults: r.pillarResults,
        hardFails: r.hardFails,
        hardFailed: r.hardFailed,
        fundamentalFlags: r.fundamentalFlags || [],
        filterFlags: r.filterFlags || [],
        isRedFlag: !!r.isRedFlag,
        unrated: !r.dataComplete,
        totalPoints: r.composite ?? 0,
        totalMax: 100,
        scorePct: r.composite != null ? Math.round(r.composite) : 0,
        breakdown: [],
        naCount: 0,
      };
    });
    state.cache[tabId] = { rows: fundCompanies, scored, meta: macroData, filtered: scored };
    return;
  }

  const [rawData, rawMeta] = await Promise.all([
    fetch(c.dataUrl).then((r) => r.json()),
    c.metaUrl ? fetch(c.metaUrl).then((r) => (r.ok ? r.json() : null)).catch(() => null) : Promise.resolve(null),
  ]);
  const parsed = c.parseData(rawData);
  const rows = parsed.rows || parsed;
  const meta = parsed.meta || rawMeta || rawData;

  // Fundamentals tab: merge insider-trades.json + governance-flags.json
  // onto each row by NSE ticker. Both files are best-effort — missing
  // or empty file degrades cleanly to N/A in the associated rule.
  if (tabId === "fundamentals") {
    let insiderByTicker = {};
    let insiderLoaded = false;
    try {
      insiderByTicker = await loadInsiderMerged();
      insiderLoaded = Object.keys(insiderByTicker).length > 0;
    } catch { /* insider file missing — rule shows N/A */ }

    let governanceByTicker = {};
    let governanceLoaded = false;
    try {
      const gov = await fetch("data/governance-flags.json").then((r) => r.json());
      governanceByTicker = gov?.flagged_companies || {};
      // Treat the file as "loaded" whenever it parses — even an empty
      // flagged_companies map is a real signal ("no SEBI proceedings"),
      // not "data missing".
      governanceLoaded = !!gov && !gov.error;
    } catch { /* governance file missing — rule shows N/A */ }

    let auditorByTicker = {};
    let auditorLoaded = false;
    try {
      const aud = await fetch("data/auditor-opinions.json").then((r) => r.json());
      auditorByTicker = aud?.companies || {};
      auditorLoaded = !!aud && Object.keys(auditorByTicker).length > 0;
    } catch { /* file missing — rule shows N/A until first refresh writes it */ }

    for (const row of rows) {
      const m = String(row["Screener URL"] || "").match(/\/company\/([^/]+)/);
      const ticker = m ? m[1].toUpperCase() : null;
      const insiderData = ticker ? insiderByTicker[ticker] : null;
      row.insider_loaded = insiderLoaded;
      if (insiderData) {
        row.insider_net_shares  = insiderData.net_shares;
        row.insider_net_value   = insiderData.net_value;
        row.insider_buy_shares  = insiderData.buy_shares;
        row.insider_sell_shares = insiderData.sell_shares;
        row.insider_transactions = insiderData.transactions;
        row.insider_pledges_excluded = insiderData.pledges_excluded || 0;
        row.insider_last_date   = insiderData.last_date;
      } else {
        row.insider_transactions = 0;
      }
      row.governance_loaded = governanceLoaded;
      row.governance_flag = ticker && governanceByTicker[ticker] ? governanceByTicker[ticker] : null;
      row.auditor_opinions_loaded = auditorLoaded;
      const auditEntry = ticker ? auditorByTicker[ticker] : null;
      row.auditor_opinion = auditEntry?.opinion || null;
      row.auditor_opinion_source = auditEntry?.source || null;
      row.auditor_firm = auditEntry?.auditor_firm || null;
      row.auditor_opinion_date = auditEntry?.auditor_opinion_date || null;
      row.auditor_report_year = auditEntry?.auditor_report_year || null;
      row.auditor_emphasis_of_matter = auditEntry?.auditor_emphasis_of_matter || null;
      row.auditor_key_concerns = auditEntry?.auditor_key_concerns || null;
      row.auditor_confidence = auditEntry?.auditor_confidence || null;
    }
  }

  // Macro tab: merge the loaded macro.json into each row as ._macro, and
  // set per-company convenience flags in_pli / in_renewable / in_china_plus_one
  // based on NSE ticker (extracted from Screener URL slug).
  if (tabId === "macro" && rawMeta) {
    const pli = new Set((rawMeta?.pli_companies || []).map((s) => String(s).toUpperCase()));
    const renew = new Set((rawMeta?.renewable_companies || []).map((s) => String(s).toUpperCase()));
    const cp1 = new Set((rawMeta?.china_plus_one_companies || []).map((s) => String(s).toUpperCase()));
    let revenueMixByTicker = {};
    try { revenueMixByTicker = (await fetch("data/company-revenue-mix.json").then((r) => r.json()))?.companies || {}; } catch {}
    for (const row of rows) {
      const m = String(row["Screener URL"] || "").match(/\/company\/([^/]+)/);
      const ticker = m ? m[1].toUpperCase() : null;
      row._macro = rawMeta;
      row.in_pli = ticker ? pli.has(ticker) : false;
      row.in_renewable = ticker ? renew.has(ticker) : false;
      row.in_china_plus_one = ticker ? cp1.has(ticker) : false;
      row._revenue_mix = ticker ? (revenueMixByTicker[ticker] || null) : null;
    }
  }

  // Sentiment & Liquidity: tab data is technicals.json (gives us ADTV +
  // F&O eligibility + bid/ask snapshot per company), and macro.json provides
  // the market-wide sentiment context (VIX, FII/DII flow, breadth). Plus
  // we fold in sentiment-extras.json — Firecrawl-sourced PCR + per-ticker
  // Impact Cost map (NSE blocks our IPs directly; Firecrawl proxies through).
  if (tabId === "sentiment" && rawMeta) {
    let extras = null;
    try { extras = await fetch("data/sentiment-extras.json").then((r) => r.json()); }
    catch { /* sentiment-extras.json missing — rules stay deferred */ }
    const impactByTicker = extras?.impact_cost?.companies || {};
    // Stamp PCR onto the macro object so the rule reads from a stable place.
    if (rawMeta && extras?.pcr?.value != null) {
      rawMeta.sentiment = rawMeta.sentiment || {};
      rawMeta.sentiment.put_call_ratio = extras.pcr.value;
      rawMeta.sentiment.put_call_ratio_source = extras.pcr.source;
      rawMeta.sentiment.put_call_ratio_stale = !!extras.pcr.stale;
    }
    for (const row of rows) {
      row._macro = rawMeta;
      // Ticker key: prefer the explicit field, else derive from Screener URL slug.
      let ticker = (row.ticker || row.symbol || "").toUpperCase();
      if (!ticker) {
        const m = String(row.screenerUrl || "").match(/\/company\/([^/]+)/);
        if (m) ticker = m[1].toUpperCase();
      }
      const ic = ticker ? impactByTicker[ticker] : null;
      // impact_cost map values can be either a bare number or { impact_cost_pct }
      if (typeof ic === "number") row.impact_cost_pct = ic;
      else if (ic && ic.impact_cost_pct != null) row.impact_cost_pct = ic.impact_cost_pct;
      // bid_ask_spread_pct already arrives on each technicals row when Yahoo's
      // snapshot meta carried it — nothing extra to do here.
    }
  }

  // Technicals + Sentiment tabs: merge ATR history per ticker so the ATR
  // Stability rule can detect declining vs rising volatility trend.
  if (tabId === "technicals" || tabId === "sentiment") {
    try {
      const atrHistory = await fetch("data/atr-history.json").then((r) => r.json());
      for (const row of rows) if (row.ticker && atrHistory[row.ticker]) row.atr_history = atrHistory[row.ticker];
    } catch { /* file may not exist yet — accumulator will populate over days */ }

    // Per-company source values from TradingView (scrape-technicals-source.mjs).
    // When available for a ticker, we silently overwrite the OHLCV-derived
    // indicator values so the dashboard shows what an analyst would see on
    // TradingView. Missing tickers keep their calculated values.
    try {
      const src = await fetch("data/technicals-source.json").then((r) => r.json());
      const bySlug = src?.companies || {};
      for (const row of rows) {
        const s = row.ticker && bySlug[row.ticker.toUpperCase()];
        if (!s) continue;
        const o = s.oscillators || {};
        const ma = s.moving_averages || {};
        // Track which fields were sourced from TradingView so rule-meta.js
        // can label the Source button accurately per-rule.
        const sourced = new Set();
        if (o.rsi_14   != null) { row.rsi14  = o.rsi_14;   sourced.add("rsi14"); }
        if (o.adx_14   != null) { row.adx14  = o.adx_14;   sourced.add("adx14"); }
        if (ma.ema_50  != null) { row.ema50  = ma.ema_50;  sourced.add("ema50"); }
        if (ma.sma_50  != null) { row.sma50  = ma.sma_50;  sourced.add("sma50"); }
        if (ma.sma_200 != null) { row.sma200 = ma.sma_200; sourced.add("sma200"); }
        if (sourced.size) row._source_tech_fields = sourced;
      }
    } catch { /* source file may not exist yet */ }
  }

  const scored = rows.map(c.score).sort((a, b) => b.totalPoints - a.totalPoints);
  state.cache[tabId] = { rows, scored, meta, filtered: scored };
}

// ---------------- rendering ----------------
function renderAll() {
  // The Strategy and Custom tabs are bespoke renderers with no rules,
  // deferred list, or scored rows -- the shared header pipeline below
  // reads all three and throws on each. They draw their own chrome.
  const c = cfg();
  if (!c || c.active || c.custom) return;
  renderMeta();
  renderStats();
  renderDeferredList();
  renderTopCards();
  applyFilters();   // also renders the table
}

function renderMeta() {
  const c = cfg(); const st = tabState();
  if (!st) return;   // tab not loaded yet — Strategy never fills this cache slot
  const m = st.meta;
  if (m && m.generated_at) {
    $("#meta-updated").textContent = new Date(m.generated_at).toLocaleString("en-IN", {
      day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", timeZone: "Asia/Kolkata",
    });
  }
  $("#meta-count").textContent = `${st.scored.length} companies`;
  // Render a friendly source label instead of dumping the raw saved-screen URL.
  $("#meta-source").textContent = sourceFriendly(c, m);
  // Top-right refresh card: show "Xh ago" + source line (replaces the
  // previous hardcoded "Daily / 06:00 IST" copy with the real freshness).
  const agoEl = $("#refresh-ago"), srcEl = $("#refresh-source");
  if (agoEl && srcEl && m && m.generated_at) {
    agoEl.textContent = relativeTimeFrom(m.generated_at);
    srcEl.textContent = sourceFriendly(c, m);
  } else if (agoEl) {
    agoEl.textContent = "—";
  }
}

// Compact relative-time formatter ("2h ago", "yesterday", "3 days ago").
function relativeTimeFrom(iso) {
  const t = new Date(iso).getTime();
  if (!Number.isFinite(t)) return "—";
  const diffMin = Math.round((Date.now() - t) / 60_000);
  if (diffMin < 1) return "just now";
  if (diffMin < 60) return `${diffMin}m ago`;
  const diffHr = Math.round(diffMin / 60);
  if (diffHr < 24) return `${diffHr}h ago`;
  const diffDay = Math.round(diffHr / 24);
  if (diffDay === 1) return "yesterday";
  if (diffDay < 7) return `${diffDay}d ago`;
  return new Date(iso).toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
}

function sourceFriendly(c, m) {
  if (c.label === "Fundamentals") return "Screener.in saved screen · ₹2,000–12,500 Cr universe";
  if (c.label === "Technicals") return "Yahoo Finance EOD · ₹2,000–12,500 Cr universe";
  if (c.label === "Macro") return "Multi-source · Yahoo + RBI + curated";
  if (c.label === "Sentiment & Liquidity") return "Yahoo + NSE + computed breadth";
  if (c.label === "AI Basket") return `${weightsLabel().names.split(" · ").length}-pillar weighted composite`;
  return c.label;
}

function renderStats() {
  const c = cfg(); const st = tabState();
  const maxCard = $("#stat-max-card");
  // Composite tab overrides the stat cards. The 4-bucket strip below the
  // header carries the rating distribution, so the stat cards here stay
  // intentionally minimal — universe size + pillar weighting.
  if (c.composite && st) {
    const counts = { strong: 0, buy: 0, watch: 0, hardfail: 0, avoid: 0, unrated: 0 };
    for (const s of st.scored) {
      if (s.hardFailed) counts.hardfail++;
      else if (s.unrated) counts.unrated++;
      else if (s.rating === "STRONG BUY") counts.strong++;
      else if (s.rating === "BUY") counts.buy++;
      else if (s.rating === "WATCH") counts.watch++;
      else counts.avoid++;
    }
    const inBasket = counts.strong + counts.buy + counts.watch;
    $("#stat-rules").textContent = `${inBasket}`;
    $("#stat-rules-note").textContent = `In AI Basket of ${st.scored.length}`;
    // Repurpose the third card for the pillar weights mini-strip (was a
    // confusing "Max Score 300" reading before).
    $("#stat-max-label").textContent = "Pillar Weights";
    // Derived, not hardcoded. This card used to print a frozen
    // "40 · 35 · 15 · 5 · 5" that ignored the actual weights entirely, so it
    // silently misreported the model the moment the mix changed. Only pillars
    // carrying weight are listed, so zeroed-out pillars stop being advertised.
    {
      const wNow = state.pillarWeights || composite.PILLAR_WEIGHTS;
      const LBL = { fundamentals: "Fund", technicals: "Tech", macro: "Macro", sentiment: "Sent", liquidity: "Liq" };
      const live = Object.keys(LBL).filter((k) => (wNow[k] || 0) > 0);
      $("#stat-max").innerHTML = `<span class="text-base font-semibold text-slate-700">${live.map((k) => wNow[k]).join(" · ")}</span>`;
      $("#stat-max-note").textContent = live.map((k) => LBL[k]).join(" · ");
    }
    if (maxCard) maxCard.classList.remove("hidden");
    // Rich title — uses innerHTML so the trophy can render properly.
    $("#top-cards-title").innerHTML = `<span class="text-amber-500">🏆</span> AI Basket — Top 10 Picks <span class="ml-2 text-xs font-normal text-slate-500">composite-weighted, hard-fails excluded</span>`;
    return;
  }
  // The Strategy and Custom tabs are bespoke renderers with no rule-count
  // stats block. Reaching here for one of them threw on c.stats.rules and
  // aborted the rest of the render.
  if (!c?.stats) return;
  $("#stat-rules").textContent = c.stats.rules;
  $("#stat-rules-note").textContent = c.stats.rulesNote;
  $("#stat-max-label").textContent = "Max Score";
  $("#stat-max").textContent = c.stats.maxScore;
  $("#stat-max-note").textContent = c.stats.maxNote;
  if (maxCard) maxCard.classList.remove("hidden");
  $("#top-cards-title").textContent = `Top 10 by ${c.label} Score`;
}

function renderDeferredList() {
  const c = cfg();
  const panel = $("#deferred-panel");
  // Composite tab has no "pending data source" concept — hide the panel entirely.
  if (c.composite || !c.deferred?.length) {
    panel?.classList.add("hidden");
    return;
  }
  panel?.classList.remove("hidden");
  $("#deferred-summary").textContent = `${c.deferred.length} parameter${c.deferred.length>1?"s":""} pending data source`;
  $("#deferred-list").innerHTML = c.deferred.map((d) => `
    <div class="flex items-start gap-3 p-3 rounded-lg bg-amber-50 ring-1 ring-amber-100">
      <div class="text-amber-500 text-lg leading-none">⚠</div>
      <div class="flex-1">
        <div class="font-semibold text-slate-900 text-sm">${escapeHtml(d.label)} <span class="text-xs font-normal text-slate-500">· ${escapeHtml(d.category)} · max ${d.max} pts</span></div>
        <div class="text-xs text-slate-600 mt-0.5">${escapeHtml(d.reason)}</div>
      </div>
    </div>
  `).join("");
}

function renderTopCards() {
  const c = cfg(); const st = tabState();
  if (c.composite) return renderCompositeTopCards();
  const top = st.scored.slice(0, 10);
  $("#top-cards").innerHTML = top.map((s, i) => {
    const name = c.name(s.company);
    const { color, initials } = avatarFor(name);
    const tier = s.hardFails.length ? "hardfail" : scoreTier(s.scorePct);
    return `
      <button data-idx="${i}" class="top-card group text-left bg-white hover:shadow-xl hover:-translate-y-1 transition-all duration-200 rounded-2xl p-4 shadow-sm ring-1 ring-slate-100 relative overflow-hidden">
        <div class="absolute top-3 right-3 text-xs font-bold text-slate-400">#${i + 1}</div>
        <div class="flex items-center gap-3 mb-3">
          <div class="w-11 h-11 rounded-xl bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-sm shadow-md">${initials}</div>
          <div class="flex-1 min-w-0">
            <div class="font-semibold text-slate-900 truncate text-sm">${escapeHtml(name)}</div>
            <div class="text-xs text-slate-500">${escapeHtml(c.marketCap(s.company))}</div>
          </div>
        </div>
        <div class="flex items-end justify-between">
          <div>
            <div class="text-3xl font-bold ${tierColor(tier)}">${s.totalPoints}<span class="text-base text-slate-400">/${s.totalMax}</span></div>
            <div class="text-xs text-slate-500 mt-0.5">${tierLabel(tier)}</div>
          </div>
          ${s.hardFails.length ? `<div class="text-rose-500 text-xl" title="Hard fail flag">⚠</div>` : ""}
        </div>
      </button>
    `;
  }).join("");
  $$("#top-cards .top-card").forEach((el) => el.addEventListener("click", () => openDrillDown(top[Number(el.dataset.idx)])));
}

// AI Basket — premium overview: a rating-distribution strip
// (visual breakdown across STRONG / BUY / WATCH / AVOID / FILTERED)
// plus a hero card grid for the top 10 picks. Designed for impact.
function renderCompositeTopCards() {
  const st = tabState();
  const all = st.scored;
  // Split the old single "Hard-Fail" bucket into a genuine fundamental
  // Red flag vs a Below-trend / illiquid filter (a market condition, not a
  // company problem) so the scary number stops conflating the two.
  const counts = { strong: 0, buy: 0, watch: 0, avoid: 0, redFlag: 0, belowTrend: 0, unrated: 0 };
  for (const s of all) {
    if (s.isRedFlag) counts.redFlag++;
    else if (s.hardFailed) counts.belowTrend++;   // only 200-DMA / liquidity filters
    else if (s.unrated) counts.unrated++;
    else if (s.rating === "STRONG BUY") counts.strong++;
    else if (s.rating === "BUY") counts.buy++;
    else if (s.rating === "WATCH") counts.watch++;
    else counts.avoid++;
  }
  const total = all.length;
  const inBasket = counts.strong + counts.buy + counts.watch;

  // Pillar-weight summary — just a compact pill showing current weights
  // with a small "Adjust" button. Full editor lives in a modal so the
  // AI basket header stays clean.
  const w = state.pillarWeights;
  const isDefaultWeights = Object.keys(composite.PILLAR_WEIGHTS)
    .every((k) => w[k] === composite.PILLAR_WEIGHTS[k]);
  const pillarWeightPill = `
    <button id="pillar-weight-btn" type="button" class="inline-flex items-center gap-2 px-3 py-1.5 text-xs rounded-full bg-white ring-1 ring-slate-200 hover:ring-indigo-300 hover:bg-indigo-50 transition shadow-sm">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-indigo-600"><path d="M3 6l3 1m0 0l-3 9a5.002 5.002 0 006.001 0M6 7l3 9M6 7l6-2m6 2l3-1m-3 1l-3 9a5.002 5.002 0 006.001 0M18 7l3 9m-3-9l-6-2m0-2v2m0 16V5m0 16l-3-1m3 1l3-1"/></svg>
      <span class="font-semibold text-slate-700">Weights</span>
      <span class="tabular-nums text-slate-500">${weightsLabel(w).nums.replace(/ /g, "")}</span>
      ${isDefaultWeights ? "" : `<span class="text-[9px] font-bold uppercase tracking-wider text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">Custom</span>`}
    </button>
  `;

  // Compact single-line breakdown — distribution bar + inline chips for
  // each bucket. One row, easy to scan, no scrolling needed before the
  // user reaches the picks below.
  const seg = (pct, klass) => pct > 0 ? `<div class="${klass} h-full" style="width:${pct}%" title="${pct.toFixed(1)}%"></div>` : "";
  const pctOf = (n) => total ? (n / total) * 100 : 0;
  // Semantic grade: positive (green) → caution (amber) → negative (red) →
  // out-of-play (greys). Same colour drives the bar segment and its legend
  // swatch so they always agree.
  const cats = [
    { count: counts.strong,     label: "Strong Buy",  bar: "bg-emerald-500" },
    { count: counts.buy,        label: "Buy",         bar: "bg-teal-500" },
    { count: counts.watch,      label: "Watch",       bar: "bg-amber-400" },
    { count: counts.avoid,      label: "Avoid",       bar: "bg-rose-500" },
    { count: counts.redFlag,    label: "Red flag",    bar: "bg-rose-700" },
    { count: counts.belowTrend, label: "Below trend", bar: "bg-slate-400" },
    { count: counts.unrated,    label: "Unrated",     bar: "bg-slate-300" },
  ];
  const distributionStrip = `
    <div class="rounded-2xl ring-1 ring-slate-200/70 bg-gradient-to-br from-white to-slate-50/60 p-4 mb-5">
      <div class="flex items-center justify-between gap-3 flex-wrap mb-3">
        <div class="flex items-baseline gap-2">
          <span class="text-2xl font-display font-extrabold text-slate-900 tabular-nums leading-none">${total}</span>
          <span class="text-xs text-slate-500">stocks scanned</span>
        </div>
        ${pillarWeightPill}
      </div>
      <div class="flex h-3 overflow-hidden rounded-full ring-1 ring-slate-200 bg-slate-100 shadow-inner" title="${cats.map(c=>`${c.count} ${c.label}`).join(" · ")}">
        ${cats.map((c) => seg(pctOf(c.count), c.bar)).join("")}
      </div>
      <div class="flex flex-wrap items-center gap-x-4 gap-y-2 mt-3 text-xs">
        ${cats.map((c) => `
          <span class="inline-flex items-center gap-1.5">
            <span class="w-2.5 h-2.5 rounded-[3px] ${c.bar}"></span>
            <span class="font-bold text-slate-900 tabular-nums">${c.count}</span>
            <span class="text-slate-500">${c.label}</span>
          </span>
        `).join("")}
        <span class="ml-auto inline-flex items-center gap-1.5 pl-2.5 pr-1 py-1 rounded-full bg-indigo-50 ring-1 ring-indigo-100">
          <span class="text-[11px] text-slate-500">In basket</span>
          <span class="inline-flex items-center justify-center min-w-[22px] px-1.5 py-0.5 rounded-full bg-indigo-600 text-white text-[11px] font-bold tabular-nums">${inBasket}</span>
        </span>
      </div>
    </div>
  `;

  // Premium hero cards for top 10 (basket only — exclude filtered/unrated/AVOID).
  const top = all.filter((s) => !s.hardFailed && !s.unrated && s.rating !== "AVOID").slice(0, 10);
  const heroCards = top.map((s, i) => {
    const co = s.company;
    const name = co.Company || "—";
    const { color, initials } = avatarFor(name);
    const theme = composite.ratingTheme(s.rating);
    const ratingTone = s.rating === "STRONG BUY" ? "text-emerald-700" : s.rating === "BUY" ? "text-blue-700" : s.rating === "WATCH" ? "text-amber-700" : "text-slate-700";
    const ribbon = i === 0 ? `<div class="absolute top-0 left-0 px-2 py-0.5 bg-gradient-to-r from-amber-500 to-orange-500 text-white text-[10px] font-bold uppercase tracking-wider rounded-br-lg shadow">🏆 #1</div>` : `<div class="absolute top-2 right-2 text-xs font-bold text-slate-400">#${i+1}</div>`;
    // Mini pillar bars
    const pillarBars = livePillars().map((k) => {
      const p = s.pillars?.[k];
      const pct = p?.pct ?? 0;
      const tier = pct >= 75 ? "bg-emerald-500" : pct >= 60 ? "bg-blue-500" : pct >= 45 ? "bg-amber-500" : "bg-rose-500";
      return `<div class="flex-1 h-1 rounded-full bg-slate-100 overflow-hidden"><div class="${tier} h-full rounded-full" style="width:${Math.min(100,pct)}%"></div></div>`;
    }).join("");
    return `
      <button data-idx="${i}" class="top-card group text-left bg-white hover:shadow-2xl hover:-translate-y-1 transition-all duration-200 rounded-2xl p-4 shadow-sm ring-1 ring-slate-100 relative overflow-hidden">
        <!-- Top accent line tinted by rating -->
        <div class="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r ${theme.from} ${theme.to}"></div>
        ${ribbon}
        <div class="flex items-center gap-3 mb-3 mt-1">
          <div class="w-11 h-11 rounded-xl bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-sm shadow-md flex-shrink-0">${initials}</div>
          <div class="flex-1 min-w-0">
            <div class="font-bold text-slate-900 truncate text-sm">${escapeHtml(name)}</div>
            <div class="text-[11px] text-slate-500 truncate">${escapeHtml(co.Sector || "")}</div>
          </div>
        </div>
        <div class="flex items-baseline gap-1 mb-1">
          <span class="text-3xl font-bold text-slate-900">${s.composite.toFixed(1)}</span>
          <span class="text-xs text-slate-400">/ 100</span>
        </div>
        <div class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-gradient-to-r ${theme.from} ${theme.to} ${theme.textOn} text-[10px] font-bold uppercase tracking-wider shadow-sm mb-3">
          <span class="w-1 h-1 rounded-full bg-white"></span>
          ${escapeHtml(s.rating)}
        </div>
        <div class="flex items-center gap-0.5" title="Pillar scores: ${weightsLabel().names}">
          ${pillarBars}
        </div>
        <div class="text-[10px] text-slate-400 mt-1 flex justify-between">
          ${weightsLabel().names.split(" · ").map((n) => `<span>${n[0]}</span>`).join("")}
        </div>
      </button>
    `;
  }).join("");

  // Print cover — only renders on @media print. Title + date + key
  // counts in a clean A4-friendly layout.
  const printCover = `
    <div class="print-only print-cover">
      <div style="padding: 30mm 10mm; text-align: center;">
        <div style="font-size: 11px; letter-spacing: 0.25em; text-transform: uppercase; color: #64748b;">VN Smallcap Screener · Basket Brief</div>
        <h1 style="font-size: 42px; font-weight: 800; margin: 12px 0 8px; color: #0f172a;">AI Basket</h1>
        <p style="font-size: 14px; color: #64748b; margin: 0;">${weightsLabel().names.split(" · ").length}-pillar weighted composite · ₹2,000–12,500 Cr universe · ${new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" })}</p>
        <div style="margin: 36px auto 0; max-width: 480px; padding: 24px; border: 1px solid #e2e8f0; border-radius: 16px; background: linear-gradient(135deg, #eef2ff 0%, #ecfeff 100%);">
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px;">
            <div>
              <div style="font-size: 28px; font-weight: 800; color: #10b981;">${counts.strong}</div>
              <div style="font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; color: #64748b;">Strong Buy</div>
            </div>
            <div>
              <div style="font-size: 28px; font-weight: 800; color: #3b82f6;">${counts.buy}</div>
              <div style="font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; color: #64748b;">Buy</div>
            </div>
            <div>
              <div style="font-size: 28px; font-weight: 800; color: #f59e0b;">${counts.watch}</div>
              <div style="font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; color: #64748b;">Watch</div>
            </div>
          </div>
          <div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid #cbd5e1; font-size: 11px; color: #64748b;">
            ${inBasket} of ${total} stocks in basket · ${counts.redFlag} red flag · ${counts.belowTrend} below trend · ${counts.unrated} unrated
          </div>
        </div>
        <p style="font-size: 11px; color: #94a3b8; margin-top: 36px;">For internal use only</p>
      </div>
    </div>
  `;

  $("#top-cards").innerHTML = `
    ${printCover}
    <div class="col-span-full">
      ${distributionStrip}
    </div>
    ${heroCards}
  `;
  $$("#top-cards .top-card").forEach((el) => el.addEventListener("click", () => openDrillDown(top[Number(el.dataset.idx)])));
  $("#pillar-weight-btn")?.addEventListener("click", () => openPillarWeightsModal());
}

// Pillar Weights modal — invoked from the AI basket header pill.
function openPillarWeightsModal() {
  const w = state.pillarWeights;
  const isDefault = Object.keys(composite.PILLAR_WEIGHTS)
    .every((k) => w[k] === composite.PILLAR_WEIGHTS[k]);
  openModal(`
    <div class="px-7 py-6">
      <div class="flex items-start justify-between gap-4 mb-1">
        <div>
          <h2 class="font-display text-xl font-bold text-slate-900">Adjust Pillar Weights</h2>
          <p class="text-sm text-slate-500 mt-1">Set the relative importance of each pillar. Composite scores update across AI Basket and Top Picks.</p>
        </div>
        <button id="modal-close-btn" class="text-slate-400 hover:text-slate-700 text-2xl leading-none">×</button>
      </div>
      <div class="mt-5 grid grid-cols-2 gap-3" style="grid-template-columns: repeat(${Math.min(livePillars(w).length, 5)}, minmax(0, 1fr));">
        ${livePillars(w).map((k) => `
          <label class="block">
            <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1">${k.charAt(0).toUpperCase() + k.slice(1)}</div>
            <div class="flex items-center gap-1 bg-slate-50 border border-slate-200 rounded-lg focus-within:ring-2 focus-within:ring-indigo-500 focus-within:bg-white">
              <input type="number" min="0" max="100" step="1" data-weight-key="${k}" value="${w[k]}"
                class="w-full px-2 py-2 text-base font-bold tabular-nums bg-transparent border-0 focus:outline-none">
              <span class="text-xs text-slate-400 pr-2">%</span>
            </div>
            <div class="text-[10px] text-slate-400 mt-1">Default: ${composite.PILLAR_WEIGHTS[k]}</div>
          </label>
        `).join("")}
      </div>
      <div class="mt-5 flex items-center justify-between gap-2 pt-4 border-t border-slate-100">
        <div class="text-sm text-slate-600">
          Total: <span id="pillar-weight-sum" class="font-bold tabular-nums text-emerald-700">${w.fundamentals+w.technicals+w.macro+w.sentiment+w.liquidity}</span>%
          <span class="text-xs text-slate-400" id="pillar-weight-hint">${isDefault ? "(framework default)" : ""}</span>
        </div>
        <div class="flex items-center gap-2">
          <button id="pillar-weight-reset" type="button" class="px-3 py-2 text-sm font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition">↺ Reset to default</button>
          <button id="pillar-weight-apply" type="button" class="px-5 py-2 text-sm font-semibold text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg shadow-sm transition">Apply</button>
        </div>
      </div>
    </div>
  `, { size: "wide" });
  $("#modal-close-btn")?.addEventListener("click", closeModal);
  $("#modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "modal-overlay") closeModal();
  }, { once: true });
  const sumEl = $("#pillar-weight-sum");
  const hintEl = $("#pillar-weight-hint");
  const inputs = $$('#modal-content input[data-weight-key]');
  function recomputeSum() {
    let s = 0;
    inputs.forEach((el) => { s += Number(el.value) || 0; });
    if (sumEl) {
      sumEl.textContent = String(s);
      sumEl.classList.toggle("text-emerald-700", s === 100);
      sumEl.classList.toggle("text-amber-700", s !== 100);
    }
    if (hintEl) hintEl.textContent = s === 100 ? "" : `(will normalise to 100 on Apply)`;
  }
  inputs.forEach((el) => el.addEventListener("input", recomputeSum));
  function commitAndClose(next) {
    state.pillarWeights = next;
    savePillarWeights(next);
    delete state.cache.composite;
    delete state.cache.topPicks;
    delete state.cache.history;            // banner state + recomputed snapshots both depend on the active weights
    state.compositeBySlug.clear();
    closeModal();
    // Re-route to whichever rating tab the user is on (composite or topPicks).
    switchTab(state.activeTab === "topPicks" ? "topPicks" : "composite");
  }
  $("#pillar-weight-apply")?.addEventListener("click", () => {
    const next = { ...composite.PILLAR_WEIGHTS };   // base on the real defaults, not a frozen copy
    inputs.forEach((el) => { next[el.dataset.weightKey] = Math.max(0, Math.min(100, Number(el.value) || 0)); });
    const s = Object.values(next).reduce((a, b) => a + b, 0);
    if (s > 0 && s !== 100) {
      const k = 100 / s;
      for (const key of Object.keys(next)) next[key] = Math.round(next[key] * k);
    }
    commitAndClose(next);
  });
  $("#pillar-weight-reset")?.addEventListener("click", () => commitAndClose({ ...composite.PILLAR_WEIGHTS }));
}

// Sources modal — single place to see every data source the dashboard
// touches, grouped by domain. Add new sources here as we wire them.
function openSourcesModal() {
  const groups = [
    {
      title: "Market data — live", icon: "💹",
      items: [
        { name: "Yahoo Finance", url: "https://finance.yahoo.com", desc: "Daily OHLCV for the universe, plus Brent crude, USD/INR and India VIX live values" },
        { name: "TradingView", url: "https://in.tradingview.com/", desc: "RSI, MACD, ADX, moving averages — scraped daily so dashboard values match the analyst's screen" },
      ],
    },
    {
      title: "Fundamentals", icon: "📊",
      items: [
        { name: "Screener.in", url: "https://www.screener.in/", desc: "Top ratios, P&L / cash-flow / shareholding series — daily scrape via the saved-screen workflow" },
        { name: "BSE — Annual Reports", url: "https://www.bseindia.com/corporates/AnnualReport_New.aspx", desc: "Auditor opinion + per-company revenue-mix MD&A extraction (weekly routine)" },
      ],
    },
    {
      title: "Shareholding · Insider · Liquidity", icon: "🤝",
      items: [
        { name: "NSE PIT (Insider Trading)", url: "https://www.nseindia.com/companies-listing/corporate-filings-insider-trading", desc: "Reg 7(2) disclosures — daily" },
        { name: "NSE F&O eligible list", url: "https://www.nseindia.com/products-services/equity-derivatives-list-underlyings", desc: "Stocks with active futures + options — routine-refreshed monthly" },
        { name: "NSE BhavCopy", url: "https://www.nseindia.com/all-reports", desc: "Daily delivery percentage from sec_bhavdata_full" },
        { name: "NSE FII / DII flow", url: "https://www.nseindia.com/reports/fii-dii", desc: "Daily institutional buying / selling" },
        { name: "NSE option chain", url: "https://www.nseindia.com/option-chain", desc: "Put-Call Ratio via Firecrawl" },
      ],
    },
    {
      title: "Governance", icon: "⚖️",
      items: [
        { name: "SEBI — Orders + Press Releases", url: "https://www.sebi.gov.in/enforcement/orders.html", desc: "Active proceedings against listed entities — weekly Firecrawl" },
        { name: "Mint / Business Standard / ET — SEBI topic pages", url: "https://www.livemint.com/topic/sebi", desc: "News-driven detection of enforcement actions" },
      ],
    },
    {
      title: "Macro · Economic", icon: "🌍",
      items: [
        { name: "MoSPI", url: "https://www.mospi.gov.in/press-release", desc: "Quarterly GDP + monthly CPI — weekly routine" },
        { name: "RBI — MPC", url: "https://www.rbi.org.in/Scripts/BS_PressReleaseDisplay.aspx", desc: "Rate-cut / hold cycle, policy rate" },
        { name: "CCIL / RBI", url: "https://www.ccilindia.com/", desc: "10-year G-Sec benchmark yield" },
        { name: "PIB / Union Budget", url: "https://pib.gov.in/", desc: "Government capex push, PLI scheme additions" },
        { name: "DPIIT / Ministry of Commerce", url: "https://www.dpiit.gov.in/", desc: "China+1 substitution + trade policy commentary" },
        { name: "NABARD / IMD", url: "https://nabard.org/", desc: "Rural-recovery + monsoon indicators" },
        { name: "MNRE", url: "https://mnre.gov.in/", desc: "Renewable energy capacity additions" },
      ],
    },
  ];
  openModal(`
    <div class="px-7 py-6 max-h-[80vh] overflow-y-auto scrollbar-thin">
      <div class="flex items-start justify-between gap-4 mb-4">
        <div>
          <h2 class="font-display text-2xl font-bold text-slate-900">All Data Sources</h2>
          <p class="text-sm text-slate-500 mt-1">Every external source that feeds this dashboard, grouped by domain.</p>
        </div>
        <button id="modal-close-btn" class="text-slate-400 hover:text-slate-700 text-2xl leading-none">×</button>
      </div>
      <div class="space-y-5">
        ${groups.map((g) => `
          <div>
            <div class="flex items-center gap-2 mb-2">
              <span class="text-base">${g.icon}</span>
              <h3 class="text-xs font-bold uppercase tracking-wider text-indigo-700">${g.title}</h3>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
              ${g.items.map((it) => `
                <a href="${it.url}" target="_blank" rel="noopener" class="group block p-3 rounded-xl bg-slate-50 hover:bg-indigo-50/60 ring-1 ring-slate-200/70 hover:ring-indigo-200 transition">
                  <div class="flex items-start justify-between gap-2 mb-0.5">
                    <span class="font-semibold text-slate-900 text-sm group-hover:text-indigo-700">${it.name}</span>
                    <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-slate-400 group-hover:text-indigo-500 flex-shrink-0 mt-1"><path d="M10 6H6a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4M14 4h6m0 0v6m0-6L10 14"/></svg>
                  </div>
                  <div class="text-[11px] text-slate-500 leading-snug">${it.desc}</div>
                </a>
              `).join("")}
            </div>
          </div>
        `).join("")}
      </div>
      <div class="mt-6 pt-4 border-t border-slate-100 text-[11px] text-slate-400 leading-relaxed">
        Each rule's drill-down has its own Source button that deep-links to the exact page (or PDF) where THAT number lives.
      </div>
    </div>
  `, { size: "magazine" });
  $("#modal-close-btn")?.addEventListener("click", closeModal);
  $("#modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "modal-overlay") closeModal();
  }, { once: true });
}

// ---------------- Top Picks (Hero) ----------------
// Client-facing premium tab. Composite ≥ 75 only, hard-fails excluded.
// Layout philosophy:
//   - Compact hero strip — single row, ~80px tall
//   - Top 3 as a "podium" with showcase cards
//   - Rest as tight horizontal row cards so 8+ picks fit per viewport
//   - Each card's visual signature is a stacked horizontal contribution
//     bar — segment widths = pillar contribution to composite, total
//     fill = composite score itself. One glance: how high + how earned.
function renderTopPicks() {
  const st = state.cache.composite;
  if (!st) return;
  const picks = st.scored.filter((s) => !s.hardFailed && !s.unrated && (s.composite ?? 0) >= 75);
  const meta = st.meta || {};

  if (picks.length === 0) {
    $("#top-picks-content").innerHTML = `
      <button id="tp-back-btn" type="button" class="mb-3 inline-flex items-center gap-1.5 text-sm font-semibold text-indigo-600 hover:text-indigo-700">← Back to AI Basket</button>
      <div class="bg-white rounded-3xl shadow-sm ring-1 ring-slate-200 p-12 text-center">
        <div class="text-6xl mb-4">🔍</div>
        <h2 class="text-2xl font-bold text-slate-900 mb-2">No stocks above 75 right now</h2>
        <p class="text-slate-600">The current macro regime + framework rules have produced no STRONG BUY picks today.<br>Try adjusting your pillar weights or wait for the next refresh.</p>
      </div>`;
    return;
  }

  const avgComposite = (picks.reduce((a, s) => a + (s.composite || 0), 0) / picks.length).toFixed(1);
  const sectors = new Set();
  picks.forEach((s) => { const sec = s.company?.["Sector"] || s.company?.["Broad Industry"]; if (sec) sectors.add(sec); });
  const topScore = picks[0]?.composite ?? 0;

  const PILLAR_KEYS = ["fundamentals", "technicals", "macro", "sentiment", "liquidity"];
  const PILLAR_NAME = { fundamentals: "Fundamentals", technicals: "Technicals", macro: "Macro", sentiment: "Sentiment", liquidity: "Liquidity" };
  // Solid tailwind classes — one per pillar — used by both the per-card
  // contribution bar and the legend below the grid.
  const PILLAR_BAR = {
    fundamentals: "bg-emerald-500",
    technicals:   "bg-indigo-500",
    macro:        "bg-violet-500",
    sentiment:    "bg-amber-500",
    liquidity:    "bg-sky-500",
  };
  const w = state.pillarWeights || composite.PILLAR_WEIGHTS;

  // The contribution bar — visual signature of each pick.
  // Each segment's width = pillar_pct × pillar_weight / 100, summed = composite.
  // Bar total width = 100, so the filled portion equals the composite score.
  function contributionBar(s) {
    const segments = PILLAR_KEYS.map((k) => {
      const pct = s.pillars?.[k]?.pct ?? 0;
      const wt = w[k] ?? composite.PILLAR_WEIGHTS[k];
      return { key: k, contribution: (pct / 100) * wt };
    });
    return `
      <div class="flex h-2.5 w-full overflow-hidden rounded-full bg-slate-100 ring-1 ring-slate-200/70">
        ${segments.map(({ key, contribution }) =>
          contribution > 0
            ? `<div class="${PILLAR_BAR[key]} h-full" style="width:${contribution.toFixed(2)}%" title="${PILLAR_NAME[key]}: ${contribution.toFixed(1)} of ${w[key]} possible"></div>`
            : "").join("")}
      </div>
    `;
  }

  // Hero — slim single-bar header. Picks start ~80px below the tab nav.
  const heroHeader = `
    <div class="relative overflow-hidden rounded-2xl mb-5 print-hide">
      <div class="absolute inset-0 bg-gradient-to-r from-slate-900 via-indigo-900 to-violet-900"></div>
      <div class="absolute inset-0 opacity-40 bg-[radial-gradient(circle_at_10%_50%,rgba(245,158,11,0.35),transparent_45%)]"></div>
      <div class="absolute inset-0 opacity-30 bg-[radial-gradient(circle_at_90%_30%,rgba(16,185,129,0.4),transparent_50%)]"></div>
      <div class="relative px-6 py-4 text-white flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center gap-3">
          <div class="text-2xl leading-none">★</div>
          <div>
            <div class="text-[10px] font-bold uppercase tracking-[0.25em] text-amber-200/80">The Top Picks</div>
            <h1 class="font-display text-2xl sm:text-[28px] font-extrabold leading-none mt-0.5">${picks.length} <span class="text-base font-bold text-white/60">stocks · composite ≥ 75</span></h1>
          </div>
        </div>
        <div class="flex items-center gap-2 text-[11px]">
          <span class="bg-white/10 backdrop-blur px-3 py-1.5 rounded-lg ring-1 ring-white/15"><span class="text-white/60">Top</span> <span class="font-bold tabular-nums text-white">${topScore.toFixed(1)}</span></span>
          <span class="bg-white/10 backdrop-blur px-3 py-1.5 rounded-lg ring-1 ring-white/15"><span class="text-white/60">Avg</span> <span class="font-bold tabular-nums text-white">${avgComposite}</span></span>
          <span class="bg-white/10 backdrop-blur px-3 py-1.5 rounded-lg ring-1 ring-white/15"><span class="text-white/60">Sectors</span> <span class="font-bold tabular-nums text-white">${sectors.size}</span></span>
          ${meta.generated_at ? `<span class="text-white/50 whitespace-nowrap">${relativeTimeFrom(meta.generated_at)}</span>` : ""}
        </div>
      </div>
    </div>
  `;

  // Podium: top 3 as showcase cards
  function podiumCard(s, i) {
    const co = s.company;
    const name = co.Company || "—";
    const { color, initials } = avatarFor(name);
    const sector = co.Sector || co["Broad Industry"] || "";
    const marketCap = co["Market Cap"] || "";
    const medals = ["🥇", "🥈", "🥉"];
    const haloByRank = [
      "from-amber-100/80 via-yellow-50 to-white",
      "from-slate-100 via-slate-50 to-white",
      "from-orange-100/70 via-amber-50 to-white",
    ];
    const accentByRank = [
      "from-amber-400 via-yellow-400 to-orange-400",
      "from-slate-300 via-slate-400 to-slate-500",
      "from-orange-400 via-amber-400 to-amber-500",
    ];
    return `
      <button data-pick-idx="${i}" class="pick-card group relative text-left w-full overflow-hidden rounded-2xl bg-gradient-to-br ${haloByRank[i]} ring-1 ring-slate-200/80 hover:ring-emerald-300 hover:shadow-xl hover:-translate-y-0.5 transition-all">
        <div class="absolute top-0 inset-x-0 h-1 bg-gradient-to-r ${accentByRank[i]}"></div>
        <div class="p-5">
          <div class="flex items-start justify-between gap-4 mb-4">
            <div class="flex items-center gap-3 min-w-0 flex-1">
              <div class="text-3xl leading-none">${medals[i]}</div>
              <div class="w-12 h-12 rounded-xl bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-base shadow-md flex-shrink-0">${initials}</div>
              <div class="min-w-0">
                <div class="font-display font-bold text-slate-900 text-base leading-tight truncate" title="${escapeHtml(name)}">${escapeHtml(name)}</div>
                <div class="text-[11px] text-slate-500 truncate mt-0.5">${escapeHtml(sector)}${marketCap ? ` · ${escapeHtml(marketCap)}` : ""}</div>
              </div>
            </div>
            <div class="text-right flex-shrink-0">
              <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold leading-none mb-1">Composite</div>
              <div class="text-[40px] font-extrabold tabular-nums leading-none bg-gradient-to-br from-emerald-600 to-teal-600 bg-clip-text text-transparent">${s.composite.toFixed(1)}</div>
            </div>
          </div>
          <div class="mb-1">${contributionBar(s)}</div>
          <div class="flex items-center justify-between mt-3">
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 ring-1 ring-emerald-200 text-[10px] font-bold uppercase tracking-wider">
              <span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>STRONG BUY
            </span>
            <span class="text-[11px] font-semibold text-slate-400 group-hover:text-emerald-600 transition-colors">View details →</span>
          </div>
        </div>
      </button>
    `;
  }

  // Compact horizontal row card — one full-width row per pick. Heavy
  // info-density: 7-8 picks visible per viewport on a desktop monitor.
  function rowCard(s, i) {
    const co = s.company;
    const name = co.Company || "—";
    const { color, initials } = avatarFor(name);
    const sector = co.Sector || co["Broad Industry"] || "";
    const marketCap = co["Market Cap"] || "";
    return `
      <button data-pick-idx="${i}" class="pick-card group w-full text-left flex items-center gap-3 sm:gap-4 px-4 py-3 rounded-xl bg-white ring-1 ring-slate-200/80 hover:ring-emerald-300 hover:shadow-md hover:-translate-y-px transition-all">
        <div class="w-8 flex-shrink-0 text-center text-xs font-bold text-slate-400 tabular-nums tracking-wider">#${String(i + 1).padStart(2, "0")}</div>
        <div class="w-10 h-10 rounded-lg bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-xs shadow-sm flex-shrink-0">${initials}</div>
        <div class="min-w-0 flex-1">
          <div class="font-display font-bold text-slate-900 text-sm leading-tight truncate" title="${escapeHtml(name)}">${escapeHtml(name)}</div>
          <div class="text-[11px] text-slate-500 truncate mt-0.5">${escapeHtml(sector)}${marketCap ? ` · ${escapeHtml(marketCap)}` : ""}</div>
        </div>
        <div class="hidden md:block flex-1 max-w-[260px]">${contributionBar(s)}</div>
        <div class="text-right flex-shrink-0 w-16">
          <div class="text-2xl font-extrabold tabular-nums leading-none bg-gradient-to-br from-emerald-600 to-teal-600 bg-clip-text text-transparent">${s.composite.toFixed(1)}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold mt-0.5">/100</div>
        </div>
        <span class="hidden sm:inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 ring-1 ring-emerald-200 text-[9px] font-bold uppercase tracking-wider flex-shrink-0">STRONG BUY</span>
        <span class="text-slate-300 group-hover:text-emerald-600 transition-colors flex-shrink-0">→</span>
      </button>
    `;
  }

  const podium = picks.slice(0, 3).map((s, i) => podiumCard(s, i)).join("");
  const rest = picks.slice(3).map((s, i) => rowCard(s, i + 3)).join("");

  // Legend tied to the contribution bar's pillar colors.
  const legend = `
    <div class="mt-5 flex flex-wrap items-center justify-center gap-x-4 gap-y-1 text-[11px] text-slate-500">
      <span class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mr-1">Pillar mix:</span>
      ${PILLAR_KEYS.map((k) => `
        <span class="inline-flex items-center gap-1.5">
          <span class="w-2.5 h-2.5 rounded-full ${PILLAR_BAR[k]}"></span>${PILLAR_NAME[k]}
        </span>
      `).join("")}
    </div>
  `;

  $("#top-picks-content").innerHTML = `
    <button id="tp-back-btn" type="button" class="mb-3 inline-flex items-center gap-1.5 text-sm font-semibold text-indigo-600 hover:text-indigo-700 print-hide">← Back to AI Basket</button>
    ${heroHeader}
    ${picks.length > 0 ? `
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-3 mb-3">
        ${podium}
      </div>` : ""}
    ${picks.length > 3 ? `
      <div class="space-y-1.5">
        ${rest}
      </div>` : ""}
    ${legend}
  `;
  // Top Picks lives outside any tab config — these are composite-scored
  // results, not generic c.score outputs — so route clicks straight to
  // the composite drill (openDrillDown would consult cfg() for the
  // current tab, which has no name/rules for Top Picks and silently
  // crashes).
  $$("#top-picks-content .pick-card").forEach((el) => el.addEventListener("click", () => openCompositeDrill(picks[Number(el.dataset.pickIdx)])));
}

// ---------------- History (predictions performance) ----------------
// Loads the snapshot manifest from public/data/snapshots/index.json,
// fetches every dated snapshot file, and reconstructs per-ticker
// timelines. For each ticker that was ever rated STRONG BUY, we anchor
// at the FIRST such day and measure return to today's close — that's
// the "we said STRONG BUY at ₹444 → today ₹555" story client wants.
// Loads + caches the snapshot trail (manifest + per-day JSON), benchmark
// closes, LKP manual picks, and the out-of-coverage OHLCV injection step.
// Called by both renderHistory() and renderActive() so they share the
// same cache. Throws on failure — caller decides how to render the
// empty / error state.
async function ensureHistoryCache() {
  if (state.cache.history) return state.cache.history;
  const idx = await fetch("data/snapshots/index.json").then((r) => {
    if (!r.ok) throw new Error("manifest missing");
    return r.json();
  });
  if (!idx.dates?.length) throw new Error("no snapshot dates");
  // Where rebuilt history stops and real forward tracking starts. Written by
  // backfill-from-history.mjs; absent on a repo that was never backfilled.
  state.cache.liveFrom = idx.live_from || null;
  const snapshots = await Promise.all(idx.dates.map((d) =>
    fetch(`data/snapshots/${d}.json`).then((r) => r.json())
  ));
  const benchmark = await fetch("data/benchmark-history.json")
    .then((r) => (r.ok ? r.json() : null))
    .catch(() => null);
  // No manual basket on this dashboard. The comparison here is AI vs
  // benchmark, not AI vs an analyst's picks, so there is nothing to load.
  // Kept as an explicit null rather than deleted because a lot of
  // downstream code already branches on `lkp` being absent; making it
  // permanently null exercises those existing paths instead of inventing
  // new ones.
  const lkp = null;
  const ooc = await fetch("data/out-of-coverage-history.json")
    .then((r) => (r.ok ? r.json() : null))
    .catch(() => null);
  // Live intraday feed (Munshot) — current price + day high/low per basket
  // ticker. Powers live prices and intraday target/SL touch detection.
  const livePricesRaw = await fetch("data/live-prices.json")
    .then((r) => (r.ok ? r.json() : null))
    .catch(() => null);
  const livePrices = livePricesRaw?.prices || {};

  // Inject synthetic stock entries for out-of-coverage tickers (LKP picks
  // below our ~Rs 12,500 Cr universe, e.g. ELECON). Once they look like
  // ordinary snapshot stocks with rating: null, every downstream lookup
  // resolves them naturally — and the active-basket sim filters them out
  // via the rating === "STRONG BUY" check.
  const oocTickers = ooc?.tickers || {};
  if (Object.keys(oocTickers).length) {
    for (const snap of snapshots) {
      for (const [ticker, closes] of Object.entries(oocTickers)) {
        const close = closes[snap.date];
        if (close == null) continue;
        if (snap.stocks.some((s) => s.ticker === ticker)) continue;
        snap.stocks.push({
          ticker, name: ticker, close,
          composite: null, rating: null,
          hardFailed: false, dataComplete: false,
          _outOfCoverage: true,
        });
      }
    }
    const promote = (arr) => {
      for (const p of (arr || [])) {
        if (p.in_universe) continue;
        const t = (p.ticker || p.selection || "").toString().trim().toUpperCase();
        if (t && oocTickers[t]) {
          p.ticker = t;
          p.in_universe = true;
          p.viaOoc = true;
        }
      }
    };
    if (lkp) {
      promote(lkp.picks);
      for (const m of Object.values(lkp.picksByMonth || {})) promote(m);
    }
  }
  state.cache.history = { idx, snapshots, benchmark, lkp, ooc };

  // If the analyst opted in to recomputing history with their custom
  // pillar weights, mutate the cached snapshots now. Reverting clears the
  // cache so the next render fetches fresh v1 snapshots.
  if (state.recomputeHistory && !weightsMatchDefault(state.pillarWeights)) {
    recomputeSnapshotsWithWeights(snapshots, state.pillarWeights);
    state.cache.history.recomputedWith = { ...state.pillarWeights };
  }

  // Per-ticker timeline of points {date, close, composite, rating, pillars}
  // + identity (name, sector). Lives in cache so any tab can resolve a
  // drill click — without this, clicking a row in the Active tab without
  // first visiting History returned null because byTicker was History-only.
  const byTicker = new Map();
  for (const snap of snapshots) {
    for (const s of snap.stocks) {
      if (!s.ticker) continue;
      const t = byTicker.get(s.ticker) || { ticker: s.ticker, name: s.name, sector: s.sector, slug: s.slug, points: [] };
      if (s.name) t.name = s.name;
      if (s.sector) t.sector = s.sector;
      if (s.slug) t.slug = s.slug;
      t.points.push({ date: snap.date, close: s.close, composite: s.composite, rating: s.rating, pillars: s.pillars || null });
      byTicker.set(s.ticker, t);
    }
  }
  // Today's close per ticker (most recent snapshot, fall back to any
  // earlier non-null close).
  const todayClose = {};
  const todaySnap = snapshots[snapshots.length - 1];
  for (const s of todaySnap.stocks) if (typeof s.close === "number") todayClose[s.ticker] = s.close;
  // Overlay the live price where we have one, so displays show the current
  // traded price rather than the end-of-day snapshot close.
  for (const [t, p] of Object.entries(livePrices)) {
    if (p && typeof p.current === "number") todayClose[t] = p.current;
  }
  state.cache.history.livePrices = livePrices;
  state.cache.history.todayClose = todayClose;
  // Kept so the UI can tell "nothing moved" apart from "nothing arrived".
  state.cache.history.livePricesGeneratedAt = livePricesRaw?.generated_at || null;
  // LKP picks indexed by ticker so manual-row clicks resolve to the
  // client-entry framing (with targets/SL) rather than the snapshot trail.
  const lkpForCache = lkpOverride() || lkp;
  const lkpPicksList = lkpForCache ? buildLkpPickList(lkpForCache.picks || [], byTicker, todayClose) : [];
  const lkpPicksByTicker = new Map(lkpPicksList.filter((p) => p.ticker).map((p) => [p.ticker, p]));
  const cohortAnchor = lkpAnchorDate(lkp, snapshots);
  Object.assign(state.cache.history, { byTicker, todayClose, lkpPicksByTicker, cohortAnchor });
  return state.cache.history;
}

async function renderHistory() {
  const host = $("#history-content");
  if (!host) return;
  host.innerHTML = `<div class="bg-white rounded-2xl ring-1 ring-slate-200 p-10 text-center text-slate-500 text-sm">Loading history…</div>`;

  try {
    await ensureHistoryCache();
  } catch (e) {
    host.innerHTML = renderHistoryEmpty(e.message);
    return;
  }
  const { idx, snapshots, benchmark } = state.cache.history;
  // Uploaded basket (localStorage) takes precedence over the committed file so
  // an Excel upload previews instantly; cleared via "Reset to published".
  const lkp = lkpOverride() || state.cache.history.lkp;

  // Per-ticker timeline of points {date, close, composite, rating, pillars}
  // plus identity (name, sector). Pillars are kept for forensics decomp
  // in the drill modal — composite delta = sum of per-pillar weighted deltas.
  const byTicker = new Map();
  for (const snap of snapshots) {
    for (const s of snap.stocks) {
      if (!s.ticker) continue;
      const t = byTicker.get(s.ticker) || { ticker: s.ticker, name: s.name, sector: s.sector, slug: s.slug, points: [] };
      // Always keep the freshest non-null name / sector
      if (s.name) t.name = s.name;
      if (s.sector) t.sector = s.sector;
      if (s.slug) t.slug = s.slug;
      t.points.push({ date: snap.date, close: s.close, composite: s.composite, rating: s.rating, pillars: s.pillars || null });
      byTicker.set(s.ticker, t);
    }
  }

  // Today's close per ticker (from the most recent snapshot, falling
  // back to any earlier non-null close if today is missing).
  const todayDate = idx.dates[idx.dates.length - 1];
  const todayClose = {};
  const todayMap = snapshots[snapshots.length - 1];
  for (const s of todayMap.stocks) if (typeof s.close === "number") todayClose[s.ticker] = s.close;

  // Benchmark series lookup — Nifty close on any date, falling back to
  // the most recent close on or before the requested date.
  const niftyClosesByDate = benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || null;
  const niftyDatesSorted = niftyClosesByDate ? Object.keys(niftyClosesByDate).sort() : null;
  function niftyOn(date) {
    if (!niftyClosesByDate) return null;
    if (niftyClosesByDate[date] != null) return niftyClosesByDate[date];
    let last = null;
    for (const d of niftyDatesSorted) { if (d <= date) last = niftyClosesByDate[d]; else break; }
    return last;
  }

  // For each ticker that had at least one STRONG BUY day, pin the FIRST
  // such day as the "we said" anchor and compute realized return + alpha.
  const picks = [];
  for (const t of byTicker.values()) {
    const firstSB = t.points.find((p) => p.rating === "STRONG BUY" && typeof p.close === "number");
    if (!firstSB) continue;
    const now = todayClose[t.ticker];
    if (typeof now !== "number") continue;
    const ret = (now / firstSB.close - 1) * 100;
    const days = daysBetween(firstSB.date, todayDate);
    // Latest snapshot point (for "current rating" + pillars-now forensics)
    let lastPoint = null;
    for (let i = t.points.length - 1; i >= 0; i--) {
      if (t.points[i].rating) { lastPoint = t.points[i]; break; }
    }
    // Alpha vs Nifty over the same window
    const nStart = niftyOn(firstSB.date), nEnd = niftyOn(todayDate);
    const niftyReturn = (nStart && nEnd) ? (nEnd / nStart - 1) * 100 : null;
    const alpha = niftyReturn != null ? ret - niftyReturn : null;
    picks.push({
      ticker: t.ticker, name: t.name, sector: t.sector,
      firstSBDate: firstSB.date,
      firstSBClose: firstSB.close,
      firstSBComposite: firstSB.composite,
      firstSBPillars: firstSB.pillars || null,
      todayClose: now,
      todayPillars: lastPoint?.pillars || null,
      ret, days,
      niftyReturn, alpha,
      currentRating: lastPoint?.rating || null,
      points: t.points,
    });
  }

  picks.sort((a, b) => b.ret - a.ret);

  // --- Performance Tracker (cohort-style, founder-spec) ---
  // Three views drive the anchor + re-lock behaviour:
  //   static  → select from prior month-end snapshot, COST BASIS at
  //             the first held-month snapshot. Per founder's call:
  //             "assume client bought June 1" — chart day-0 = June 1,
  //             not May 31. AI top 7 still picked from May 31's
  //             composite ranking.
  //   monthly → anchor + cost basis at the client upload date
  //             (handles mid-month uploads).
  //   weekly  → anchor at upload date, AI basket re-locks each Mon.
  const view = state.cohortView;
  let cohortAnchor, cohortSelection;
  if (view === "static") {
    cohortSelection = staticAnchorDate(snapshots);
    // Cost basis = first snapshot of the IST-today calendar month after
    // the selection date. Falls back to the selection date itself
    // (first day of new month before snapshot lands) so the tracker
    // still renders something.
    const heldMonth = istTodayDate().slice(0, 7);
    const firstHeld = snapshots.find((s) => s.date.slice(0, 7) === heldMonth && s.date > (cohortSelection || ""));
    cohortAnchor = firstHeld ? firstHeld.date : cohortSelection;
  } else {
    cohortAnchor = lkpAnchorDate(lkp, snapshots);
    cohortSelection = cohortAnchor;
  }
  const buildMode = view === "weekly" ? "weekly" : "monthly";
  const cohort = buildCohort(snapshots, cohortAnchor, buildMode, { selectionDate: cohortSelection });
  if (cohort) cohort.anchorMode = view;             // for the subtitle copy

  // Manual basket source: prefer month-keyed picksByMonth (each month
  // locks its own client basket); fall back to top-level lkp.picks
  // (legacy single-basket file).
  //
  // Held month resolves PER VIEW:
  //   static  → IST today's calendar month (held month differs from
  //             the prior month-end anchor)
  //   monthly → cohortAnchor's month (= client's upload month).
  //             When client uploads July's picks the anchor moves to
  //             July, picksByMonth lookup follows.
  //   weekly  → same as monthly (upload-month based)
  const manualMonthKey = view === "static"
    ? istTodayDate().slice(0, 7)
    : (cohortAnchor ? cohortAnchor.slice(0, 7) : null);
  const manualPicks = (manualMonthKey && lkp?.picksByMonth?.[manualMonthKey])
    || lkp?.picks
    || [];
  const cohortSeriesData = cohort ? buildCohortSeries(cohort, manualPicks, niftyOn) : null;
  const segCount = cohort?.segments?.length || 0;
  const selectedSegIdx = (state.cohortSegmentIdx != null && state.cohortSegmentIdx >= 0 && state.cohortSegmentIdx < segCount)
    ? state.cohortSegmentIdx
    : Math.max(0, segCount - 1);
  const cohortTracker = cohort
    ? renderCohortTracker(cohort, cohortSeriesData, state.cohortView, selectedSegIdx)
    : "";
  const prevMonthCard = "";        // dormant — restored once we have multi-period history

  // No STRONG BUYs yet? Render the cohort tracker on its own so a flat
  // period (like today) doesn't hide the founder's feature.
  if (picks.length === 0) {
    if (cohort) {
      // Same cohort-row lookup cache as the full-render branch — mirror
      // the manualPicks list (picksByMonth-aware) used by the table.
      const earlyLkpPicks = buildLkpPickList(manualPicks, byTicker, todayClose);
      state.cache.history.byTicker = byTicker;
      state.cache.history.todayClose = todayClose;
      state.cache.history.lkpPicksByTicker = new Map(earlyLkpPicks.filter((p) => p.ticker).map((p) => [p.ticker, p]));
      state.cache.history.cohortAnchor = cohortAnchor;
      // No STRONG BUYs yet, but cohort exists — show cohort tracker
      // + accuracy/history sub-switch (Accuracy still works because
      // it pulls picks from cohort segments, not from STRONG BUY history).
      const historyView = state.historyView;
      const accuracyData = buildAccuracyData(cohort, manualPicks, snapshots);
      const subViewSwitch = renderHistoryViewSwitch(historyView, accuracyData, view);
      const subBody = historyView === "accuracy"
        ? renderAccuracyView(accuracyData)
        : renderHistoryEmpty(`Snapshots loaded (${idx.dates.length} days) but no STRONG BUY picks have been recorded yet.`);
      const weightsBanner = renderPillarWeightsBanner();
      host.innerHTML = weightsBanner + subViewSwitch + cohortTracker + prevMonthCard + subBody;
      wireCohortHandlers(cohortSeriesData);
      wireHistorySubViewSwitch();
      wirePillarWeightsBanner();
    } else {
      host.innerHTML = renderHistoryEmpty(`Snapshots loaded (${idx.dates.length} days) but no STRONG BUY picks have been recorded yet.`);
    }
    return;
  }

  const winners = picks.filter((p) => p.ret > 0);
  const avg = picks.reduce((a, b) => a + b.ret, 0) / picks.length;
  const best = picks[0];
  const dateRange = idx.dates.length > 1 ? `${idx.dates[0]} → ${idx.dates[idx.dates.length - 1]}` : idx.dates[0];

  // Portfolio ₹-backtest removed — funds judge on relative % / alpha, not the
  // absolute "₹X invested → ₹Y today" simulation. Avg return %, win rate, and
  // per-pick alpha vs Nifty carry the relative story.

  // --- Hero strip ---
  const heroHeader = `
    <div class="relative overflow-hidden rounded-2xl mb-4 print-hide">
      <div class="absolute inset-0 bg-gradient-to-r from-slate-900 via-indigo-900 to-purple-900"></div>
      <div class="absolute inset-0 opacity-40 bg-[radial-gradient(circle_at_15%_50%,rgba(34,211,238,0.35),transparent_45%)]"></div>
      <div class="absolute inset-0 opacity-30 bg-[radial-gradient(circle_at_85%_30%,rgba(16,185,129,0.4),transparent_50%)]"></div>
      <div class="relative px-6 py-5 text-white">
        <div class="flex flex-wrap items-start justify-between gap-3 mb-4">
          <div>
            <div class="text-[10px] font-bold uppercase tracking-[0.25em] text-cyan-200/80">Predictions · Performance</div>
            <h1 class="font-display text-2xl sm:text-3xl font-extrabold leading-tight mt-1">${picks.length} past STRONG BUY picks ${avg >= 0 ? "earning" : "down"} <span class="${avg >= 0 ? "text-emerald-300" : "text-rose-300"}">${avg >= 0 ? "+" : ""}${avg.toFixed(1)}%</span> on average</h1>
            <p class="text-white/70 text-xs mt-1">From ${dateRange} · realized return from first STRONG BUY date → today's close</p>
          </div>
        </div>
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
          ${heroStat("Picks", picks.length, `${idx.dates.length} day window`)}
          ${heroStat("Avg return", `${avg >= 0 ? "+" : ""}${avg.toFixed(1)}%`, "since first STRONG BUY")}
          ${heroStat("Win rate", `${Math.round(winners.length / picks.length * 100)}%`, `${winners.length} of ${picks.length} up`)}
          ${heroStat("Best pick", `${best.ret >= 0 ? "+" : ""}${best.ret.toFixed(1)}%`, best.name)}
        </div>
      </div>
    </div>
  `;

  // --- Table of past picks (one row per ticker, sortable visually) ---
  const rows = picks.map((p, i) => {
    const { color, initials } = avatarFor(p.name || "—");
    const ratingChipCls = composite.ratingClass(p.currentRating || "—");
    const retCls = p.ret >= 0 ? "text-emerald-700" : "text-rose-700";
    const retBg = p.ret >= 0 ? "bg-emerald-50 ring-emerald-100" : "bg-rose-50 ring-rose-100";
    const alphaCls = p.alpha == null ? "" : p.alpha >= 0 ? "text-emerald-700" : "text-rose-700";
    return `
      <button data-pick="${i}" class="hist-row w-full text-left grid grid-cols-12 items-center gap-3 px-4 py-3 rounded-xl bg-white ring-1 ring-slate-200/80 hover:ring-indigo-300 hover:shadow-md hover:-translate-y-px transition-all">
        <div class="col-span-1 sm:col-span-1 flex items-center gap-2">
          <span class="text-xs font-bold text-slate-400 tabular-nums">#${String(i + 1).padStart(2, "0")}</span>
        </div>
        <div class="col-span-5 sm:col-span-4 flex items-center gap-3 min-w-0">
          <div class="w-9 h-9 rounded-lg bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-xs shadow-sm flex-shrink-0">${initials}</div>
          <div class="min-w-0">
            <div class="font-display font-bold text-slate-900 text-sm leading-tight truncate" title="${escapeHtml(p.name || "")}">${escapeHtml(p.name || "—")}</div>
            <div class="text-[11px] text-slate-500 truncate">${escapeHtml(p.sector || "")}</div>
          </div>
        </div>
        <div class="hidden sm:block col-span-3">
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400">We said STRONG BUY</div>
          <div class="text-xs text-slate-700 mt-0.5">
            <span class="font-semibold">${p.firstSBDate}</span>
            <span class="text-slate-400">·</span>
            <span>₹${formatPrice(p.firstSBClose)}</span>
            ${p.firstSBComposite != null ? `<span class="text-slate-400">·</span><span class="font-bold tabular-nums">${p.firstSBComposite.toFixed(1)}</span>` : ""}
          </div>
        </div>
        <div class="hidden sm:block col-span-2">
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Today</div>
          <div class="text-xs text-slate-700 mt-0.5">
            <span>₹${formatPrice(p.todayClose)}</span>
            <span class="text-slate-400">·</span>
            <span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap ${ratingChipCls}">${escapeHtml(p.currentRating || "—")}</span>
          </div>
        </div>
        <div class="col-span-6 sm:col-span-2 flex flex-col items-end gap-0.5">
          <div class="flex items-center gap-2">
            <span class="inline-flex items-center px-2.5 py-1 rounded-lg ring-1 text-sm font-extrabold tabular-nums ${retCls} ${retBg}">${p.ret >= 0 ? "+" : ""}${p.ret.toFixed(1)}%</span>
            <span class="text-[10px] text-slate-400 tabular-nums whitespace-nowrap">${p.days}d</span>
          </div>
          ${p.alpha != null
            ? `<span class="text-[10px] font-semibold tabular-nums ${alphaCls}">α ${p.alpha >= 0 ? "+" : ""}${p.alpha.toFixed(2)}% vs Nifty</span>`
            : ""}
        </div>
      </button>
    `;
  }).join("");

  const lkpCard = "";   // manual-basket card removed — AI vs benchmark only

  // Cache for cohort-row click handlers (Performance Tracker drill).
  // Crucially the click cache mirrors the SAME `manualPicks` list the
  // cohort table renders (picksByMonth when present, else lkp.picks),
  // so covered Manual rows always resolve to the right entry/targets
  // pick instead of falling through.
  const cohortLkpPicks = buildLkpPickList(manualPicks, byTicker, todayClose);
  state.cache.history.byTicker = byTicker;
  state.cache.history.todayClose = todayClose;
  state.cache.history.lkpPicksByTicker = new Map(cohortLkpPicks.filter((p) => p.ticker).map((p) => [p.ticker, p]));
  state.cache.history.cohortAnchor = cohortAnchor;

  // Below the Performance Tracker, a prominent tabbed header lets the
  // analyst flip between History (past STRONG BUYs + LKP card) and
  // Accuracy (target / stop-loss hit tracker). Bell-icon dropdown on
  // the right surfaces every recent hit with a badge for today's count.
  const historyView = state.historyView;
  const accuracyData = buildAccuracyData(cohort, manualPicks, snapshots);
  const accuracyView = renderAccuracyView(accuracyData);
  const subViewSwitch = renderHistoryViewSwitch(historyView, accuracyData, view);
  const weightsBanner = renderPillarWeightsBanner();
  const historySubViewHtml = historyView === "accuracy"
    ? accuracyView
    : `
      ${heroHeader}
      <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 sm:p-5">
        <div class="flex items-center justify-between mb-3">
          <h2 class="font-display font-bold text-slate-900 text-base">Past STRONG BUYs · realized return</h2>        </div>
        <div class="space-y-1.5">${rows}</div>
      </div>
      ${lkpCard}
    `;

  host.innerHTML = `
    ${weightsBanner}
    ${subViewSwitch}
    ${cohortTracker}
    ${prevMonthCard}
    ${historySubViewHtml}
  `;
  $$("#history-content .hist-row").forEach((el) => el.addEventListener("click", () => openHistoryDrill(picks[Number(el.dataset.pick)])));
  $$("#history-content .lkp-row").forEach((el) => el.addEventListener("click", () => openHistoryDrill(lkpPicks[Number(el.dataset.lkp)])));
  // LKP Excel-upload controls
  $("#lkp-upload-btn")?.addEventListener("click", () => $("#lkp-file-input")?.click());
  $("#lkp-file-input")?.addEventListener("change", (e) => { const f = e.target.files?.[0]; if (f) handleLkpExcelUpload(f); e.target.value = ""; });
  $("#lkp-download-btn")?.addEventListener("click", downloadLkpJson);
  $("#lkp-reset-btn")?.addEventListener("click", () => { clearLkpOverride(); renderHistory(); });
  wireHistorySubViewSwitch();
  wirePillarWeightsBanner();
  // Performance Tracker controls (view toggle, week pills, hover)
  wireCohortHandlers(cohortSeriesData);

}

// Cohort tracker — segment pills (which week's AI basket to show in the
// table) + chart hover crosshair. View toggle (Static/Monthly/Weekly)
// lives in the History/Accuracy header now and is wired by
// wireHistorySubViewSwitch.
function wireCohortHandlers(seriesData) {
  $$("#history-content [data-seg]").forEach((btn) => btn.addEventListener("click", () => {
    state.cohortSegmentIdx = Number(btn.dataset.seg);
    renderHistory();
  }));
  // Cohort rows → existing drill modal (price line + rating tape +
  // forensics). Manual rows reuse the LKP pick (client-entry framing
  // with targets/SL); AI rows build a pick from byTicker (STRONG BUY
  // framing if any STRONG BUY history, else cohort-entry framing).
  $$("#history-content [data-cohort-row]").forEach((el) => el.addEventListener("click", () => {
    const ticker = el.dataset.ticker;
    const side = el.dataset.cohortSide;
    const segAnchor = el.dataset.segAnchor || null;        // weekly mode: per-segment anchor
    if (!ticker) return;
    const pick = buildCohortClickPick(ticker, side, segAnchor);
    if (pick) openHistoryDrill(pick);
  }));
  if (seriesData) setupCohortHover(seriesData);
}

// Resolve a cohort-table row click to a pick object that openHistoryDrill
// can render. Manual rows prefer the LKP pick (with entry / targets / SL
// framing) when its today-close is usable; otherwise fall back to a
// snapshot-trail pick so the drill still opens cleanly. AI rows always
// build from the snapshot trail, anchored at the first STRONG BUY day
// if any, else at the segment anchor (= selected week start in weekly
// mode; the cohort upload date in monthly).
function buildCohortClickPick(ticker, side, segAnchor) {
  const cache = state.cache.history || {};
  const lkpByTicker = cache.lkpPicksByTicker;
  if (side === "manual" && lkpByTicker?.has(ticker)) {
    const lkpPick = lkpByTicker.get(ticker);
    // Only return the LKP pick if it has usable data — otherwise the
    // drill modal would crash on null .toFixed / null close. Fall
    // through to the snapshot-trail builder below.
    if (lkpPick && lkpPick.ret != null && lkpPick.todayClose != null) return lkpPick;
  }

  const byTicker = cache.byTicker;
  const tk = byTicker?.get(ticker);
  // Was < 2, which meant that on the first day of tracking -- exactly when
  // a new cohort is locked in -- every row in the AI picks list did nothing
  // at all when clicked. No error, no message. A single point is enough to
  // show entry price, composite and pillar breakdown; only the return chart
  // needs two, and that degrades on its own.
  if (!tk || !Array.isArray(tk.points) || tk.points.length < 1) return null;

  // When a segment anchor is supplied (Accuracy / Performance row click),
  // the row's target / SL / return were all computed from the close at
  // that exact date — use the same close as the drill modal's anchor so
  // the chart overlay matches the row's numbers. Without an anchor we
  // fall back to the first STRONG BUY across the trail (used by hist-row
  // clicks where the row IS that first STRONG BUY pick).
  let anchorPoint = null;
  let isCohortLookup = false;
  if (segAnchor) {
    anchorPoint = tk.points.find((p) => p.date === segAnchor && typeof p.close === "number")
      || tk.points.find((p) => p.date >= segAnchor && typeof p.close === "number");
    isCohortLookup = !anchorPoint || anchorPoint.rating !== "STRONG BUY";
  }
  if (!anchorPoint) {
    const firstSB = tk.points.find((p) => p.rating === "STRONG BUY" && typeof p.close === "number");
    const cohortAnchor = cache.cohortAnchor;
    anchorPoint = firstSB
      || (cohortAnchor && tk.points.find((p) => p.date >= cohortAnchor && typeof p.close === "number"))
      || tk.points.find((p) => typeof p.close === "number");
    isCohortLookup = !firstSB;
  }
  if (!anchorPoint) return null;

  let todayPoint = null;
  for (let i = tk.points.length - 1; i >= 0; i--) {
    if (typeof tk.points[i].close === "number") { todayPoint = tk.points[i]; break; }
  }
  if (!todayPoint) todayPoint = anchorPoint;
  let currentRating = null;
  for (let i = tk.points.length - 1; i >= 0; i--) {
    if (tk.points[i].rating) { currentRating = tk.points[i].rating; break; }
  }

  return {
    ticker, name: tk.name, sector: tk.sector,
    firstSBDate: anchorPoint.date,
    firstSBClose: anchorPoint.close,
    firstSBComposite: anchorPoint.composite,
    firstSBPillars: anchorPoint.pillars,
    todayClose: todayPoint.close,
    todayPillars: todayPoint.pillars,
    currentRating,
    ret: (anchorPoint.close && todayPoint.close) ? (todayPoint.close / anchorPoint.close - 1) * 100 : null,
    days: daysBetween(anchorPoint.date, todayPoint.date),
    points: tk.points,
    isCohortLookup,
  };
}

// Build pick objects for the LKP basket from the committed file + snapshot
// timelines, so each COVERED pick can drive the same drill-down chart as our
// STRONG BUY rows (isLkp flags the client-entry framing in openHistoryDrill).
// Out-of-coverage picks carry no points and stay non-clickable.
function buildLkpPicks(lkp, byTicker, todayClose) {
  if (!lkp || !Array.isArray(lkp.picks)) return [];
  return buildLkpPickList(lkp.picks, byTicker, todayClose);
}

// Same shape, takes an explicit picks list. Used by the cohort tracker
// so the click cache can be built from picksByMonth when present (the
// cohort table reads from the same resolved list).
function buildLkpPickList(picks, byTicker, todayClose) {
  if (!Array.isArray(picks)) return [];
  return picks.map((pk) => {
    const t = pk.ticker;
    let name = pk.selection, sector = "", rating = null, points = [];
    if (t && byTicker.has(t)) {
      const tk = byTicker.get(t);
      name = tk.name || pk.selection;
      sector = tk.sector || "";
      points = tk.points || [];
      for (let j = points.length - 1; j >= 0; j--) { if (points[j].rating) { rating = points[j].rating; break; } }
    }
    const close = t ? (todayClose[t] ?? null) : null;
    const covered = !!pk.in_universe && close != null;
    const ret = covered ? (close / pk.entry - 1) * 100 : null;
    return { ...pk, isLkp: true, name, sector, ticker: t, points, close, todayClose: close, currentRating: rating, covered, ret };
  });
}

// LKP Manual picks card — the client's hand-curated basket, tracked with the
// SAME snapshot data as our STRONG BUY history. Covered rows are clickable and
// open the same price+rating drill chart; picks below our market-cap coverage
// render greyed with the client's levels but no live tracking. Returns "" when
// there are no picks.
function renderLkpCard(picks, isOverride) {
  if ((!picks || !picks.length) && !isOverride) return "";
  const body = picks.map((p, i) => {
    const covered = p.covered;
    const { color, initials } = avatarFor(p.name || p.selection || "—");
    const retCls = p.ret == null ? "text-slate-400" : p.ret >= 0 ? "text-emerald-700" : "text-rose-700";
    const retBg  = p.ret == null ? "bg-slate-50 ring-slate-100" : p.ret >= 0 ? "bg-emerald-50 ring-emerald-100" : "bg-rose-50 ring-rose-100";
    const ratingChip = covered && p.currentRating
      ? `<span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap ${composite.ratingClass(p.currentRating)}">${escapeHtml(p.currentRating)}</span>`
      : `<span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap bg-amber-50 text-amber-700 ring-amber-200" title="${escapeHtml(p.out_reason || "Outside our coverage universe")}">NOT COVERED</span>`;
    const pill = (label, val, isSL) => {
      if (val == null) return "";
      if (!covered) return `<span class="px-1.5 py-0.5 rounded bg-slate-50 ring-1 ring-slate-100 text-slate-400 text-[10px] font-semibold tabular-nums">${label} ₹${formatPrice(val)}</span>`;
      const hit = isSL ? p.close <= val : p.close >= val;
      const dist = (val / p.close - 1) * 100;
      const cls = hit
        ? (isSL ? "bg-rose-100 ring-rose-200 text-rose-700" : "bg-emerald-100 ring-emerald-200 text-emerald-700")
        : "bg-slate-50 ring-slate-100 text-slate-600";
      const txt = hit ? "hit" : `${dist >= 0 ? "+" : ""}${dist.toFixed(1)}%`;
      return `<span class="px-1.5 py-0.5 rounded ring-1 ${cls} text-[10px] font-semibold tabular-nums">${label} ₹${formatPrice(val)} · ${txt}</span>`;
    };
    const tag = covered ? "button" : "div";
    const attrs = covered
      ? `class="lkp-row w-full text-left grid grid-cols-12 items-center gap-3 px-4 py-3 rounded-xl bg-white ring-1 ring-slate-200/80 hover:ring-indigo-300 hover:shadow-md hover:-translate-y-px transition-all" data-lkp="${i}"`
      : `class="grid grid-cols-12 items-center gap-3 px-4 py-3 rounded-xl bg-white ring-1 ring-slate-200/80 opacity-70"`;
    return `
      <${tag} ${attrs}>
        <div class="col-span-1 flex items-center">
          <span class="text-xs font-bold text-slate-400 tabular-nums">#${String(i + 1).padStart(2, "0")}</span>
        </div>
        <div class="col-span-7 sm:col-span-4 flex items-center gap-3 min-w-0">
          <div class="w-9 h-9 rounded-lg bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-xs shadow-sm flex-shrink-0">${initials}</div>
          <div class="min-w-0">
            <div class="flex items-center gap-1.5">
              <span class="font-display font-bold text-slate-900 text-sm leading-tight truncate" title="${escapeHtml(p.name)}">${escapeHtml(p.selection)}</span>
              ${ratingChip}
            </div>
            <div class="text-[11px] text-slate-500 truncate">${covered ? escapeHtml(p.sector || p.name) : escapeHtml(p.out_reason || "Below our coverage")}</div>
          </div>
        </div>
        <div class="hidden sm:flex sm:col-span-3 flex-col gap-0.5">
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Entry → Today</div>
          <div class="text-xs text-slate-700">
            <span class="font-semibold tabular-nums">₹${p.entry_low}–${p.entry_high}</span>
            <span class="text-slate-300">→</span>
            <span class="tabular-nums">${covered ? "₹" + formatPrice(p.close) : "—"}</span>
          </div>
        </div>
        <div class="col-span-4 flex flex-col items-end gap-1">
          <div class="flex items-center gap-1.5">
            <span class="inline-flex items-center px-2.5 py-1 rounded-lg ring-1 text-sm font-extrabold tabular-nums ${retCls} ${retBg}">${p.ret == null ? "—" : (p.ret >= 0 ? "+" : "") + p.ret.toFixed(1) + "%"}</span>
            ${covered ? `<span class="text-slate-300 text-base leading-none">›</span>` : ""}
          </div>
          <div class="flex flex-wrap items-center justify-end gap-1">
            ${pill("T1", p.tgt1, false)}
            ${pill("T2", p.tgt2, false)}
            ${pill("SL", p.sl, true)}
          </div>
        </div>
      </${tag}>`;
  }).join("");

  const coveredCount = picks.filter((p) => p.in_universe).length;
  const uploadBtn = `
    <input id="lkp-file-input" type="file" accept=".xlsx,.xls,.csv" class="hidden" />
    <button id="lkp-upload-btn" type="button" class="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-[11px] font-semibold rounded-lg bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm transition">⬆ Upload Excel</button>`;
  const overrideBanner = isOverride ? `
    <div class="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-xl bg-amber-50 ring-1 ring-amber-200 px-3 py-2">
      <div class="text-[11px] text-amber-800 leading-relaxed"><span class="font-bold">Preview only</span> — this uploaded basket is saved on <span class="font-semibold">this device</span>, not yet published to other viewers. Download it and commit as <code class="bg-amber-100 px-1 rounded">public/data/lkp-manual-picks.json</code> (or send it to me) to publish.</div>
      <div class="flex items-center gap-1.5 flex-shrink-0">
        <button id="lkp-download-btn" type="button" class="px-2.5 py-1 text-[11px] font-semibold rounded-lg bg-amber-600 text-white hover:bg-amber-700">Download JSON</button>
        <button id="lkp-reset-btn" type="button" class="px-2.5 py-1 text-[11px] font-semibold rounded-lg bg-white text-slate-600 ring-1 ring-slate-200 hover:ring-slate-300">Reset to published</button>
      </div>
    </div>` : "";
  const emptyState = !picks.length ? `<div class="text-center text-slate-500 text-sm py-8">No picks yet — click <span class="font-semibold">Upload Excel</span> to load the client's basket.<br><span class="text-[11px] text-slate-400">Expected columns: Selection, Entry (range ok), TGT1, TGT2, SL.</span></div>` : "";
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 sm:p-5 mt-4">
      <div class="flex flex-wrap items-center justify-between gap-2 mb-1">
        <h2 class="font-display font-bold text-slate-900 text-base">LKP Manual picks</h2>
        <div class="flex items-center gap-3">
          ${picks.length ? `<span class="text-[11px] text-slate-500 hidden sm:inline">${coveredCount}/${picks.length} in coverage · click a covered row for its chart</span>` : ""}
          ${uploadBtn}
        </div>
      </div>
      ${overrideBanner}
      <div class="space-y-1.5">${body}</div>
      ${emptyState}
    </div>`;
}

// Targets / stop-loss block shown inside the LKP drill modal, in place of the
// STRONG BUY pillar forensics. Distances are vs today's close; "hit" once the
// level is crossed.
function renderLkpTargets(pick) {
  const close = pick.close;
  const card = (label, val, isSL) => {
    if (val == null) return "";
    const hit = isSL ? close <= val : close >= val;
    const dist = (val / close - 1) * 100;
    const distCls = hit ? (isSL ? "text-rose-700" : "text-emerald-700") : "text-slate-600";
    const ringCls = hit ? (isSL ? "ring-rose-200 bg-rose-50" : "ring-emerald-200 bg-emerald-50") : "ring-slate-200 bg-white";
    return `<div class="rounded-xl ring-1 ${ringCls} px-3 py-2 text-center">
      <div class="text-[9px] font-bold uppercase tracking-wider text-slate-400">${label}</div>
      <div class="text-sm font-display font-bold text-slate-900 mt-0.5 tabular-nums">₹${formatPrice(val)}</div>
      <div class="text-[11px] font-semibold tabular-nums mt-0.5 ${distCls}">${hit ? "✓ hit" : (dist >= 0 ? "+" : "") + dist.toFixed(1) + "%"}</div>
    </div>`;
  };
  return `<div class="grid grid-cols-3 gap-2.5 mb-3">
    ${card("Target 1", pick.tgt1, false)}
    ${card("Target 2", pick.tgt2, false)}
    ${card("Stop-loss", pick.sl, true)}
  </div>`;
}

// ---------------- LKP Excel upload ----------------
// Client self-service: parse an uploaded Excel/CSV in the browser, match each
// row to our NSE symbol, and preview the basket instantly (localStorage). The
// committed public/data/lkp-manual-picks.json stays the shared source of
// truth — "Download JSON" exports the parsed basket so it can be committed.
const LKP_OVERRIDE_KEY = "lkp_manual_override_v1";
function lkpOverride() { try { const s = localStorage.getItem(LKP_OVERRIDE_KEY); return s ? JSON.parse(s) : null; } catch { return null; } }
function saveLkpOverride(d) { try { localStorage.setItem(LKP_OVERRIDE_KEY, JSON.stringify(d)); } catch {} }
function clearLkpOverride() { try { localStorage.removeItem(LKP_OVERRIDE_KEY); } catch {} }

const lkpNk = (s) => String(s == null ? "" : s).toUpperCase().replace(/[^A-Z0-9]/g, "");
function lkpNum(v) { const n = Number(String(v == null ? "" : v).replace(/[,₹\s]/g, "")); return Number.isFinite(n) ? n : null; }
function lkpParseRange(s) {
  const str = String(s || "");
  const m = str.match(/(\d[\d,]*\.?\d*)\s*(?:[-–—]|to)\s*(\d[\d,]*\.?\d*)/i);
  if (m) return [lkpNum(m[1]), lkpNum(m[2])];
  const one = lkpNum(str);
  return one != null ? [one, one] : [null, null];
}

// Lazy-load SheetJS from CDN, only when an .xlsx is actually uploaded. CSV is
// parsed natively (no dependency), so it always works even offline.
let _xlsxPromise = null;
function loadSheetJS() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (_xlsxPromise) return _xlsxPromise;
  _xlsxPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    s.onload = () => resolve(window.XLSX);
    s.onerror = () => reject(new Error("Couldn't load the Excel parser — check your connection, or save the sheet as .csv and upload that."));
    document.head.appendChild(s);
  });
  return _xlsxPromise;
}
function lkpParseCsv(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) return [];
  const cells = (l) => (l.match(/("(?:[^"]|"")*"|[^,]*)(?=,|$)/g) || []).map((c) => c.replace(/^"|"$/g, "").replace(/""/g, '"').trim());
  const header = cells(lines[0]);
  return lines.slice(1).map((l) => { const c = cells(l); const o = {}; header.forEach((h, i) => (o[h] = c[i] ?? "")); return o; });
}
async function lkpParseSheet(file) {
  if ((file.name || "").toLowerCase().endsWith(".csv")) return lkpParseCsv(await file.text());
  const XLSX = await loadSheetJS();
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
}

// Map an uploaded label to our NSE symbol using the latest snapshot universe.
// Conservative: exact ticker, curated aliases, exact/unique normalized name —
// no loose fuzzy that could mis-match.
function buildLkpUniverse() {
  const snaps = state.cache?.history?.snapshots || [];
  const last = snaps[snaps.length - 1];
  const m = new Map();
  if (last) for (const s of last.stocks) if (s.ticker) m.set(s.ticker.toUpperCase(), lkpNk(s.name || ""));
  return m;
}
const LKP_ALIAS = { EXIDE: "EXIDEIND", FEDERALBANK: "FEDERALBNK", ATHERENERGY: "ATHERENERG", ATHER: "ATHERENERG", LANDT: "LT", BAJAJFINANCE: "BAJFINANCE" };
function lkpMatchTicker(label, universe) {
  const L = lkpNk(label);
  if (!L) return null;
  if (universe.has(L)) return L;                                    // exact ticker
  if (LKP_ALIAS[L] && universe.has(LKP_ALIAS[L])) return LKP_ALIAS[L]; // curated alias
  for (const [t, nName] of universe) if (nName && nName === L) return t; // exact normalized name
  if (L.length >= 4) {
    const tp = [...universe.keys()].filter((t) => t.startsWith(L));
    if (tp.length === 1) return tp[0];                              // unique ticker prefix
    const np = [...universe].filter(([, nName]) => nName && nName.startsWith(L));
    if (np.length === 1) return np[0][0];                           // unique name prefix
  }
  return null;
}
function lkpRowToPick(row, universe) {
  const get = (aliases) => { for (const [k, v] of Object.entries(row)) if (aliases.includes(lkpNk(k))) return v; return ""; };
  const sel = String(get(["SELECTION", "STOCK", "TICKER", "SYMBOL", "SCRIP", "NAME", "COMPANY"]) || "").trim();
  if (!sel) return null;
  const [low, high] = lkpParseRange(get(["ENTRY", "ENTRYPRICE", "BUYRANGE", "BUY", "ENTRYRANGE", "RANGE"]));
  const entry = (low != null && high != null) ? (low + high) / 2 : (low ?? high);
  if (entry == null) return null;
  const ticker = lkpMatchTicker(sel, universe);
  return {
    selection: sel.toUpperCase(),
    ticker,
    in_universe: !!ticker,
    out_reason: ticker ? undefined : "Not matched to our coverage — check the NSE symbol, or it's below our market-cap floor",
    entry_low: low ?? entry, entry_high: high ?? entry, entry,
    tgt1: lkpNum(get(["TGT1", "TARGET1", "T1", "TARGET", "TGT"])),
    tgt2: lkpNum(get(["TGT2", "TARGET2", "T2"])),
    sl: lkpNum(get(["SL", "STOPLOSS", "STOP", "STOPLOSSPRICE"])),
  };
}
async function handleLkpExcelUpload(file) {
  try {
    const rows = await lkpParseSheet(file);
    const universe = buildLkpUniverse();
    const picks = rows.map((r) => lkpRowToPick(r, universe)).filter(Boolean);
    if (!picks.length) { alert("No valid rows found. Expected columns like: Selection, Entry, TGT1, TGT2, SL."); return; }
    saveLkpOverride({ label: "LKP Manual picks", source: "Excel upload (browser preview)", generated_at: new Date().toISOString(), picks });
    // Re-render whichever tab is showing the manual basket.
    const t = state.activeTab;
    if (t === "custom") renderCustom(); else if (t === "active") renderActive(); else renderHistory();
    const matched = picks.filter((p) => p.in_universe).length;
    if (matched < picks.length) {
      const miss = picks.filter((p) => !p.in_universe).map((p) => p.selection).join(", ");
      alert(`Loaded ${picks.length} picks. ${matched} matched our coverage; ${picks.length - matched} not matched (${miss}). Those show greyed — fix the NSE symbol in the sheet and re-upload if needed.`);
    }
  } catch (e) {
    alert("Couldn't read that file: " + (e?.message || e));
  }
}
function downloadLkpJson() {
  const data = lkpOverride();
  if (!data) return;
  const blob = new Blob([JSON.stringify(data, null, 2) + "\n"], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "lkp-manual-picks.json";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Tiny portfolio-value sparkline drawn inside the backtest card. Shows
// the running portfolio return (%) day-over-day, with each pick-entry
// date marked as a small green dot on the line.
function renderPortfolioSparkline(series, H) {
  if (!series || series.length < 2) return `<div class="text-[11px] text-slate-400 text-center py-4">Not enough days to draw a curve yet.</div>`;
  const W = 800;
  const M = { top: 8, right: 8, bottom: 22, left: 36 };
  const innerW = W - M.left - M.right;
  const innerH = H - M.top - M.bottom;
  const rets = series.map((p) => p.ret);
  const yMin = Math.min(0, ...rets);
  const yMax = Math.max(0, ...rets);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.1, yHi = yMax + ySpan * 0.1;
  const xAt = (i) => M.left + (i / (series.length - 1)) * innerW;
  const yAt = (v) => M.top + innerH - ((v - yLo) / (yHi - yLo)) * innerH;
  const path = series.map((p, i) => `${i === 0 ? "M" : "L"} ${xAt(i).toFixed(1)} ${yAt(p.ret).toFixed(1)}`).join(" ");
  const lastRet = series[series.length - 1].ret;
  const lineColor = lastRet >= 0 ? "#10b981" : "#f43f5e";
  const areaPath = `${path} L ${xAt(series.length - 1).toFixed(1)} ${yAt(yLo).toFixed(1)} L ${M.left.toFixed(1)} ${yAt(yLo).toFixed(1)} Z`;
  const zeroLineY = yAt(0).toFixed(1);
  const dateTickEvery = Math.max(1, Math.ceil(series.length / 6));
  const xLabels = series.map((p, i) => {
    if (i % dateTickEvery !== 0 && i !== series.length - 1) return "";
    return `<text x="${xAt(i).toFixed(1)}" y="${(M.top + innerH + 14).toFixed(1)}" text-anchor="middle" font-size="9" fill="#94a3b8">${p.date.slice(5)}</text>`;
  }).join("");
  return `
    <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" class="w-full" style="max-height:${H + 20}px">
      <defs>
        <linearGradient id="portArea" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="${lineColor}" stop-opacity="0.22"/>
          <stop offset="100%" stop-color="${lineColor}" stop-opacity="0"/>
        </linearGradient>
      </defs>
      <line x1="${M.left}" x2="${(M.left + innerW).toFixed(1)}" y1="${zeroLineY}" y2="${zeroLineY}" stroke="#cbd5e1" stroke-width="0.6" stroke-dasharray="3 4" />
      <text x="${(M.left - 6).toFixed(1)}" y="${zeroLineY}" text-anchor="end" dominant-baseline="middle" font-size="9" fill="#94a3b8">0%</text>
      <text x="${(M.left - 6).toFixed(1)}" y="${(M.top + 6).toFixed(1)}" text-anchor="end" dominant-baseline="middle" font-size="9" fill="#94a3b8">${yHi >= 0 ? "+" : ""}${yHi.toFixed(1)}%</text>
      <text x="${(M.left - 6).toFixed(1)}" y="${(M.top + innerH - 2).toFixed(1)}" text-anchor="end" dominant-baseline="middle" font-size="9" fill="#94a3b8">${yLo >= 0 ? "+" : ""}${yLo.toFixed(1)}%</text>
      <path d="${areaPath}" fill="url(#portArea)" />
      <path d="${path}" fill="none" stroke="${lineColor}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" />
      <circle cx="${xAt(series.length - 1).toFixed(1)}" cy="${yAt(lastRet).toFixed(1)}" r="3.5" fill="${lineColor}" stroke="#fff" stroke-width="1.5" />
      ${xLabels}
    </svg>
  `;
}

function portfolioStat(label, value, sub, valueCls = "text-slate-900") {
  return `
    <div class="rounded-xl bg-slate-50 ring-1 ring-slate-100 px-3 py-2">
      <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400">${escapeHtml(label)}</div>
      <div class="text-base font-display font-extrabold tabular-nums mt-0.5 leading-tight ${valueCls}">${value}</div>
      <div class="text-[10px] text-slate-500 mt-0.5 truncate" title="${escapeHtml(sub)}">${escapeHtml(sub)}</div>
    </div>
  `;
}

// Indian-rupee compact formatter — ₹X,XX,XXX up to a lakh, then
// ₹X.XL / ₹X.XCr for bigger numbers (matches Indian convention).
function formatINR(n) {
  if (n == null || !isFinite(n)) return "—";
  const abs = Math.abs(n);
  const sign = n < 0 ? "−" : "";
  if (abs >= 10000000) return `${sign}${(abs / 10000000).toFixed(2)}Cr`;
  if (abs >= 100000)   return `${sign}${(abs / 100000).toFixed(2)}L`;
  return `${sign}${Math.round(abs).toLocaleString("en-IN")}`;
}

// Score forensics — decompose the composite change between the
// STRONG BUY anchor day and today into per-pillar weighted-delta
// contributions. The sum of all five deltas equals the composite
// delta, so a fund manager can see *why* the model changed its mind
// (or didn't). Quiet if either endpoint lacks pillar data.
function renderScoreForensics(pick) {
  const a = pick.firstSBPillars;
  const b = pick.todayPillars;
  if (!a || !b) return "";
  const PILLAR_KEYS = ["fundamentals", "technicals", "macro", "sentiment", "liquidity"];
  const LABEL = { fundamentals: "Fundamentals", technicals: "Technicals", macro: "Macro", sentiment: "Sentiment", liquidity: "Liquidity" };
  const COLOR = { fundamentals: "bg-emerald-500", technicals: "bg-indigo-500", macro: "bg-violet-500", sentiment: "bg-amber-500", liquidity: "bg-sky-500" };
  const deltas = PILLAR_KEYS.map((k) => {
    const aw = a[k]?.weighted, bw = b[k]?.weighted;
    const ap = a[k]?.pct,      bp = b[k]?.pct;
    if (aw == null || bw == null) return { key: k, missing: true };
    return { key: k, missing: false, aw, bw, ap, bp, delta: bw - aw };
  });
  const compDelta = deltas.reduce((s, d) => s + (d.missing ? 0 : d.delta), 0);
  const compTotalA = pick.firstSBComposite;
  const compTotalB = (a && b && pick.todayPillars)
    ? deltas.reduce((s, d) => s + (d.missing ? 0 : d.bw), 0)
    : null;
  const maxAbs = Math.max(0.5, ...deltas.map((d) => d.missing ? 0 : Math.abs(d.delta)));
  const deltaSign = compDelta >= 0 ? "+" : "−";
  const deltaCls = compDelta >= 0 ? "text-emerald-700" : "text-rose-700";

  // How often each pillar's inputs actually move, so a reader knows a flat
  // Fundamentals / Macro is expected and the day-to-day swing is the fast ones.
  const CADENCE = { fundamentals: "quarterly", technicals: "daily", macro: "monthly", sentiment: "daily", liquidity: "daily" };
  const CADENCE_CLS = { daily: "bg-sky-50 text-sky-700 ring-sky-200", monthly: "bg-violet-50 text-violet-700 ring-violet-200", quarterly: "bg-slate-100 text-slate-500 ring-slate-200" };
  const cadenceChip = (k) => `<span class="text-[8px] font-bold uppercase tracking-wide px-1 py-px rounded ring-1 ${CADENCE_CLS[CADENCE[k]]} whitespace-nowrap">${CADENCE[k]}</span>`;

  // Movers first (largest absolute contribution) so the eye lands on what
  // actually changed instead of a wall of "+0.00".
  const ordered = [...deltas].sort((x, y) => (x.missing ? -1 : Math.abs(x.delta)) < (y.missing ? -1 : Math.abs(y.delta)) ? 1 : -1);

  const rowsHtml = ordered.map((d) => {
    if (d.missing) {
      return `
        <div class="grid grid-cols-12 items-center gap-2 py-1">
          <div class="col-span-6 flex items-center gap-1.5"><span class="w-1.5 h-1.5 rounded-full bg-slate-300 flex-shrink-0"></span><span class="text-[11px] font-semibold text-slate-400">${LABEL[d.key]}</span>${cadenceChip(d.key)}</div>
          <div class="col-span-6 text-[10px] text-slate-400 text-right">no data</div>
        </div>`;
    }
    const pos = d.delta >= 0;
    const changed = Math.abs(d.delta) >= 0.01;
    const widthPct = (Math.abs(d.delta) / maxAbs) * 50;          // 0..50% of the bar (centered)
    const valCls = !changed ? "text-slate-400" : pos ? "text-emerald-700" : "text-rose-700";
    const barCls = !changed ? "bg-slate-300" : pos ? "bg-emerald-500" : "bg-rose-500";
    return `
      <div class="grid grid-cols-12 items-center gap-2 py-1">
        <div class="col-span-5 flex items-center gap-1.5 min-w-0">
          <span class="w-1.5 h-1.5 rounded-full ${COLOR[d.key]} flex-shrink-0"></span>
          <span class="text-[11px] font-semibold text-slate-800 truncate">${LABEL[d.key]}</span>
          ${cadenceChip(d.key)}
        </div>
        <div class="col-span-3 text-[10px] tabular-nums text-right whitespace-nowrap"><span class="text-slate-400">${d.ap}%</span> <span class="text-slate-300">→</span> <span class="font-semibold text-slate-700">${d.bp}%</span></div>
        <div class="col-span-4 flex items-center justify-end gap-2">
          <div class="relative h-1.5 w-12 sm:w-16 bg-slate-100 rounded-full overflow-hidden flex-shrink-0">
            <div class="absolute top-0 bottom-0 left-1/2 w-px bg-slate-300"></div>
            <div class="absolute top-0 bottom-0 ${barCls}" style="${pos ? `left:50%;width:${widthPct.toFixed(2)}%` : `right:50%;width:${widthPct.toFixed(2)}%`}"></div>
          </div>
          <span class="text-[11px] font-bold tabular-nums ${valCls} w-11 text-right">${changed ? (pos ? "+" : "−") + Math.abs(d.delta).toFixed(2) : "0.00"}</span>
        </div>
      </div>`;
  }).join("");

  return `
    <details class="rounded-2xl ring-1 ring-slate-200 bg-white mb-2 group">
      <summary class="cursor-pointer select-none flex flex-wrap items-baseline justify-between gap-2 px-3 py-2 sm:px-4 sm:py-2.5 hover:bg-slate-50 rounded-2xl">
        <div class="flex items-center gap-2">
          <span class="text-slate-400 group-open:rotate-90 transition-transform text-xs">▸</span>
          <span class="font-display font-bold text-slate-900 text-sm">What moved the score</span>
        </div>
        <div class="flex items-baseline gap-1.5 text-xs tabular-nums">
          <span class="text-slate-500">${compTotalA != null ? compTotalA.toFixed(1) : "—"}</span>
          <span class="text-slate-400">→</span>
          <span class="font-bold text-slate-900">${compTotalB != null ? compTotalB.toFixed(1) : "—"}</span>
          <span class="font-bold ${deltaCls}">${deltaSign}${Math.abs(compDelta).toFixed(2)}</span>
          <span class="text-[10px] text-slate-400 font-normal">since ${fmtDateDMY(pick.firstSBDate)}</span>
        </div>
      </summary>
      <div class="px-3 pb-3 sm:px-4 sm:pb-3 pt-0">
        <div class="grid grid-cols-12 gap-2 px-0.5 pb-1 mb-0.5 border-b border-slate-100 text-[9px] font-bold uppercase tracking-wider text-slate-400">
          <div class="col-span-5">Pillar · refresh</div>
          <div class="col-span-3 text-right">Score</div>
          <div class="col-span-4 text-right">Contribution</div>
        </div>
        <div class="divide-y divide-slate-50">${rowsHtml}</div>
        <p class="text-[10px] text-slate-400 leading-relaxed mt-2 pt-2 border-t border-slate-100"><span class="font-semibold text-slate-500">Contribution</span> = how much each pillar pushed the composite (the five sum to the total change). The chip is how often it refreshes — <span class="font-semibold">Fundamentals</span> (quarterly) and <span class="font-semibold">Macro</span> (monthly) barely move day-to-day, so short-term swings come from <span class="font-semibold">Technicals / Sentiment / Liquidity</span> (daily).</p>
      </div>
    </details>`;
}

function heroStat(label, value, sub) {
  return `
    <div class="bg-white/10 backdrop-blur rounded-xl p-3 ring-1 ring-white/15">
      <div class="text-[10px] uppercase tracking-wider text-white/60 font-semibold">${escapeHtml(label)}</div>
      <div class="text-xl font-bold mt-1 tabular-nums">${escapeHtml(String(value))}</div>
      <div class="text-[10px] text-white/60 mt-0.5 truncate" title="${escapeHtml(sub)}">${escapeHtml(sub)}</div>
    </div>
  `;
}

// ---------------- Monthly Cohort Tracker ----------------
// At each month-end, lock the AI basket's top 7 (highest composite,
// regardless of band) as the cohort that's HELD through the next
// month. Compute daily average return for that locked group, alongside
// the client's LKP picks (same entry-month logic, midpoint of entry
// range as cost basis) and Nifty over the same window. Three lines
// on one chart so a fund manager can see the basket vs the client's
// pick vs the index for each calendar month.

const COHORT_RATING_BG = {
  "STRONG BUY": "bg-emerald-100 text-emerald-700 ring-emerald-200",
  "BUY":        "bg-blue-100 text-blue-700 ring-blue-200",
  "WATCH":      "bg-amber-100 text-amber-700 ring-amber-200",
  "AVOID":      "bg-rose-100 text-rose-700 ring-rose-200",
  "FILTERED":   "bg-rose-50 text-rose-700 ring-rose-200",
};

// IST calendar today (YYYY-MM-DD). Used by Static mode + the manual
// basket month lookup so both stay aligned through a month rollover
// (e.g. July 1 morning before the first July snapshot is written).
function istTodayDate() {
  const istMs = Date.now() + 5.5 * 3600 * 1000;
  return new Date(istMs).toISOString().slice(0, 10);
}

// Last snapshot of the calendar month immediately before today's date.
// Used by the Static view ("month-end") so the AI top 7 is locked from
// e.g. 31-05-26 for the entire month of June. If only one month of
// snapshots exists, falls back to the earliest available snapshot so
// the tracker still renders something.
//
// "Today" is computed in IST (markets we track), not from the latest
// snapshot date — so on the morning of a month rollover (e.g. July 1
// before the first July snapshot is written) Static mode anchors at
// June 30 instead of staying stale at May 31.
function staticAnchorDate(snapshots) {
  if (!snapshots.length) return null;
  const todayMonth = istTodayDate().slice(0, 7);
  for (let i = snapshots.length - 1; i >= 0; i--) {
    if (snapshots[i].date.slice(0, 7) !== todayMonth) return snapshots[i].date;
  }
  return snapshots[0].date;
}

// Derive the anchor date — the day the client uploaded their basket.
// Founder's rule: "the clock starts the day client gave us the basket".
// Falls back to most-recent month-end snapshot if no LKP file exists
// yet, so the tracker still shows something (AI line + Nifty line) for
// internal review.
// Months the client has baskets for (oldest → newest). Uses picksByMonth
// keys when present; otherwise the single legacy basket's anchor month.
function availableManualMonths(lkp, snapshots) {
  if (!lkp) return [];
  if (lkp.picksByMonth && typeof lkp.picksByMonth === "object") {
    return Object.keys(lkp.picksByMonth)
      .filter((m) => Array.isArray(lkp.picksByMonth[m]) && lkp.picksByMonth[m].length)
      .sort();
  }
  const d = lkpAnchorDate(lkp, snapshots);
  return d ? [d.slice(0, 7)] : [];
}

// Anchor (cost-basis) date for a given basket month. Prefer the recorded
// anchorByMonth; else the first snapshot in that month; else the file's
// generated_at.
function manualMonthAnchor(lkp, month, snapshots) {
  if (!month) return lkpAnchorDate(lkp, snapshots);
  if (lkp?.anchorByMonth?.[month]) return lkp.anchorByMonth[month];
  const s = snapshots.find((x) => x.date.slice(0, 7) === month);
  if (s) return s.date;
  return lkpAnchorDate(lkp, snapshots);
}

function monthLabelYM(ym) {
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const [y, m] = ym.split("-").map(Number);
  return `${MONTHS[m - 1] || ym} ${y}`;
}

// Add n calendar months to a YYYY-MM-DD date (UTC).
function addMonthsStr(dateStr, n) {
  const [y, m, d] = dateStr.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1 + n, d)).toISOString().slice(0, 10);
}

// Report-month selector — one pill per client basket month so June stays
// visible after July is uploaded (each keeps its own anchor + track
// record). The latest month is the default; hidden when only one exists.
function renderManualMonthSelector(months, selectedMonth, lkp, snapshots) {
  if (!months || months.length <= 1) return "";
  const pills = months.map((m) => {
    const on = m === selectedMonth;
    const isLatest = m === months[months.length - 1];
    return `
      <button type="button" data-manual-month="${m}" class="px-3 py-1.5 rounded-lg text-sm font-semibold whitespace-nowrap transition ${on ? "bg-slate-900 text-white shadow-sm" : "bg-slate-50 text-slate-600 ring-1 ring-slate-200 hover:ring-slate-400 hover:text-slate-900"}">
        ${escapeHtml(monthLabelYM(m))}${isLatest ? ` <span class="text-[9px] font-bold uppercase ${on ? "text-slate-300" : "text-emerald-600"}">latest</span>` : ""}
      </button>`;
  }).join("");
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-2 sm:p-3 flex flex-wrap items-center gap-2">
      <div class="flex items-center gap-1.5 px-2">
        <span class="text-base">📅</span>
        <span class="text-[11px] font-bold uppercase tracking-wider text-slate-500">Report month</span>
      </div>
      <div class="flex items-center gap-1 overflow-x-auto -my-1 py-1">${pills}</div>
    </div>
  `;
}

// Cohort anchor: the first snapshot of the CURRENT month. The basket is
// picked on the first trading day of each month and held for the month, so
// that date is both the selection day and the cost basis.
//
// Previously this keyed off lkp.generated_at -- the day the client handed
// over their basket. With no manual basket that date does not exist, and the
// old fallback walked BACK to the previous month's last snapshot, which
// would have anchored an August cohort to a July date.
//
// The lkp argument is retained so existing call sites keep working.
function lkpAnchorDate(lkp, snapshots) {
  if (!snapshots.length) return null;
  const ym = snapshots[snapshots.length - 1].date.slice(0, 7);
  const firstOfMonth = snapshots.find((s) => s.date.slice(0, 7) === ym);
  // Snapshots are ordered oldest -> newest, so find() gives the FIRST one in
  // the current month. Every snapshot belongs to some month, so this always
  // resolves once the trail is non-empty; the fallback is belt-and-braces.
  return firstOfMonth ? firstOfMonth.date : snapshots[0].date;
}

// Build the cohort. Returns a single object with `segments` — one entry
// for monthly mode (the whole window since anchor), N entries for
// weekly mode (one per 7-day window). Each segment owns its own AI
// top-7; in weekly mode the basket re-locks at each segment boundary.
//
// `opts.selectionDate` decouples the basket SELECTION day from the
// COST BASIS day. In Static mode the AI top 7 is picked from the
// prior month-end snapshot (selectionDate) but the performance
// window starts at the first held-month snapshot (anchorDate) — per
// founder's call, "assume client bought June 1" even though the
// screener output is from May 31's close.
function buildCohort(snapshots, anchorDate, mode, opts = {}) {
  if (!snapshots.length || !anchorDate) return null;
  const today = snapshots[snapshots.length - 1].date;
  if (anchorDate > today) anchorDate = today;

  const anchorSnap = snapshots.find((s) => s.date >= anchorDate);
  if (!anchorSnap) return null;
  const selectionSnap = opts.selectionDate
    ? (snapshots.find((s) => s.date >= opts.selectionDate) || anchorSnap)
    : anchorSnap;

  const pickTop7 = (snap) => snap.stocks
    .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed)
    .sort((a, b) => b.composite - a.composite)
    .slice(0, 7);

  // Pick the top 7 from the selection day, but re-base each stock's
  // close to the cost-basis day (anchorSnap) so chart day-0 = entry
  // day and returns measure from there. Selection-day rating is
  // preserved as the display tag.
  const remapTop7 = (sel, entry) => pickTop7(sel).map((s) => {
    const eS = entry.stocks.find((x) => x.ticker === s.ticker);
    return { ...s, close: eS?.close ?? s.close, _selectionDate: sel.date };
  });

  if (mode === "monthly") {
    const tracking = snapshots.filter((s) => s.date >= anchorSnap.date);
    return {
      mode: "monthly",
      anchorDate,
      effectiveStart: anchorSnap.date,
      selectionDate: selectionSnap.date,
      segments: [{
        index: 0,
        label: "Since upload",
        startDate: anchorSnap.date,
        endDate: tracking[tracking.length - 1].date,
        entrySnap: anchorSnap,
        tracking,
        top7: remapTop7(selectionSnap, anchorSnap),
      }],
    };
  }

  // Weekly: rolling 7-day windows starting at the anchor. Weekly mode
  // ignores opts.selectionDate — each week's basket is picked from
  // that week's own entry snapshot, no cross-window decoupling.
  const fmt = (ms) => new Date(ms).toISOString().slice(0, 10);
  const dateToMs = (d) => Date.UTC(
    Number(d.slice(0, 4)), Number(d.slice(5, 7)) - 1, Number(d.slice(8, 10)),
  );
  let cursorMs = dateToMs(anchorDate);
  const todayMs = dateToMs(today);
  const segments = [];
  while (cursorMs <= todayMs) {
    const winStart = fmt(cursorMs);
    const winEnd = fmt(cursorMs + 6 * 86400000);
    const entry = snapshots.find((s) => s.date >= winStart && s.date <= winEnd);
    if (!entry) { cursorMs += 7 * 86400000; continue; }
    const tracking = snapshots.filter((s) => s.date >= entry.date && s.date <= winEnd);
    segments.push({
      index: segments.length,
      label: `Week ${segments.length + 1}`,
      startDate: entry.date,
      endDate: tracking[tracking.length - 1].date,
      windowStart: winStart,
      windowEnd: winEnd,
      entrySnap: entry,
      tracking,
      top7: pickTop7(entry),
    });
    cursorMs += 7 * 86400000;
  }
  if (segments.length === 0) return null;
  return {
    mode: "weekly",
    anchorDate,
    effectiveStart: segments[0].startDate,
    selectionDate: selectionSnap.date,
    segments,
  };
}

// Per-month client basket lookup.
// Preferred shape: lkp.picksByMonth["YYYY-MM"] -> array of pick objects.
// Legacy fallback: lkp.picks (single basket). The fallback only applies
// to the MOST RECENT cohort month, so uploading a new latest basket
// doesn't silently rewrite client baskets in older months' charts.
function lkpPicksForMonth(lkp, cohortMonth, mostRecentMonth) {
  if (!lkp) return [];
  if (lkp.picksByMonth && Array.isArray(lkp.picksByMonth[cohortMonth])) {
    return lkp.picksByMonth[cohortMonth];
  }
  // The manual basket is "sticky" — the last uploaded picks keep showing
  // (anchored at their upload date) until the client uploads a fresh set,
  // even after the calendar rolls into a new month.
  if (Array.isArray(lkp.picks)) return lkp.picks;
  return [];
}

// Compute the time series for the chart + per-segment tables.
//
// Three series get plotted: AI (cumulative if cohort.mode==="weekly",
// holding-period if monthly), Manual (always anchored at the upload
// snapshot's close, static basket), Nifty 50.
//
// For weekly mode the AI line "re-locks" each Monday and the
// cumulative factor = product over completed segments of their final
// factor × current segment's partial factor. The line stays
// mathematically continuous through each boundary (each segment's
// own factor starts at 1.0, multiplied by the running cumulative);
// boundary lines are added on the chart as honest visual markers.
//
// Forward-fill behaviour preserved from the prior version: a missing
// daily close uses the last known price so the basket denominator
// stays at its locked size.
function buildCohortSeries(cohort, manualPicksInput, niftyOn) {
  if (!cohort) return null;
  const anchorSnap = cohort.segments[0]?.entrySnap;
  if (!anchorSnap) return null;
  const anchorDate = anchorSnap.date;

  // Manual basket: in-universe picks only count for the average. Cost
  // basis = anchor snapshot close (NOT the LKP file's entry range).
  const manualAll = manualPicksInput || [];
  const manualInUni = manualAll.filter((p) => p && p.in_universe && p.ticker);
  const manualEntry = {};
  const manualLastRating = {};
  for (const p of manualInUni) {
    const s = anchorSnap.stocks.find((x) => x.ticker === p.ticker);
    if (s && s.close != null) {
      manualEntry[p.ticker] = s.close;
      manualLastRating[p.ticker] = s.rating || null;
    }
  }
  const manualLastClose = { ...manualEntry };

  const niftyAtEntry = niftyOn(anchorDate);

  // Pre-compute each segment's per-day per-stock detail (with forward
  // fill) AND the segment's end-factor for the cumulative AI math.
  for (const seg of cohort.segments) {
    seg.entryCloses = {};
    seg.entryRatings = {};
    for (const s of seg.top7) {
      seg.entryCloses[s.ticker] = s.close;
      seg.entryRatings[s.ticker] = s.rating || null;
    }
    seg.lastClose = { ...seg.entryCloses };
    seg.lastRating = { ...seg.entryRatings };
    seg.dailyPerStock = {};

    for (const day of seg.tracking) {
      const perStock = {};
      for (const ticker of Object.keys(seg.entryCloses)) {
        const s = day.stocks.find((x) => x.ticker === ticker);
        let close = s?.close;
        let stale = false;
        if (close == null) { close = seg.lastClose[ticker]; stale = true; }
        else { seg.lastClose[ticker] = close; }
        if (s?.rating != null) seg.lastRating[ticker] = s.rating;
        if (close == null) {
          perStock[ticker] = { close: null, ret: null, rating: seg.lastRating[ticker], stale: true };
          continue;
        }
        const ret = (close / seg.entryCloses[ticker] - 1) * 100;
        perStock[ticker] = {
          close, ret, rating: seg.lastRating[ticker],
          hardFailed: !!s?.hardFailed, stale,
        };
      }
      seg.dailyPerStock[day.date] = perStock;
    }
    // End-factor = avg(today's close / entry close) across the basket.
    const endDate = seg.tracking[seg.tracking.length - 1]?.date;
    seg.endFactor = endDate ? avgFactorFromPerStock(seg.dailyPerStock[endDate], seg.entryCloses) : null;
  }

  // Top-level cumulative series (one per snapshot date in the cohort)
  const points = [];
  // Day 0 = anchor
  points.push({
    date: anchorDate,
    aiCum: 0,
    manualCum: Object.keys(manualEntry).length ? 0 : null,
    niftyCum: niftyAtEntry != null ? 0 : null,
    segmentIdx: 0,
    isBoundary: false,
  });

  let prevSegFactor = 1.0;
  for (let segIdx = 0; segIdx < cohort.segments.length; segIdx++) {
    const seg = cohort.segments[segIdx];
    for (const day of seg.tracking) {
      if (day.date === anchorDate) continue;
      const perStock = seg.dailyPerStock[day.date];
      const curFactor = avgFactorFromPerStock(perStock, seg.entryCloses);
      const aiCum = curFactor != null ? (prevSegFactor * curFactor - 1) * 100 : null;

      // Manual: always anchored at anchorDate
      let manualSum = 0, manualN = 0;
      for (const ticker of Object.keys(manualEntry)) {
        const s = day.stocks.find((x) => x.ticker === ticker);
        let close = s?.close;
        if (close == null) close = manualLastClose[ticker];
        else manualLastClose[ticker] = close;
        if (s?.rating != null) manualLastRating[ticker] = s.rating;
        if (close == null) continue;
        manualSum += close / manualEntry[ticker];
        manualN++;
      }
      const manualCum = manualN > 0 ? (manualSum / manualN - 1) * 100 : null;

      const niftyClose = niftyOn(day.date);
      const niftyCum = (niftyAtEntry != null && niftyClose != null)
        ? (niftyClose / niftyAtEntry - 1) * 100
        : null;

      const isBoundary = (segIdx > 0 && day.date === seg.startDate);
      points.push({ date: day.date, aiCum, manualCum, niftyCum, segmentIdx: segIdx, isBoundary });
    }
    // Lock in this segment's end factor for the next segment's cumulative.
    if (seg.endFactor != null && segIdx < cohort.segments.length - 1) {
      prevSegFactor *= seg.endFactor;
    }
  }

  // Manual per-day per-stock (used by the Manual table). Same forward-
  // fill semantics as AI segments.
  const manualLastCloseTable = { ...manualEntry };
  const manualDailyPerStock = {};
  manualDailyPerStock[anchorDate] = Object.fromEntries(
    Object.entries(manualEntry).map(([t, c]) => [t, { close: c, ret: 0, rating: manualLastRating[t] }])
  );
  // Walk all snapshot dates from the anchor onwards
  const allTrackingDates = [];
  const seen = new Set();
  for (const seg of cohort.segments) for (const d of seg.tracking) {
    if (!seen.has(d.date)) { allTrackingDates.push(d); seen.add(d.date); }
  }
  for (const day of allTrackingDates) {
    if (day.date === anchorDate) continue;
    const perStock = {};
    for (const ticker of Object.keys(manualEntry)) {
      const s = day.stocks.find((x) => x.ticker === ticker);
      let close = s?.close;
      let stale = false;
      if (close == null) { close = manualLastCloseTable[ticker]; stale = true; }
      else { manualLastCloseTable[ticker] = close; }
      if (s?.rating != null) manualLastRating[ticker] = s.rating;
      if (close == null) {
        perStock[ticker] = { close: null, ret: null, rating: manualLastRating[ticker], stale: true };
        continue;
      }
      perStock[ticker] = {
        close,
        ret: (close / manualEntry[ticker] - 1) * 100,
        rating: manualLastRating[ticker],
        hardFailed: !!s?.hardFailed,
        stale,
      };
    }
    manualDailyPerStock[day.date] = perStock;
  }

  return {
    cohort,
    anchorDate,
    points,
    manualEntry,
    manualPicksAll: manualAll,
    manualPicksInUni: manualInUni,
    manualDailyPerStock,
    hasManual: Object.keys(manualEntry).length > 0,
    hasNifty: niftyAtEntry != null,
  };
}

function avgFactorFromPerStock(perStock, entryCloses) {
  if (!perStock || !entryCloses) return null;
  let sum = 0, count = 0;
  for (const [ticker, ec] of Object.entries(entryCloses)) {
    if (ec == null) continue;
    const ps = perStock[ticker];
    if (!ps || ps.close == null) continue;
    sum += ps.close / ec;
    count++;
  }
  return count > 0 ? sum / count : null;
}

const COHORT_COLOR = { ai: "#6366f1", manual: "#f59e0b", nifty: "#64748b" };

// Banner shown above the History tab when the analyst's current pillar
// weights differ from v1 defaults. Lets them opt in to recomputing
// history with their custom weights instead of seeing snapshot
// composites scored at v1. Silently absent when weights match default.
function renderPillarWeightsBanner() {
  const w = state.pillarWeights;
  const d = composite.PILLAR_WEIGHTS;
  if (weightsMatchDefault(w)) return "";
  const recomputed = !!state.recomputeHistory;
  const curWeights = `${w.fundamentals}·${w.technicals}·${w.macro}·${w.sentiment}·${w.liquidity}`;
  const defaultWeights = `${d.fundamentals}·${d.technicals}·${d.macro}·${d.sentiment}·${d.liquidity}`;
  if (recomputed) {
    return `
      <div class="rounded-xl bg-indigo-50 ring-1 ring-indigo-200 px-3 py-2 mb-3 flex flex-wrap items-center gap-2">
        <span class="text-indigo-700 text-base leading-none">⚙</span>
        <div class="flex-1 min-w-0">
          <div class="text-xs font-bold text-indigo-900">History recomputed with your custom weights (${curWeights})</div>
          <div class="text-[10px] text-indigo-700/80 mt-0.5">All composite scores + ratings below were re-derived from each snapshot's stored per-pillar percentages × your weights.</div>
        </div>
        <button id="pillar-weights-revert-btn" type="button" class="px-2.5 py-1 text-[11px] font-semibold rounded-lg bg-white text-indigo-700 ring-1 ring-indigo-300 hover:bg-indigo-100">↺ Revert to v1 (${defaultWeights})</button>
      </div>`;
  }
  return `
    <div class="rounded-xl bg-amber-50 ring-1 ring-amber-200 px-3 py-2 mb-3 flex flex-wrap items-center gap-2">
      <span class="text-amber-700 text-base leading-none">⚠</span>
      <div class="flex-1 min-w-0">
        <div class="text-xs font-bold text-amber-900">History below is scored at v1 weights (${defaultWeights}), not your custom weights (${curWeights})</div>
        <div class="text-[10px] text-amber-700/80 mt-0.5">Snapshots are stored with the framework defaults. Opt in to re-derive every composite + rating using your current weights.</div>
      </div>
      <button id="pillar-weights-recompute-btn" type="button" class="px-2.5 py-1 text-[11px] font-semibold rounded-lg bg-amber-600 text-white hover:bg-amber-700 shadow-sm">Recompute with my weights →</button>
    </div>`;
}

// Prominent header bar that toggles the History sub-view (History vs
// Accuracy). Rendered as a real tab strip with an active underline +
// a bell-icon dropdown on the right surfacing every recent hit. Badge
// count = today's fresh hits.
function renderHistoryViewSwitch(activeView, accuracyData, cohortView) {
  const allHits = accuracyData?.allHits || [];
  const todayCount = accuracyData?.todayHits?.length || 0;
  const tabBtn = (view, label, sub) => {
    const active = activeView === view;
    return `
      <button data-history-view="${view}" type="button" class="relative px-4 py-2.5 text-sm font-semibold transition ${active ? "text-indigo-700" : "text-slate-500 hover:text-slate-900"}">
        ${label}
        ${sub ? `<span class="ml-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">${sub}</span>` : ""}
        ${active ? `<span class="absolute left-3 right-3 -bottom-px h-0.5 bg-indigo-600 rounded-full"></span>` : ""}
      </button>`;
  };
  // Cohort view pills (Static / Monthly / Weekly). Lifted from inside
  // the Performance Tracker card up here so they sit with the main
  // section controls — they drive what's shown below, so they belong
  // in the header.
  const cohortBtnCls = (active) => `px-2.5 py-1 rounded-md transition ${active ? "bg-white text-indigo-700 shadow-sm" : "text-slate-600 hover:text-slate-900"}`;
  const cohortToggle = `
    <div id="cohort-view-toggle" class="inline-flex bg-slate-100 rounded-lg p-0.5 text-[11px] font-semibold">
      <button data-view="static" type="button" class="${cohortBtnCls(cohortView === "static")}" title="AI top 7 picked on the first trading day of the month · held all month">Monthly</button>
      <button data-view="weekly" type="button" class="${cohortBtnCls(cohortView === "weekly")}" title="AI re-locks every Monday">Weekly</button>
    </div>
  `;
  const dropdown = allHits.length ? `
    <div id="hits-dropdown" class="hidden absolute right-0 top-full mt-1.5 w-80 bg-white rounded-xl ring-1 ring-slate-200 shadow-2xl z-50 max-h-96 overflow-y-auto">
      <div class="sticky top-0 px-3 py-2 bg-slate-50 border-b border-slate-100 text-[11px] font-bold uppercase tracking-wider text-slate-600 flex items-center justify-between">
        <span>Recent hits · ${allHits.length} total</span>
        ${todayCount > 0 ? `<span class="inline-flex items-center px-1.5 py-0 rounded bg-amber-500 text-white text-[9px] font-bold uppercase tracking-wider animate-pulse">${todayCount} today</span>` : ""}
      </div>
      <div class="py-1">
        ${allHits.map((h) => `
          <div class="px-3 py-1.5 hover:bg-slate-50 text-xs ${h.hitToday ? "bg-amber-50/40" : ""}">
            <div class="flex items-center justify-between gap-2">
              <span class="font-semibold text-slate-900 truncate">${escapeHtml(h.name || h.ticker)}</span>
              <span class="text-[10px] tabular-nums text-slate-400 whitespace-nowrap">${fmtDateDMY(h.hitDate)}</span>
            </div>
            <div class="text-[10px] text-slate-500 mt-0.5">
              <span class="inline-flex items-center gap-1">
                <span class="w-1.5 h-1.5 rounded-full" style="background:${h.basket === "AI" ? COHORT_COLOR.ai : COHORT_COLOR.manual}"></span>
                ${escapeHtml(h.basket)}
              </span>
              ·
              <span class="${h.status === "TARGET_HIT" ? "text-emerald-700 font-bold" : "text-rose-700 font-bold"}">${h.status === "TARGET_HIT" ? "🎯 Target" : "⚠ SL"}</span>
              at ₹${formatPrice(h.exitPrice)} · ${h.daysToHit}d
              ${h.hitToday ? `<span class="ml-1 inline-flex items-center px-1 py-0 rounded bg-amber-500 text-white text-[8px] font-bold uppercase">just hit</span>` : ""}
            </div>
          </div>`).join("")}
      </div>
    </div>` : `
    <div id="hits-dropdown" class="hidden absolute right-0 top-full mt-1.5 w-80 bg-white rounded-xl ring-1 ring-slate-200 shadow-2xl z-50">
      <div class="px-3 py-4 text-[11px] text-slate-500 text-center">No hits yet — once a stock crosses target or SL, it'll land here.</div>
    </div>`;
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 mb-3 overflow-visible">
      <div class="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 px-2 sm:px-3">
        <div id="history-view-toggle" class="flex items-center gap-1">
          ${tabBtn("history", "History", "past picks")}
          ${tabBtn("accuracy", "Accuracy", "target / sl")}
        </div>
        <div class="flex items-center gap-2 py-1.5">
          ${cohortToggle}
          <div class="relative">
            <button id="hits-alert-btn" type="button" class="relative inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg ring-1 ring-slate-200 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-700">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>
              <span class="hidden sm:inline">Alerts</span>
              ${todayCount > 0
                ? `<span class="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-amber-500 text-white text-[9px] font-bold tabular-nums">${todayCount}</span>`
                : allHits.length ? `<span class="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-slate-200 text-slate-600 text-[9px] font-bold tabular-nums">${allHits.length}</span>` : ""}
            </button>
            ${dropdown}
          </div>
        </div>
      </div>
    </div>`;
}

// Wire the pillar-weights banner buttons (Recompute / Revert). Both
// clear the cached history snapshots so the next renderHistory pulls
// fresh v1 data — the recompute path then re-mutates in place.
function wirePillarWeightsBanner() {
  $("#pillar-weights-recompute-btn")?.addEventListener("click", () => {
    state.recomputeHistory = true;
    saveRecomputeHistory(true);
    delete state.cache.history;
    renderHistory();
  });
  $("#pillar-weights-revert-btn")?.addEventListener("click", () => {
    state.recomputeHistory = false;
    saveRecomputeHistory(false);
    delete state.cache.history;
    renderHistory();
  });
}

// Wire up the History sub-view toggle + the bell-icon dropdown. Idempotent —
// safe to call multiple times since renderHistory rebuilds the DOM.
function wireHistorySubViewSwitch() {
  $$("#history-content [data-history-view]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.historyView;
    if (!v || v === state.historyView) return;
    state.historyView = v;
    saveHistoryView(v);
    renderHistory();
    // After the render lands, scroll the Accuracy section into view so
    // the analyst doesn't have to manually scroll past the Performance
    // Tracker to find the target/SL tracker.
    if (v === "accuracy") {
      requestAnimationFrame(() => {
        $("#accuracy-section")?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    }
  }));
  // Cohort view toggle (Static / Monthly / Weekly) now lives in the
  // main tab strip alongside the History/Accuracy switch.
  $$("#history-content #cohort-view-toggle [data-view]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.view;
    if (!v || v === state.cohortView) return;
    state.cohortView = v;
    saveCohortView(v);
    state.cohortSegmentIdx = null;
    renderHistory();
  }));
  const alertBtn = $("#hits-alert-btn");
  const dropdown = $("#hits-dropdown");
  if (alertBtn && dropdown) {
    alertBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      dropdown.classList.toggle("hidden");
    });
    document.addEventListener("click", (e) => {
      if (!dropdown.classList.contains("hidden") && !dropdown.contains(e.target) && !alertBtn.contains(e.target)) {
        dropdown.classList.add("hidden");
      }
    }, { once: true });
  }
}

// ---------------- Accuracy view ----------------
// Per-stock target / stop-loss tracker. AI picks use the framework
// uniform +5% target / −20% SL (from the founder call). Manual picks
// use TGT1 + SL columns from the LKP upload (per-stock).
//
// For each pick we scan forward through the snapshot trail from its
// entry date and record the FIRST close that crossed target or SL.
// Status: TARGET_HIT, SL_HIT, OPEN. Includes daysToHit and a flag
// for "hit happened in today's snapshot" so the UI can highlight
// fresh outcomes.
const AI_TARGET_PCT = 0.05;      // +5%
const AI_SL_PCT     = 0.20;      // −20%

// Detection runs on daily CLOSES for past days -- all a snapshot stores --
// but for TODAY it widens to the live intraday HIGH / LOW, matching what
// buildSegmentedEquityCurve already did. Without this the two disagreed:
// the basket curve booked Deep Industries at its +5% target off an intraday
// high of 646, while this panel called the same position OPEN because the
// close had not printed yet. A level that trades is a level that fills.
function computeHitStatus(ticker, entryDate, entryPrice, target, sl, snapshots, todayDate) {
  if (!ticker || entryPrice == null || target == null || sl == null) {
    return { status: "OPEN", currentClose: null };
  }
  const liveQ = liveQuotesToday();
  const live = liveQ ? state.cache.history?.livePrices?.[ticker] : null;
  // Walk forward from entry day onwards looking for the first close
  // that hits either side.
  let lastClose = null;
  let prevDateBeforeToday = null;
  for (const snap of snapshots) {
    if (snap.date < entryDate) continue;
    const s = snap.stocks.find((x) => x.ticker === ticker);
    if (!s || s.close == null) continue;
    lastClose = s.close;
    // Today only: a target touched intraday counts as hit even if the
    // close came back under it.
    const isToday = snap.date === todayDate || (liveQ && snap.date === snapshots[snapshots.length - 1].date);
    const hi = isToday && live?.dayHigh != null ? Math.max(s.close, live.dayHigh) : s.close;
    const lo = isToday && live?.dayLow != null ? Math.min(s.close, live.dayLow) : s.close;
    if (hi >= target) {
      return {
        status: "TARGET_HIT",
        hitDate: snap.date,
        daysToHit: daysBetween(entryDate, snap.date),
        exitPrice: target,
        intraday: hi > s.close,
        hitToday: true,
      };
    }
    if (lo <= sl) {
      return {
        status: "SL_HIT",
        hitDate: snap.date,
        daysToHit: daysBetween(entryDate, snap.date),
        exitPrice: sl,
        intraday: lo < s.close,
        hitToday: true,
      };
    }
    if (s.close >= target) {
      return {
        status: "TARGET_HIT",
        hitDate: snap.date,
        daysToHit: daysBetween(entryDate, snap.date),
        exitPrice: s.close,
        hitToday: snap.date === todayDate,
      };
    }
    if (s.close <= sl) {
      return {
        status: "SL_HIT",
        hitDate: snap.date,
        daysToHit: daysBetween(entryDate, snap.date),
        exitPrice: s.close,
        hitToday: snap.date === todayDate,
      };
    }
    if (snap.date < todayDate) prevDateBeforeToday = snap.date;
  }
  return { status: "OPEN", currentClose: lastClose };
}

// Manual-basket milestone detection. Unlike the AI cohort's single band,
// the client basket has TWO targets — T1 (first target) and T2 (the max /
// freeze level) — plus SL. We surface T1 as a milestone while the pick
// keeps running toward T2, and freeze/book at T2 (or SL). Detection is on
// the daily CLOSE for past days (all we store historically), but for TODAY
// it widens to the live intraday HIGH / LOW from Munshot when available —
// so a target that was only TOUCHED intraday (the client's definition of
// "hit": Paytm spiked to 1359 over T2 1350 but closed 1341.8) still counts.
// Manual-basket booking. The desk closes at the FIRST target (T1) or SL —
// whichever the price reaches first — and freezes at that LEVEL (not the
// overshooting close). Detection is on the daily CLOSE for past days (all
// we store historically); for TODAY it widens to the live intraday HIGH /
// LOW from Munshot, so a target only TOUCHED intraday still counts as hit
// (the client's definition: Paytm spiked over its target then closed under).
function computeManualMilestones(ticker, entryDate, entryPrice, target, sl, snapshots, todayDate, live) {
  const out = { status: "OPEN", currentClose: null };
  if (!ticker || entryPrice == null || target == null || sl == null) return out;
  let lastClose = null;
  for (const snap of snapshots) {
    if (snap.date < entryDate) continue;
    const s = snap.stocks.find((x) => x.ticker === ticker);
    const isToday = snap.date === todayDate;
    let close = (s && s.close != null) ? s.close : null;
    if (close == null && isToday && live?.current != null) close = live.current;
    if (close == null) continue;
    lastClose = close;
    // Reach range for the day: historically we only have the close (hi=lo=
    // close); today we widen to the live intraday high/low + current.
    let hi = close, lo = close;
    if (isToday && live) {
      if (live.dayHigh != null) hi = Math.max(hi, live.dayHigh);
      if (live.dayLow  != null) lo = Math.min(lo, live.dayLow);
      if (live.current != null) { hi = Math.max(hi, live.current); lo = Math.min(lo, live.current); }
    }
    // Booking event — first of Target 1 (high) or SL (low). Freeze there.
    if (hi >= target) {
      return { ...out, status: "TARGET_HIT", hitDate: snap.date, daysToHit: daysBetween(entryDate, snap.date), exitPrice: target, hitToday: isToday, currentClose: lastClose };
    }
    if (lo <= sl) {
      return { ...out, status: "SL_HIT", hitDate: snap.date, daysToHit: daysBetween(entryDate, snap.date), exitPrice: sl, hitToday: isToday, currentClose: lastClose };
    }
  }
  out.currentClose = lastClose;
  return out;
}

// Per-stock peak excursion — independent of target / SL. Walks the
// snapshot trail from the entry date and records the best (highest
// close) and worst (lowest close) the stock reached relative to entry,
// plus how many days each took. Surfaces "how high did it go, and when"
// so the desk can judge the optimal rebalance horizon (founder ask:
// max upside + max downside + days-to-peak, stockwise). The entry day
// itself counts as 0% at day 0, so maxUpside ≥ 0 and maxDownside ≤ 0.
function computePeakStats(ticker, entryDate, entryPrice, snapshots, todayDate) {
  if (!ticker || entryPrice == null || !entryDate) return null;
  let maxUpPct = null, maxUpDate = null;
  let maxDownPct = null, maxDownDate = null;
  for (const snap of snapshots) {
    if (snap.date < entryDate) continue;
    if (todayDate && snap.date > todayDate) break;
    const s = snap.stocks.find((x) => x.ticker === ticker);
    if (!s || s.close == null) continue;
    const pct = (s.close / entryPrice - 1) * 100;
    if (maxUpPct == null || pct > maxUpPct) { maxUpPct = pct; maxUpDate = snap.date; }
    if (maxDownPct == null || pct < maxDownPct) { maxDownPct = pct; maxDownDate = snap.date; }
  }
  if (maxUpPct == null) return null;
  return {
    maxUpsidePct: maxUpPct,
    daysToMaxUpside: daysBetween(entryDate, maxUpDate),
    maxUpsideDate: maxUpDate,
    maxDownsidePct: maxDownPct,
    daysToMaxDownside: daysBetween(entryDate, maxDownDate),
    maxDownsideDate: maxDownDate,
  };
}

// Build the rows + summary for the Accuracy view.
//   cohort      → the AI cohort currently active (Static/Monthly/Weekly).
//                  Each segment's top 7 contributes 7 picks (5% TGT / 20% SL).
//                  In Weekly mode that's one row per (week × stock), so the
//                  table grows as more weeks accumulate (founder's ask).
//   manualPicks → the LKP basket (in-universe rows have ticker + tgt1 + sl).
//   snapshots   → full snapshot trail.
//
// Rows are scored by proximity to the target / stop-loss bands:
//   proximityScore = (currentReturn% − slPct) / (targetPct − slPct)
//     1.0 = at target · 0.5 = midpoint · 0.0 = at SL
// Sort order: TARGET_HIT (top) → OPEN by proximityScore desc → SL_HIT
// (bottom). Greens up high, reds down low.
function buildAccuracyData(cohort, manualPicks, snapshots) {
  if (!snapshots.length) return null;
  const todayDate = snapshots[snapshots.length - 1].date;

  // Resolve a row's status, current-close return %, and proximity score
  // — all needed for the sort + tint pass below.
  function enrichRow(row) {
    if (row.notCovered) return row;
    const range = row.targetPct - row.slPct;                          // e.g. 5 − (−20) = 25
    const curRet = row.currentClose != null && row.entryPrice
      ? (row.currentClose / row.entryPrice - 1) * 100
      : null;
    let proximity = null;
    if (row.status === "TARGET_HIT") proximity = 1.5;                 // pinned top
    else if (row.status === "SL_HIT") proximity = -0.5;               // pinned bottom
    else if (curRet != null && range > 0) proximity = (curRet - row.slPct) / range;
    return { ...row, currentReturnPct: curRet, proximity };
  }

  // -- AI picks --
  const aiRowsRaw = [];
  for (const seg of (cohort?.segments || [])) {
    for (const s of seg.top7) {
      if (!s.ticker || s.close == null) continue;
      const target = s.close * (1 + AI_TARGET_PCT);
      const sl = s.close * (1 - AI_SL_PCT);
      const status = computeHitStatus(s.ticker, seg.startDate, s.close, target, sl, snapshots, todayDate);
      aiRowsRaw.push({
        ticker: s.ticker,
        name: s.name || s.ticker,
        entryDate: seg.startDate,
        entryPrice: s.close,
        target, sl,
        targetPct: AI_TARGET_PCT * 100,
        slPct: -AI_SL_PCT * 100,
        cohortLabel: cohort?.segments?.length > 1 ? seg.label : null,
        ...status,
      });
    }
  }

  // -- Manual picks --
  const manualRowsRaw = [];
  for (const p of (manualPicks || [])) {
    if (!p.in_universe || !p.ticker) {
      manualRowsRaw.push({
        ticker: p.ticker,
        name: p.selection || p.ticker || "—",
        notCovered: true,
        outReason: p.out_reason || "Below market-cap floor / not matched",
      });
      continue;
    }
    const cohortEntrySnap = cohort?.segments?.[0]?.entrySnap;
    const entryFromSnap = cohortEntrySnap?.stocks?.find((x) => x.ticker === p.ticker)?.close;
    const entryPrice = entryFromSnap ?? p.entry ?? null;
    const entryDate = cohortEntrySnap?.date || cohort?.anchorDate || todayDate;
    const target = p.tgt1 ?? null;
    const sl = p.sl ?? null;
    if (entryPrice == null || target == null || sl == null) {
      manualRowsRaw.push({
        ticker: p.ticker, name: p.selection || p.ticker,
        entryDate, entryPrice, target, sl,
        targetPct: target != null && entryPrice ? (target / entryPrice - 1) * 100 : null,
        slPct: sl != null && entryPrice ? (sl / entryPrice - 1) * 100 : null,
        status: "OPEN", currentClose: null,
      });
      continue;
    }
    const status = computeHitStatus(p.ticker, entryDate, entryPrice, target, sl, snapshots, todayDate);
    manualRowsRaw.push({
      ticker: p.ticker,
      name: p.selection || p.ticker,
      entryDate, entryPrice, target, sl,
      targetPct: (target / entryPrice - 1) * 100,
      slPct: (sl / entryPrice - 1) * 100,
      tgt2: p.tgt2 ?? null,
      ...status,
    });
  }

  // Sort: pinned target-hit at top, OPEN by proximity desc, SL-hit at
  // bottom. Not-Covered float to the very bottom so they don't break
  // the visual flow.
  const sortByProximity = (rows) => rows
    .map(enrichRow)
    .sort((a, b) => {
      if (a.notCovered && !b.notCovered) return 1;
      if (!a.notCovered && b.notCovered) return -1;
      return (b.proximity ?? -99) - (a.proximity ?? -99);
    });
  const aiRows = sortByProximity(aiRowsRaw);
  const manualRows = sortByProximity(manualRowsRaw);

  // Summaries — count target/SL/open per basket.
  const summarise = (rows) => {
    const eligible = rows.filter((r) => !r.notCovered && r.status);
    const targetHits = eligible.filter((r) => r.status === "TARGET_HIT");
    const slHits = eligible.filter((r) => r.status === "SL_HIT");
    const open = eligible.filter((r) => r.status === "OPEN");
    const avgDaysToTarget = targetHits.length
      ? targetHits.reduce((a, b) => a + (b.daysToHit || 0), 0) / targetHits.length
      : null;
    const avgDaysToSL = slHits.length
      ? slHits.reduce((a, b) => a + (b.daysToHit || 0), 0) / slHits.length
      : null;
    return {
      total: eligible.length,
      targetHits: targetHits.length,
      slHits: slHits.length,
      open: open.length,
      targetHitRate: eligible.length ? (targetHits.length / eligible.length) * 100 : null,
      slHitRate: eligible.length ? (slHits.length / eligible.length) * 100 : null,
      avgDaysToTarget, avgDaysToSL,
    };
  };
  const aiSummary = summarise(aiRows);
  const manualSummary = summarise(manualRows);

  // "Just hit today" picks across both baskets — surfaced as an alert.
  const todayHits = [...aiRows, ...manualRows].filter((r) => r.hitToday);

  // All-time hits ledger for the bell-icon dropdown — newest first,
  // labelled with basket.
  const allHits = [
    ...aiRows.filter((r) => !r.notCovered && (r.status === "TARGET_HIT" || r.status === "SL_HIT"))
      .map((r) => ({ ...r, basket: "AI" })),
    ...manualRows.filter((r) => !r.notCovered && (r.status === "TARGET_HIT" || r.status === "SL_HIT"))
      .map((r) => ({ ...r, basket: "Manual" })),
  ].sort((a, b) => (b.hitDate || "").localeCompare(a.hitDate || ""));

  return { aiRows, manualRows, aiSummary, manualSummary, todayHits, allHits, todayDate };
}

function renderAccuracyView(data) {
  if (!data) return `<div class="bg-white rounded-2xl ring-1 ring-slate-100 p-6 text-center text-slate-500 text-sm">No accuracy data yet — snapshots not loaded.</div>`;

  const fmtPct = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
  const fmtDays = (v) => v == null ? "—" : `${Math.round(v)}d`;

  const summaryCard = (label, summary, accent) => `
    <div class="rounded-xl bg-white ring-1 ring-slate-200 p-3">
      <div class="flex items-center gap-1.5 mb-2">
        <span class="inline-block w-2 h-2 rounded-full" style="background:${accent}"></span>
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">${escapeHtml(label)}</div>
        <span class="ml-auto text-[10px] text-slate-400 tabular-nums">${summary.total} pick${summary.total === 1 ? "" : "s"}</span>
      </div>
      <div class="grid grid-cols-3 gap-2 text-center">
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-emerald-700">${summary.targetHits}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">Target hits</div>
          <div class="text-[10px] text-slate-400 tabular-nums">${summary.targetHitRate != null ? summary.targetHitRate.toFixed(0) + "%" : "—"}</div>
        </div>
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-rose-700">${summary.slHits}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">SL hits</div>
          <div class="text-[10px] text-slate-400 tabular-nums">${summary.slHitRate != null ? summary.slHitRate.toFixed(0) + "%" : "—"}</div>
        </div>
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-slate-700">${summary.open}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">Open</div>
          <div class="text-[10px] text-slate-400 tabular-nums">avg ${fmtDays(summary.avgDaysToTarget)}→T · ${fmtDays(summary.avgDaysToSL)}→SL</div>
        </div>
      </div>
    </div>
  `;

  const todayHitsBanner = data.todayHits.length ? `
    <div class="rounded-xl bg-gradient-to-r from-amber-50 to-orange-50 ring-1 ring-amber-200 px-3 py-2 mb-3">
      <div class="flex items-center gap-2">
        <span class="text-amber-600 text-base">🔔</span>
        <span class="font-bold text-amber-900 text-sm">${data.todayHits.length} pick${data.todayHits.length === 1 ? "" : "s"} just hit ${data.todayHits.length === 1 ? "an outcome" : "outcomes"} on ${fmtDateDMY(data.todayDate)}</span>
        <div class="ml-auto flex flex-wrap items-center gap-1.5">
          ${data.todayHits.map((h) => `
            <span class="inline-flex items-center gap-1 px-1.5 py-0.5 rounded ring-1 text-[10px] font-bold ${h.status === "TARGET_HIT" ? "bg-emerald-100 text-emerald-700 ring-emerald-200" : "bg-rose-100 text-rose-700 ring-rose-200"}">
              ${h.status === "TARGET_HIT" ? "🎯" : "⚠"} ${escapeHtml(h.name || h.ticker)}
            </span>
          `).join("")}
        </div>
      </div>
    </div>` : "";

  function renderRow(r, basket) {
    if (r.notCovered) {
      return `
        <div class="grid grid-cols-12 items-center gap-2 py-1.5 px-2 rounded-lg bg-slate-50/40">
          <div class="col-span-5 min-w-0">
            <div class="font-semibold text-slate-500 text-xs truncate" title="${escapeHtml(r.outReason || "")}">${escapeHtml(r.name)}</div>
            <div class="text-[10px] text-slate-400 truncate">${escapeHtml(r.outReason || "")}</div>
          </div>
          <div class="col-span-7 text-right">
            <span class="inline-flex items-center px-1.5 py-0 rounded bg-slate-200 text-slate-600 ring-1 ring-slate-300 text-[9px] font-bold uppercase tracking-wider">Not Covered</span>
          </div>
        </div>`;
    }
    const statusCls = r.status === "TARGET_HIT" ? "bg-emerald-100 text-emerald-700 ring-emerald-200"
      : r.status === "SL_HIT" ? "bg-rose-100 text-rose-700 ring-rose-200"
      : "bg-slate-100 text-slate-600 ring-slate-200";
    const statusLabel = r.status === "TARGET_HIT" ? "🎯 Target Hit"
      : r.status === "SL_HIT" ? "⚠ SL Hit"
      : "Open";
    const hitTodayBadge = r.hitToday
      ? `<span class="inline-flex items-center gap-1 px-1.5 py-0 rounded bg-amber-500 text-white text-[9px] font-bold uppercase tracking-wider ml-1 animate-pulse">JUST HIT</span>`
      : "";
    const exit = r.exitPrice != null ? `₹${formatPrice(r.exitPrice)} · ${fmtDateDMY(r.hitDate)} · ${r.daysToHit}d`
      : r.currentClose != null ? `now ₹${formatPrice(r.currentClose)}` : "—";
    // Proximity tint — greens climb toward target, reds slide toward SL.
    // Tints are subtle so the table reads at a glance without screaming.
    let rowTint = "hover:bg-slate-50";
    if (r.status === "TARGET_HIT") rowTint = "bg-emerald-50 ring-1 ring-emerald-200";
    else if (r.status === "SL_HIT") rowTint = "bg-rose-50 ring-1 ring-rose-200";
    else if (r.proximity != null && r.proximity >= 0.75) rowTint = "bg-emerald-50/40 hover:bg-emerald-50/70";
    else if (r.proximity != null && r.proximity <= 0.25) rowTint = "bg-rose-50/40 hover:bg-rose-50/70";
    if (r.hitToday) rowTint += " ring-2 ring-amber-300";
    const currentReturnHtml = r.currentReturnPct != null
      ? `<span class="text-[10px] tabular-nums font-semibold ${r.currentReturnPct >= 0 ? "text-emerald-700" : "text-rose-700"} ml-1">${fmtPct(r.currentReturnPct)}</span>`
      : "";
    // Rows with a ticker are clickable — `data-cohort-row` is the
    // existing hook the History tab uses to open the drill modal
    // (price chart + rating timeline + forensics) via wireCohortHandlers.
    // Same plumbing as the Performance Tracker tables; tying into it
    // means hover affordance, modal layout, and forensic data all
    // come for free.
    const clickable = !!r.ticker;
    const baseCls = `grid grid-cols-12 items-center gap-2 py-1.5 px-2 rounded-lg ${rowTint}`;
    const cellsHtml = `
        <div class="col-span-5 min-w-0">
          <div class="font-semibold text-slate-900 text-xs truncate">${escapeHtml(r.name)}${r.cohortLabel ? ` <span class="text-[9px] text-slate-400 font-normal">· ${escapeHtml(r.cohortLabel)}</span>` : ""}</div>
          <div class="text-[10px] text-slate-500 tabular-nums">Entry ${fmtDateDMY(r.entryDate)} · ₹${formatPrice(r.entryPrice)}${currentReturnHtml}</div>
        </div>
        <div class="col-span-3 text-[10px] tabular-nums text-slate-600 text-right">
          <div>T: ₹${formatPrice(r.target)} <span class="text-emerald-600">${fmtPct(r.targetPct)}</span></div>
          <div>SL: ₹${formatPrice(r.sl)} <span class="text-rose-600">${fmtPct(r.slPct)}</span></div>
        </div>
        <div class="col-span-4 text-right">
          <div class="flex items-center justify-end flex-wrap gap-1">
            <span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 ${statusCls}">${statusLabel}</span>
            ${hitTodayBadge}
          </div>
          <div class="text-[10px] text-slate-500 mt-0.5 truncate">${exit}</div>
        </div>`;
    if (clickable) {
      return `
      <button type="button" data-cohort-row data-cohort-side="${basket}" data-ticker="${escapeHtml(r.ticker)}" data-seg-anchor="${escapeHtml(r.entryDate || "")}" class="w-full text-left ${baseCls} cursor-pointer transition hover:ring-1 hover:ring-indigo-200">
        ${cellsHtml}
      </button>`;
    }
    return `
      <div class="${baseCls}">
        ${cellsHtml}
      </div>`;
  }

  const aiRowsHtml = data.aiRows.length
    ? data.aiRows.map((r) => renderRow(r, "ai")).join("")
    : `<div class="text-[11px] text-slate-400 text-center py-3">No AI picks for this cohort.</div>`;
  const manualRowsHtml = data.manualRows.length
    ? data.manualRows.map((r) => renderRow(r, "manual")).join("")
    : `<div class="text-[11px] text-slate-400 text-center py-3">No manual basket loaded.</div>`;

  return `
    <div id="accuracy-section" class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 sm:p-5 scroll-mt-4">
      <div class="flex flex-wrap items-baseline justify-between gap-2 mb-3">
        <h2 class="font-display font-bold text-slate-900 text-base">Target / Stop-loss tracker</h2>
      </div>
      ${todayHitsBanner}
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
        ${summaryCard("AI Picks · summary", data.aiSummary, COHORT_COLOR.ai)}
        ${summaryCard("Manual Picks · summary", data.manualSummary, COHORT_COLOR.manual)}
      </div>
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <div>
          <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5 flex items-center gap-1.5">
            <span class="inline-block w-2 h-2 rounded-full" style="background:${COHORT_COLOR.ai}"></span>
            AI Picks · per-pick status
          </div>
          <div class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 px-1 py-1">${aiRowsHtml}</div>
        </div>
        <div>
          <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5 flex items-center gap-1.5">
            <span class="inline-block w-2 h-2 rounded-full" style="background:${COHORT_COLOR.manual}"></span>
            Manual Picks · per-pick status
          </div>
          <div class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 px-1 py-1">${manualRowsHtml}</div>
        </div>
      </div>
    </div>
  `;
}

function renderCohortTracker(cohort, series, view, selectedSegIdx) {
  if (!cohort || !series) return "";
  const days = series.points.length - 1;
  const last = series.points[series.points.length - 1];
  const fmtPct = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const sign = (v) => v == null ? "text-slate-500" : v >= 0 ? "text-emerald-700" : "text-rose-700";
  const alpha = (last.aiCum != null && last.niftyCum != null) ? last.aiCum - last.niftyCum : null;

  const statBlock = (label, value, sub, valueCls = "text-slate-900", dotColor = null) => `
    <div class="rounded-xl bg-slate-50 ring-1 ring-slate-100 px-3 py-2">
      <div class="flex items-center gap-1.5">
        ${dotColor ? `<span class="inline-block w-2 h-2 rounded-full" style="background:${dotColor}"></span>` : ""}
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">${escapeHtml(label)}</div>
      </div>
      <div class="text-xl font-display font-extrabold tabular-nums leading-tight mt-0.5 ${valueCls}">${value}</div>
      <div class="text-[10px] text-slate-500 mt-0.5 truncate" title="${escapeHtml(sub)}">${escapeHtml(sub)}</div>
    </div>
  `;
  const aiSub = view === "weekly"
    ? `${cohort.segments.length} week${cohort.segments.length === 1 ? "" : "s"} · re-locks each Mon`
    : view === "static"
      ? `Frozen at prior month-end · ${days}d`
      : `Held since upload · ${days}d`;
  const stats = `
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mb-3">
      ${statBlock("AI Picks", fmtPct(last.aiCum), aiSub, sign(last.aiCum), COHORT_COLOR.ai)}
      ${statBlock(
        "Manual Picks",
        series.hasManual ? fmtPct(last.manualCum) : "—",
        series.hasManual ? `${series.manualPicksInUni.length} in coverage · static basket` : "Upload basket below",
        sign(last.manualCum),
        COHORT_COLOR.manual,
      )}
      ${statBlock(
        "Smallcap 250",
        series.hasNifty ? fmtPct(last.niftyCum) : "—",
        series.hasNifty ? `Benchmark over same window` : "Loading daily",
        sign(last.niftyCum),
        COHORT_COLOR.nifty,
      )}
      ${statBlock("Alpha vs Nifty", alpha == null ? "—" : fmtPct(alpha), "AI − Nifty", sign(alpha))}
    </div>
  `;

  // Weekly mode: pill picker for selecting which segment's table to show.
  const segPills = view === "weekly" && cohort.segments.length > 0 ? `
    <div class="flex flex-wrap items-center gap-1 mb-3 overflow-x-auto">
      <span class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mr-1">AI basket per week</span>
      ${cohort.segments.map((seg, i) => `
        <button data-seg="${i}" type="button" class="px-2.5 py-1 rounded-lg text-[11px] font-semibold whitespace-nowrap transition ${i === selectedSegIdx ? "bg-indigo-600 text-white shadow-sm" : "bg-white text-slate-600 ring-1 ring-slate-200 hover:ring-indigo-300 hover:text-indigo-700"}">
          ${escapeHtml(seg.label)} · ${fmtDateDM(seg.startDate)}–${fmtDateDM(seg.endDate)}
        </button>
      `).join("")}
    </div>` : "";

  // AI table — for the SELECTED segment (week pill in weekly mode, the
  // single segment in monthly).
  const seg = cohort.segments[selectedSegIdx] || cohort.segments[cohort.segments.length - 1];
  const segLastDate = seg.tracking[seg.tracking.length - 1]?.date;
  const segLastPerStock = seg.dailyPerStock[segLastDate] || {};
  const aiRows = seg.top7.map((s) => {
    const cur = segLastPerStock[s.ticker];
    const ret = cur?.ret;
    const retCls = ret == null ? "text-slate-500" : ret >= 0 ? "text-emerald-700" : "text-rose-700";
    const rating = cur?.rating || s.rating;
    const ratingCls = COHORT_RATING_BG[rating] || "bg-slate-100 text-slate-700 ring-slate-200";
    const fail = !!cur?.hardFailed;
    return `
      <button type="button" data-cohort-row data-cohort-side="ai" data-ticker="${escapeHtml(s.ticker)}" data-seg-anchor="${escapeHtml(seg.startDate)}" class="w-full text-left grid grid-cols-12 items-center gap-2 py-1.5 px-2 rounded-lg cursor-pointer transition ${fail ? "bg-rose-50/40 ring-1 ring-rose-200 hover:ring-rose-300" : "hover:bg-indigo-50/40 hover:ring-1 hover:ring-indigo-200"}">
        <div class="col-span-6 sm:col-span-5 min-w-0">
          <div class="font-semibold text-slate-900 text-xs truncate" title="${escapeHtml(s.name || s.ticker)}">${escapeHtml(s.name || s.ticker)}</div>
          <div class="text-[10px] text-slate-500 tabular-nums">₹${formatPrice(s.close)} → ₹${cur?.close != null ? formatPrice(cur.close) : "—"}</div>
        </div>
        <div class="col-span-3 sm:col-span-3 text-right tabular-nums text-[11px] font-bold ${retCls}">${ret == null ? "—" : (ret >= 0 ? "+" : "") + ret.toFixed(2) + "%"}</div>
        <div class="col-span-3 sm:col-span-2 text-right">
          <span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap ${ratingCls}">${escapeHtml(rating || "—")}</span>
        </div>
        <div class="hidden sm:block col-span-2 text-right">
          ${fail ? `<span class="inline-flex items-center gap-1 px-1.5 py-0 rounded bg-rose-100 text-rose-700 ring-1 ring-rose-200 text-[9px] font-bold uppercase tracking-wider">⚠ hard-fail</span>` : ""}
        </div>
      </button>`;
  }).join("");
  const aiHeader = view === "weekly"
    ? `${escapeHtml(seg.label)} · ${fmtDateDM(seg.startDate)} → ${fmtDateDM(seg.endDate)} · 7 stocks`
    : `Held since ${fmtDateDMY(cohort.effectiveStart)} · 7 stocks`;

  // Manual table — ALWAYS shows the same basket (founder: client updates
  // monthly even in weekly mode). Out-of-universe picks now appear here
  // with a "NOT COVERED" badge instead of being silently dropped.
  const lastManualPerStock = series.manualDailyPerStock[segLastDate] || series.manualDailyPerStock[series.anchorDate] || {};
  const manualRows = series.manualPicksAll.length
    ? series.manualPicksAll.map((p) => {
        if (!p.in_universe || !p.ticker || series.manualEntry[p.ticker] == null) {
          // Not-covered: shown with reason; doesn't enter the average.
          return `
            <div class="grid grid-cols-12 items-center gap-2 py-1.5 px-2 rounded-lg bg-slate-50/40">
              <div class="col-span-8 min-w-0">
                <div class="font-semibold text-slate-500 text-xs truncate" title="${escapeHtml(p.out_reason || "Not in our coverage universe")}">${escapeHtml(p.selection || p.ticker || "—")}</div>
                <div class="text-[10px] text-slate-400 truncate">${escapeHtml(p.out_reason || "Below market-cap floor / not matched")}</div>
              </div>
              <div class="col-span-4 text-right">
                <span class="inline-flex items-center px-1.5 py-0 rounded bg-slate-200 text-slate-600 ring-1 ring-slate-300 text-[9px] font-bold uppercase tracking-wider">Not Covered</span>
              </div>
            </div>`;
        }
        const cur = lastManualPerStock[p.ticker];
        const entryClose = series.manualEntry[p.ticker];
        const ret = cur?.ret;
        const retCls = ret == null ? "text-slate-500" : ret >= 0 ? "text-emerald-700" : "text-rose-700";
        return `
          <button type="button" data-cohort-row data-cohort-side="manual" data-ticker="${escapeHtml(p.ticker)}" class="w-full text-left grid grid-cols-12 items-center gap-2 py-1.5 px-2 rounded-lg cursor-pointer transition hover:bg-amber-50/40 hover:ring-1 hover:ring-amber-200">
            <div class="col-span-8 min-w-0">
              <div class="font-semibold text-slate-900 text-xs truncate">${escapeHtml(p.selection || p.ticker)}</div>
              <div class="text-[10px] text-slate-500 tabular-nums">₹${entryClose != null ? formatPrice(entryClose) : "—"} → ₹${cur?.close != null ? formatPrice(cur.close) : "—"}</div>
            </div>
            <div class="col-span-4 text-right tabular-nums text-[11px] font-bold ${retCls}">${ret == null ? "—" : (ret >= 0 ? "+" : "") + ret.toFixed(2) + "%"}</div>
          </button>`;
      }).join("")
    : `<div class="text-[11px] text-slate-400 text-center py-4 leading-relaxed">No client basket loaded.<br>Use the LKP picks card below to upload one.</div>`;

  // Static mode shows BOTH the selection day (prior month-end snapshot)
  // and the cost-basis day (first held-month snapshot) so the analyst
  // sees the apples-to-apples reading: "screener said this on 31-05,
  // assume client bought 01-06". Monthly/Weekly views have selection
  // === cost basis, so we just show one date.
  const anchorLabel = cohort.anchorMode === "static" && cohort.selectionDate && cohort.selectionDate !== cohort.effectiveStart
    ? `selected ${fmtDateDMY(cohort.selectionDate)} · cost basis ${fmtDateDMY(cohort.effectiveStart)}`
    : `${fmtDateDMY(cohort.anchorDate)}${cohort.effectiveStart !== cohort.anchorDate ? ` · snapped to first trading day ${fmtDateDMY(cohort.effectiveStart)}` : ""}`;
  const anchorSourceLabel = cohort.anchorMode === "static"
    ? "frozen at prior month-end"
    : cohort.anchorMode === "weekly"
      ? "weekly re-lock since upload"
      : "client upload date";
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 sm:p-5 mb-4">
      <div class="mb-3">
        <h2 class="font-display font-bold text-slate-900 text-base">Performance Tracker</h2>
        <div class="text-[11px] text-slate-500 mt-0.5">AI Picks vs Manual Picks vs Nifty · ${anchorSourceLabel} (<span class="font-semibold">${anchorLabel}</span>) · ${days} trading day${days === 1 ? "" : "s"}</div>
      </div>

      ${stats}

      <div id="cohort-chart-container" class="relative rounded-xl bg-slate-50 ring-1 ring-slate-100 p-3 mb-3">
        ${renderCohortChart(series, view)}
        <div id="cohort-tooltip" class="hidden absolute z-10 pointer-events-none -translate-x-1/2 -translate-y-[calc(100%+12px)] bg-slate-900/95 backdrop-blur text-white text-xs rounded-xl shadow-2xl ring-1 ring-slate-700/60 px-3 py-2 whitespace-nowrap"></div>
      </div>

      ${segPills}

      <div class="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <div>
          <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5 flex items-center gap-1.5">
            <span class="inline-block w-2 h-2 rounded-full" style="background:${COHORT_COLOR.ai}"></span>
            AI Picks · <span class="text-slate-400 normal-case font-medium">${escapeHtml(aiHeader)}</span>
          </div>
          <div class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 px-1 py-1">${aiRows}</div>
        </div>
        <div>
          <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5 flex items-center gap-1.5">
            <span class="inline-block w-2 h-2 rounded-full" style="background:${COHORT_COLOR.manual}"></span>
            Manual Picks · <span class="text-slate-400 normal-case font-medium">${series.manualPicksAll.length} stocks · ${series.manualPicksInUni.length} in coverage</span>
          </div>
          <div class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 px-1 py-1">${manualRows}</div>
        </div>
      </div>
    </div>
  `;
}

// 3-line cohort chart — AI / Manual / Nifty. In weekly mode, vertical
// dotted boundary lines mark each AI re-balance ("rebalance" markers).
// Hidden hover-guide elements + a capture rect so setupCohortHover()
// can wire crosshair + tooltip after render.
function renderCohortChart(series, view) {
  const pts = series.points;
  if (pts.length < 2) return `<div class="text-[11px] text-slate-400 text-center py-6">Not enough days to draw a curve yet — come back tomorrow.</div>`;
  const W = 800, H = 200;
  const M = { top: 14, right: 16, bottom: 30, left: 44 };
  const innerW = W - M.left - M.right;
  const innerH = H - M.top - M.bottom;
  const allRets = [];
  for (const p of pts) {
    if (p.aiCum != null) allRets.push(p.aiCum);
    if (p.manualCum != null) allRets.push(p.manualCum);
    if (p.niftyCum != null) allRets.push(p.niftyCum);
  }
  const yMin = Math.min(0, ...allRets);
  const yMax = Math.max(0, ...allRets);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.12, yHi = yMax + ySpan * 0.12;
  const xAt = (i) => M.left + (i / (pts.length - 1)) * innerW;
  const yAt = (v) => M.top + innerH - ((v - yLo) / (yHi - yLo)) * innerH;

  function linePath(key) {
    let d = "";
    let pending = "M";
    for (let i = 0; i < pts.length; i++) {
      const v = pts[i][key];
      if (v == null) { pending = "M"; continue; }
      d += `${pending} ${xAt(i).toFixed(2)} ${yAt(v).toFixed(2)} `;
      pending = "L";
    }
    return d.trim();
  }

  const aiPath = linePath("aiCum");
  const manualPath = series.hasManual ? linePath("manualCum") : "";
  const niftyPath = series.hasNifty ? linePath("niftyCum") : "";

  // Filled area under AI line for visual prominence
  let aiArea = "";
  const aiPoints = pts.map((p, i) => p.aiCum == null ? null : [xAt(i), yAt(p.aiCum)]).filter(Boolean);
  if (aiPoints.length >= 2) {
    const first = aiPoints[0], lastPt = aiPoints[aiPoints.length - 1];
    const baseY = yAt(0).toFixed(2);
    const d = aiPoints.map(([x, y], i) => `${i === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`).join(" ");
    aiArea = `${d} L ${lastPt[0].toFixed(2)} ${baseY} L ${first[0].toFixed(2)} ${baseY} Z`;
  }

  // Vertical boundary markers — weekly only.
  const boundaryLines = view === "weekly"
    ? pts.map((p, i) => p.isBoundary
        ? `<line x1="${xAt(i).toFixed(2)}" x2="${xAt(i).toFixed(2)}" y1="${M.top}" y2="${(M.top + innerH).toFixed(2)}" stroke="#94a3b8" stroke-width="0.8" stroke-dasharray="3 4" opacity="0.6"><title>Rebalance — AI basket changed (${p.date})</title></line>`
        : "").join("")
    : "";
  // Tiny label above each boundary line
  const boundaryLabels = view === "weekly"
    ? pts.map((p, i) => p.isBoundary
        ? `<text x="${xAt(i).toFixed(2)}" y="${(M.top - 2).toFixed(2)}" text-anchor="middle" font-size="8" font-weight="600" fill="#94a3b8">↻</text>`
        : "").join("")
    : "";

  // Y-axis
  const yTicks = [0, 1, 2, 3, 4].map((step) => {
    const v = yLo + (yHi - yLo) * (step / 4);
    const yy = (M.top + innerH - (step / 4) * innerH).toFixed(2);
    const isZero = Math.abs(v) < 0.05;
    return `
      <line x1="${M.left}" x2="${(W - M.right).toFixed(2)}" y1="${yy}" y2="${yy}" stroke="${isZero ? "#94a3b8" : "#e2e8f0"}" stroke-width="${isZero ? 0.9 : 0.6}" stroke-dasharray="${isZero ? "0" : "3 4"}" />
      <text x="${(M.left - 8).toFixed(2)}" y="${yy}" text-anchor="end" dominant-baseline="middle" font-size="10" font-weight="500" fill="#94a3b8">${v >= 0 ? "+" : ""}${v.toFixed(1)}%</text>`;
  }).join("");

  // X-axis date ticks (DD-MM, Indian convention)
  const tickEvery = Math.max(1, Math.ceil(pts.length / 6));
  const xTicks = pts.map((p, i) => {
    if (i % tickEvery !== 0 && i !== pts.length - 1) return "";
    return `<text x="${xAt(i).toFixed(2)}" y="${(M.top + innerH + 16).toFixed(2)}" text-anchor="middle" font-size="10" fill="#64748b">${fmtDateDM(p.date)}</text>`;
  }).join("");

  return `
    <svg id="cohort-svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" class="w-full select-none" style="max-height:240px">
      <defs>
        <linearGradient id="cohortAiArea" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="${COHORT_COLOR.ai}" stop-opacity="0.22"/>
          <stop offset="100%" stop-color="${COHORT_COLOR.ai}" stop-opacity="0"/>
        </linearGradient>
      </defs>
      ${yTicks}
      ${boundaryLines}
      ${aiArea ? `<path d="${aiArea}" fill="url(#cohortAiArea)" />` : ""}
      ${niftyPath ? `<path d="${niftyPath}" fill="none" stroke="${COHORT_COLOR.nifty}" stroke-width="1.6" stroke-linejoin="round" stroke-dasharray="4 4" stroke-linecap="round" />` : ""}
      ${manualPath ? `<path d="${manualPath}" fill="none" stroke="${COHORT_COLOR.manual}" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round" />` : ""}
      <path d="${aiPath}" fill="none" stroke="${COHORT_COLOR.ai}" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round" />
      ${boundaryLabels}
      ${xTicks}
      <line id="cohort-guide" x1="0" y1="${M.top}" x2="0" y2="${(M.top + innerH).toFixed(2)}" stroke="#94a3b8" stroke-width="0.7" stroke-dasharray="2 3" opacity="0" />
      <circle id="cohort-hover-ai" cx="0" cy="0" r="4" fill="#fff" stroke="${COHORT_COLOR.ai}" stroke-width="2" opacity="0" />
      <circle id="cohort-hover-manual" cx="0" cy="0" r="3" fill="#fff" stroke="${COHORT_COLOR.manual}" stroke-width="2" opacity="0" />
      <circle id="cohort-hover-nifty" cx="0" cy="0" r="3" fill="#fff" stroke="${COHORT_COLOR.nifty}" stroke-width="2" opacity="0" />
      <rect id="cohort-hover-capture" x="0" y="0" width="${W}" height="${H}" fill="transparent" />
    </svg>
  `;
}

function setupCohortHover(series) {
  const W = 800, H = 200;
  const M = { left: 44, right: 16, top: 14, bottom: 30 };
  const innerW = W - M.left - M.right;
  const pts = series.points;
  if (pts.length < 2) return;
  const allRets = [];
  for (const p of pts) {
    if (p.aiCum != null) allRets.push(p.aiCum);
    if (p.manualCum != null) allRets.push(p.manualCum);
    if (p.niftyCum != null) allRets.push(p.niftyCum);
  }
  const yMin = Math.min(0, ...allRets);
  const yMax = Math.max(0, ...allRets);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.12, yHi = yMax + ySpan * 0.12;
  const innerH = H - M.top - M.bottom;
  const xAt = (i) => M.left + (i / (pts.length - 1)) * innerW;
  const yAt = (v) => M.top + innerH - ((v - yLo) / (yHi - yLo)) * innerH;

  const container = $("#cohort-chart-container");
  const svg = $("#cohort-svg");
  const capture = $("#cohort-hover-capture");
  const guide = $("#cohort-guide");
  const ha = $("#cohort-hover-ai");
  const hm = $("#cohort-hover-manual");
  const hn = $("#cohort-hover-nifty");
  const tip = $("#cohort-tooltip");
  if (!container || !svg || !capture || !tip) return;

  function show(idx) {
    const p = pts[idx];
    const x = xAt(idx);
    guide.setAttribute("x1", x); guide.setAttribute("x2", x); guide.setAttribute("opacity", "1");
    if (p.aiCum != null) {
      ha.setAttribute("cx", x); ha.setAttribute("cy", yAt(p.aiCum)); ha.setAttribute("opacity", "1");
    } else ha.setAttribute("opacity", "0");
    if (p.manualCum != null) {
      hm.setAttribute("cx", x); hm.setAttribute("cy", yAt(p.manualCum)); hm.setAttribute("opacity", "1");
    } else hm.setAttribute("opacity", "0");
    if (p.niftyCum != null) {
      hn.setAttribute("cx", x); hn.setAttribute("cy", yAt(p.niftyCum)); hn.setAttribute("opacity", "1");
    } else hn.setAttribute("opacity", "0");

    const line = (label, val, color) =>
      `<div class="flex items-center gap-2">
        <span class="inline-block w-2 h-2 rounded-full" style="background:${color}"></span>
        <span class="text-slate-300">${label}</span>
        <span class="ml-auto tabular-nums font-bold ${val == null ? "text-slate-400" : val >= 0 ? "text-emerald-300" : "text-rose-300"}">${val == null ? "—" : (val >= 0 ? "+" : "") + val.toFixed(2) + "%"}</span>
      </div>`;
    const boundaryNote = p.isBoundary
      ? `<div class="text-[10px] text-amber-300 mt-1">↻ Rebalance (basket changed)</div>` : "";
    tip.innerHTML = `
      <div class="font-bold text-sm leading-tight mb-1">${fmtDateDMY(p.date)}</div>
      <div class="space-y-0.5 min-w-[160px]">
        ${line("AI Picks", p.aiCum, COHORT_COLOR.ai)}
        ${series.hasManual ? line("Manual Picks", p.manualCum, COHORT_COLOR.manual) : ""}
        ${series.hasNifty ? line("Smallcap 250", p.niftyCum, COHORT_COLOR.nifty) : ""}
      </div>
      ${boundaryNote}
    `;
    tip.classList.remove("hidden");
    // Anchor via the live CTM (letterbox / zoom / DPR safe), not x/W*rect.width.
    const cr = container.getBoundingClientRect();
    const ctm = svg.getScreenCTM();
    let px, py;
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = x; sp.y = M.top + innerH / 2;
      const scr = sp.matrixTransform(ctm);
      px = scr.x - cr.left; py = scr.y - cr.top;
    } else {
      const rect = svg.getBoundingClientRect();
      px = (x / W) * rect.width; py = (M.top + innerH / 2) / H * rect.height;
    }
    tip.style.left = `${px}px`;
    tip.style.top = `${py}px`;
    tip.style.transform = "translate(-50%, calc(-100% - 16px))";
    requestAnimationFrame(() => {
      const tr = tip.getBoundingClientRect();
      let dx = 0;
      if (tr.left < cr.left + 6) dx = (cr.left + 6) - tr.left;
      else if (tr.right > cr.right - 6) dx = (cr.right - 6) - tr.right;
      if (dx !== 0) tip.style.transform = `translate(calc(-50% + ${dx}px), calc(-100% - 16px))`;
    });
  }
  function hide() {
    guide.setAttribute("opacity", "0");
    ha.setAttribute("opacity", "0"); hm.setAttribute("opacity", "0"); hn.setAttribute("opacity", "0");
    tip.classList.add("hidden");
  }
  function eventToIdx(e) {
    const t = e.touches ? e.touches[0] : e;
    let xView;
    const ctm = svg.getScreenCTM();
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = t.clientX; sp.y = t.clientY;
      xView = sp.matrixTransform(ctm.inverse()).x;   // letterbox-aware
    } else {
      const rect = svg.getBoundingClientRect();
      xView = ((t.clientX - rect.left) / rect.width) * W;
    }
    let bestI = 0, bestD = Infinity;
    for (let i = 0; i < pts.length; i++) {
      const d = Math.abs(xAt(i) - xView);
      if (d < bestD) { bestD = d; bestI = i; }
    }
    return bestI;
  }
  capture.addEventListener("mousemove", (e) => show(eventToIdx(e)));
  capture.addEventListener("mouseleave", hide);
  capture.addEventListener("touchstart", (e) => { show(eventToIdx(e)); e.preventDefault(); }, { passive: false });
  capture.addEventListener("touchmove",  (e) => { show(eventToIdx(e)); e.preventDefault(); }, { passive: false });
  capture.addEventListener("touchend", hide);
  // Surface the most-recent point on initial render so the tooltip isn't empty.
  show(pts.length - 1);
}

function formatYearMonth(ym) {
  const [y, m] = ym.split("-");
  const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${names[Number(m) - 1] || m} ${y}`;
}

// Indian-convention date formatters: DD-MM (compact, for chart ticks +
// inline cohort labels) and DD-MM-YY (full, for tooltips + subtitles).
// Both accept a YYYY-MM-DD string.
function fmtDateDM(d) {
  if (!d || typeof d !== "string") return d;
  const parts = d.split("-");
  if (parts.length !== 3) return d;
  return `${parts[2]}-${parts[1]}`;
}
function fmtDateDMY(d) {
  if (!d || typeof d !== "string") return d;
  const parts = d.split("-");
  if (parts.length !== 3) return d;
  return `${parts[2]}-${parts[1]}-${parts[0].slice(2)}`;
}

// Detect which current-cohort tickers also appeared in any PRIOR
// cohort's locked top 7. For each repeater, return the list of prior
// months it was in — the cohort table renders a 🔁 badge with that
// list as the tooltip. Empty array until we have ≥ 2 cohorts (organic
// activation come July 1).
function findRepeaters(currentCohort, allCohorts) {
  if (!currentCohort || allCohorts.length < 2) return new Map();
  const currentTickers = new Set(currentCohort.top7.map((s) => s.ticker));
  const priorByTicker = new Map();                       // ticker → [month, month, ...]
  for (const c of allCohorts) {
    if (c.month >= currentCohort.month) continue;
    for (const s of c.top7) {
      if (!currentTickers.has(s.ticker)) continue;
      if (!priorByTicker.has(s.ticker)) priorByTicker.set(s.ticker, []);
      priorByTicker.get(s.ticker).push(c.month);
    }
  }
  return priorByTicker;
}

// Previous month summary card — final outcome of the cohort that just
// closed. Renders only when ≥ 2 cohorts exist; today (June 2026) is
// the first cohort so this is silently absent. Comes alive
// automatically once July 1's snapshot lands and June becomes the
// "previous" month.
function renderPrevMonthSummary(cohorts, getClientPicks, niftyOn, currentMonth) {
  if (!cohorts || cohorts.length < 2) return "";
  // The "previous" month is the cohort immediately before the one
  // currently selected. If user scrubs back via the picker, we always
  // show the cohort right before that one.
  const idx = cohorts.findIndex((c) => c.month === currentMonth);
  if (idx <= 0) return "";
  const prev = cohorts[idx - 1];

  const clientPicks = getClientPicks(prev.month);
  const series = cohortSeries(prev, clientPicks, niftyOn);
  const final = series.points[series.points.length - 1];
  if (!final) return "";

  const fmt = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const cls = (v) => v == null ? "text-slate-500" : v >= 0 ? "text-emerald-700" : "text-rose-700";
  const alpha = (final.ourAvg != null && final.niftyRet != null) ? final.ourAvg - final.niftyRet : null;
  const daysHeld = series.points.length - 1;

  // Stock-by-stock row for the prior cohort
  const stockRows = prev.top7.map((s) => {
    const cur = final.perStockOurs[s.ticker];
    const ret = cur?.ret;
    const retCls = ret == null ? "text-slate-500" : ret >= 0 ? "text-emerald-700" : "text-rose-700";
    return `
      <div class="flex items-center justify-between gap-2 px-2 py-1 rounded hover:bg-slate-50 text-[11px]">
        <span class="font-semibold text-slate-700 truncate" title="${escapeHtml(s.name || s.ticker)}">${escapeHtml(s.name || s.ticker)}</span>
        <span class="tabular-nums font-bold ${retCls}">${ret == null ? "—" : (ret >= 0 ? "+" : "") + ret.toFixed(2) + "%"}</span>
      </div>
    `;
  }).join("");

  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 sm:p-5 mb-4">
      <div class="flex flex-wrap items-start justify-between gap-2 mb-3">
        <div>
          <div class="text-[10px] font-bold uppercase tracking-[0.2em] text-slate-400">Previous month · closed</div>
          <h2 class="font-display font-bold text-slate-900 text-base mt-0.5">${formatYearMonth(prev.month)} cohort · final outcome</h2>
          <div class="text-[11px] text-slate-500 mt-0.5">Entry ${prev.entryDate} → close of ${final.date} · ${daysHeld} trading day${daysHeld === 1 ? "" : "s"}</div>
        </div>
        <div class="flex flex-wrap items-center gap-2">
          <span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg ring-1 ring-slate-200 bg-slate-50 text-[11px]"><span class="w-2 h-2 rounded-full" style="background:${COHORT_COLOR.ours}"></span><span class="text-slate-500">Ours</span> <span class="font-bold tabular-nums ${cls(final.ourAvg)}">${fmt(final.ourAvg)}</span></span>
          ${series.hasClient ? `<span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg ring-1 ring-slate-200 bg-slate-50 text-[11px]"><span class="w-2 h-2 rounded-full" style="background:${COHORT_COLOR.client}"></span><span class="text-slate-500">Client</span> <span class="font-bold tabular-nums ${cls(final.clientAvg)}">${fmt(final.clientAvg)}</span></span>` : ""}
          ${series.hasNifty ? `<span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg ring-1 ring-slate-200 bg-slate-50 text-[11px]"><span class="w-2 h-2 rounded-full" style="background:${COHORT_COLOR.nifty}"></span><span class="text-slate-500">Nifty</span> <span class="font-bold tabular-nums ${cls(final.niftyRet)}">${fmt(final.niftyRet)}</span></span>` : ""}
          ${alpha != null ? `<span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg ring-1 ring-slate-200 bg-slate-50 text-[11px]"><span class="text-slate-500">α</span> <span class="font-bold tabular-nums ${cls(alpha)}">${fmt(alpha)}</span></span>` : ""}
        </div>
      </div>
      <div class="rounded-xl bg-slate-50 ring-1 ring-slate-100 p-3 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-3 gap-y-0.5">
        ${stockRows}
      </div>
    </div>
  `;
}

function renderHistoryEmpty(reason) {
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-10 sm:p-12 text-center">
      <div class="inline-flex items-center justify-center w-14 h-14 rounded-2xl bg-gradient-to-br from-indigo-400 to-purple-500 text-white text-2xl mb-4 shadow-lg">📈</div>
      <h2 class="text-xl font-bold font-display text-slate-900 mb-2">History is building up</h2>
      <p class="text-sm text-slate-600 max-w-lg mx-auto leading-relaxed">${escapeHtml(reason || "We need a few daily snapshots before this view becomes useful.")}<br><br>Each daily refresh appends a snapshot to <code class="text-[11px] bg-slate-100 px-1.5 py-0.5 rounded">public/data/snapshots/</code>. As STRONG BUY picks accumulate, this tab will show realized return per pick + a chart with rating markers overlaid on the price line.</p>
    </div>
  `;
}

// ============================================================
// ACTIVE BASKET — daily-rebalancing backtest
// ============================================================
// Walks the snapshot trail and simulates an equal-weight portfolio that
// rebalances every snapshot day to hold exactly today's STRONG BUY set:
//   - Day 1: open positions in every STRONG BUY at that day's close.
//   - Day N: exit anything no longer STRONG BUY at today's close, then
//            rebalance the entire remaining basket to NAV ÷ N — trim
//            winners that drifted overweight, top up losers that
//            drifted underweight, fund new entrants at the same weight.
//
// Buys and sells transact at the same close used for mark-to-market on
// that day, so opening a position generates no synthetic gain/loss; the
// equity curve reflects pure model behaviour.
//
// Rebalance runs every day (not just basket-change days) to honour the
// equal-weight invariant promised by the "Daily Rebalance" label —
// otherwise winners drift overweight and losers shrink between events.
// Top-ups update the position's entry price to a weighted average of
// the lots bought, so realized-return stats (hit rate, avg winner /
// loser) reflect the cash actually allocated rather than just the
// first lot's price. Trims leave the basis unchanged (FIFO / weighted-
// average equivalence when selling, not buying). Internal weight
// adjustments stay invisible in the trade log; only real basket adds /
// drops surface. No transaction costs in v1 — "model truth" without
// friction.
const ACTIVE_INITIAL_CAPITAL = 100000;

function simulateActiveBasket(snapshots, anchorDate, simP = simPrefs) {
  if (!snapshots?.length) return null;
  // Anchor at the client upload date when supplied (the founder's
  // "everything starts from upload" rule). Filter to snapshots ON or
  // AFTER the anchor so the first BUY event is at the upload close.
  let sorted = [...snapshots].sort((a, b) => a.date.localeCompare(b.date));
  if (anchorDate) sorted = sorted.filter((s) => s.date >= anchorDate);
  if (!sorted.length) return null;

  // Real-world inputs: total pot, an idle cash buffer held back from
  // deployment, and a per-side transaction charge applied to every BUY
  // (position open) and SELL (position exit). Internal daily equal-weight
  // trims / top-ups are model housekeeping — charging them would explode
  // costs unrealistically — so charges land only on real basket changes,
  // the round-trip the client actually pays per holding.
  const capital = simP?.capital ?? ACTIVE_INITIAL_CAPITAL;
  const bufferAmt = capital * Math.max(0, simP?.bufferPct ?? 0) / 100;
  const chg = simP === ZERO_CHARGER ? ZERO_CHARGER : makeCharger(simP);
  let cash = capital;
  let totalCharges = 0;
  const holdings = new Map();   // ticker → { units, entryDate, entryPrice, name, sector }
  const trades = [];            // BUY (new entrant) / SELL (basket exit), time-ordered
  const equity = [];            // { date, value, holdingCount, cash }

  for (const snap of sorted) {
    const date = snap.date;
    const closeByTicker = new Map();
    for (const s of snap.stocks) if (typeof s.close === "number") closeByTicker.set(s.ticker, s.close);

    // Today's target = STRONG BUY rated stocks with a close. (Out-of-
    // coverage injects have rating: null so they're naturally excluded.)
    const target = snap.stocks
      .filter((s) => s.rating === "STRONG BUY" && typeof s.close === "number" && !s.hardFailed && s.ticker);
    const targetByTicker = new Map(target.map((s) => [s.ticker, s]));

    // 1. Exit positions not in today's target — book SELL with realized
    //    return based on weighted-average entry price.
    for (const [ticker, pos] of [...holdings.entries()]) {
      if (targetByTicker.has(ticker)) continue;
      const exitPrice = closeByTicker.get(ticker);
      if (typeof exitPrice !== "number") continue;   // carry to next day if missing close
      const gross = pos.units * exitPrice;
      const fee = chg.sell(gross);
      cash += gross - fee;
      totalCharges += fee;
      const ret = (exitPrice / pos.entryPrice - 1) * 100;
      trades.push({
        action: "SELL",
        ticker, name: pos.name, sector: pos.sector,
        date, price: exitPrice,
        entryDate: pos.entryDate, entryPrice: pos.entryPrice,
        days: daysBetween(pos.entryDate, date),
        ret,
      });
      holdings.delete(ticker);
    }

    // 2. Daily rebalance to equal weight across continuing + new
    //    positions. Runs unconditionally so prices drifting between
    //    snapshots doesn't quietly skew weights between basket events.
    const newEntries = target.filter((s) => !holdings.has(s.ticker));
    if (target.length > 0) {
      let nav = cash;
      for (const [ticker, pos] of holdings.entries()) {
        const px = closeByTicker.get(ticker) ?? pos.entryPrice;
        nav += pos.units * px;
      }
      // Hold the buffer back as idle cash; deploy the rest equal-weight.
      const deployable = Math.max(0, nav - bufferAmt);
      const targetValuePerStock = deployable / target.length;

      // Adjust continuing positions to target value. On top-ups, update
      // entry price to weighted-average cost basis so closed-trade
      // returns reflect the cash actually allocated. On trims, basis
      // stays the same — the realized portion of any gain/loss is
      // captured in the equity curve via the cash flow.
      for (const [ticker, pos] of holdings.entries()) {
        const px = closeByTicker.get(ticker);
        if (typeof px !== "number") continue;
        const currentValue = pos.units * px;
        const valueDelta = targetValuePerStock - currentValue;   // +ve = buy more, -ve = trim
        if (Math.abs(valueDelta) < 0.01) continue;
        const unitDelta = valueDelta / px;
        if (unitDelta > 0) {
          // Top-up: weighted-average cost basis update.
          const newUnits = pos.units + unitDelta;
          pos.entryPrice = (pos.units * pos.entryPrice + unitDelta * px) / newUnits;
          pos.units = newUnits;
        } else {
          // Trim: units down, basis unchanged.
          pos.units += unitDelta;
        }
        cash -= valueDelta;
      }

      // Open new entrants at the target weight. Cap by remaining cash
      // so floating-point dust can't push cash negative.
      for (const e of newEntries) {
        // Cap so buyValue + its charge can't overdraw cash.
        const buyValue = Math.min(targetValuePerStock, Math.max(0, cash) / (1 + buyRate(chg.prefs || {})));
        if (buyValue < 0.01) continue;
        const fee = chg.buy(buyValue);
        const units = buyValue / e.close;
        holdings.set(e.ticker, {
          units, entryDate: date, entryPrice: e.close,
          name: e.name || e.ticker, sector: e.sector || null,
        });
        trades.push({
          action: "BUY",
          ticker: e.ticker, name: e.name || e.ticker, sector: e.sector || null,
          date, price: e.close,
        });
        cash -= buyValue + fee;
        totalCharges += fee;
      }
    }

    // 3. Mark-to-market today's portfolio value.
    let value = cash;
    for (const [ticker, pos] of holdings.entries()) {
      const px = closeByTicker.get(ticker) ?? pos.entryPrice;
      value += pos.units * px;
    }
    equity.push({ date, value, holdingCount: holdings.size, cash });
  }

  return { equity, trades, holdings, startCapital: capital, totalCharges };
}

// Derived stats from a simulation result + Nifty 50 closes for alpha.
function computeActiveStats(sim, niftyOn) {
  if (!sim || !sim.equity.length) return null;
  const start = sim.equity[0];
  const end = sim.equity[sim.equity.length - 1];
  const totalRet = (end.value / start.value - 1) * 100;

  // Max drawdown — peak-to-trough on the equity curve.
  let peak = start.value, maxDD = 0;
  for (const e of sim.equity) {
    if (e.value > peak) peak = e.value;
    const dd = (e.value / peak - 1) * 100;
    if (dd < maxDD) maxDD = dd;
  }

  // Closed-trade stats: every SELL has a realized return.
  const sells = sim.trades.filter((t) => t.action === "SELL");
  const winners = sells.filter((t) => t.ret > 0);
  const hitRate = sells.length ? (winners.length / sells.length) * 100 : null;
  const avgHoldDays = sells.length ? sells.reduce((a, t) => a + t.days, 0) / sells.length : null;
  const avgWin = winners.length ? winners.reduce((a, t) => a + t.ret, 0) / winners.length : null;
  const losers = sells.filter((t) => t.ret <= 0);
  const avgLoss = losers.length ? losers.reduce((a, t) => a + t.ret, 0) / losers.length : null;

  // Nifty over the same window — drives the alpha tile.
  const nStart = niftyOn ? niftyOn(start.date) : null;
  const nEnd = niftyOn ? niftyOn(end.date) : null;
  const niftyRet = (nStart && nEnd) ? (nEnd / nStart - 1) * 100 : null;
  const alpha = niftyRet != null ? totalRet - niftyRet : null;

  const buys = sim.trades.filter((t) => t.action === "BUY");
  return {
    totalRet, maxDD, hitRate, avgHoldDays, avgWin, avgLoss,
    niftyRet, alpha,
    days: sim.equity.length,
    startDate: start.date, endDate: end.date,
    tradeCount: sim.trades.length,
    buyCount: buys.length,
    sellCount: sells.length,
    liveHoldings: sim.holdings.size,
    finalValue: end.value,
  };
}

// Active tab renderer. Three cadences (Daily / Weekly / Monthly) share
// a common shell:
//   - sub-pill toggle in the header
//   - hero card with total return + alpha vs Nifty
//   - cumulative return chart (active basket vs Nifty)
//   - overall hit summary (founder ask — daily basket churns so per-pick
//     hits across the full history are the headline number)
//   - per-pick accuracy rows (every entry = one tracked pick)
//
// All three anchor at lkp.generated_at (client upload date). The Daily
// cell reuses the equal-weight NAV-tracking simulator (simulateActiveBasket).
// Weekly and Monthly cells re-use buildCohort with weekly mode + a new
// 30-day chain extension so segments re-lock at the named cadence and
// hold frozen during each segment.
async function renderActive() {
  const host = $("#active-content");
  if (!host) return;
  host.innerHTML = `<div class="bg-white rounded-2xl ring-1 ring-slate-200 p-10 text-center text-slate-500 text-sm">Loading active strategy…</div>`;

  // Outer try/catch around the WHOLE render — without it, an exception
  // in any of the post-await helpers (buildActiveView, render*, etc.)
  // is swallowed silently by the async function's rejected promise and
  // the user is left staring at "Loading…" forever. Show the error +
  // log to console so we can see what's wrong.
  try {
    try {
      await ensureHistoryCache();
      // ADTV lives only in technicals.json (snapshots carry techVals, which
      // omits it). Fetch once per session; the Trade Plan needs it to flag
      // names too thin to place a market order in.
      if (!state.cache.planAdtv) {
        try {
          const tj = await fetch("data/technicals.json").then((r) => (r.ok ? r.json() : null));
          const m = {};
          for (const c of (tj?.companies || [])) if (c.ticker && c.adtv_20d_cr != null) m[c.ticker] = c.adtv_20d_cr;
          state.cache.planAdtv = m;
        } catch { state.cache.planAdtv = {}; }
      }
      // Strategy definitions + their precomputed backtest, and the pre-open
      // market context. All optional: a missing file degrades one panel
      // rather than taking the whole tab down with it.
      await strat5.loadStrategies();
      // A #1 chosen in this browser but not yet committed still has to win
      // here, or the crown and the Trade Plan would disagree after a reload.
      try {
        const override = localStorage.getItem("vn-primary-override-v1");
        if (override && strat5.strategyById(override)) strat5.setPrimaryId(override);
      } catch {}
      if (state.cache.strategyBacktest === undefined) {
        try { state.cache.strategyBacktest = await fetch("data/strategy-backtest.json").then((r) => (r.ok ? r.json() : null)); }
        catch { state.cache.strategyBacktest = null; }
      }
      if (state.cache.marketContext === undefined) {
        try { state.cache.marketContext = await fetch("data/market-context.json").then((r) => (r.ok ? r.json() : null)); }
        catch { state.cache.marketContext = null; }
      }
    } catch (e) {
      host.innerHTML = renderHistoryEmpty(e.message);
      return;
    }
    const { snapshots, benchmark, lkp } = state.cache.history;
    if (!snapshots.length) {
      host.innerHTML = renderHistoryEmpty("No snapshots loaded.");
      return;
    }

    // Report-month selection — June stays reachable after July is uploaded.
    // Default = latest month; each month anchors at its own report date.
    const lkpForMonths = lkpOverride() || lkp;
    const manualMonths = availableManualMonths(lkpForMonths, snapshots);
    const selectedMonth = (state.manualMonth && manualMonths.includes(state.manualMonth))
      ? state.manualMonth
      : (manualMonths[manualMonths.length - 1] || null);
    const anchorDate = manualMonthAnchor(lkpForMonths, selectedMonth, snapshots);
    const monthSelectorHtml = renderManualMonthSelector(manualMonths, selectedMonth, lkpForMonths, snapshots);
    // A basket is tracked for its 4-month life. Once past the hold window,
    // stop the clock at month 4 (both baskets) — any still-open pick is
    // closed at that day's market price, matching the client's hold rule.
    const HOLD_MONTHS = 4;
    const fullToday = snapshots[snapshots.length - 1].date;
    const capDate = anchorDate ? addMonthsStr(anchorDate, HOLD_MONTHS) : fullToday;
    const capped = capDate < fullToday;
    const todayDate = capped ? capDate : fullToday;
    const viewSnaps = capped ? snapshots.filter((s) => s.date <= todayDate) : snapshots;
    const holdNote = capped
      ? `<div class="bg-slate-50 rounded-2xl ring-1 ring-slate-200 px-3 py-2 text-[11px] text-slate-600 flex items-center gap-2"><span>🔒</span><span>This basket completed its <strong>4-month hold</strong> — tracking closed on ${fmtDateDMY(todayDate)}. Open positions were marked at that day's close.</span></div>`
      : "";

    const niftyClosesByDate = benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || null;
    const niftyDatesSorted = niftyClosesByDate ? Object.keys(niftyClosesByDate).sort() : null;
    function niftyOn(date) {
      if (!niftyClosesByDate) return null;
      if (niftyClosesByDate[date] != null) return niftyClosesByDate[date];
      let last = null;
      for (const d of niftyDatesSorted) { if (d <= date) last = niftyClosesByDate[d]; else break; }
      return last;
    }
    // Midcap 150 — secondary reference; the Rs 12,500 Cr ceiling straddles it.
    const nifty500ClosesByDate = benchmark?.indices?.["NIFTYMIDCAP150.NS"]?.closes || null;
    const nifty500DatesSorted = nifty500ClosesByDate ? Object.keys(nifty500ClosesByDate).sort() : null;
    function nifty500On(date) {
      if (!nifty500ClosesByDate) return null;
      if (nifty500ClosesByDate[date] != null) return nifty500ClosesByDate[date];
      let last = null;
      for (const d of nifty500DatesSorted) { if (d <= date) last = nifty500ClosesByDate[d]; else break; }
      return last;
    }

    // Strategy mode (top-level: active vs passive) decides the buildView
    // contract. Active mode passes the user-chosen cadence; Passive mode
    // passes "passive" so the view builder produces a single-segment chain
    // anchored at upload (no re-locking — AI basket fixed forever).
    const mode = SHOW_ROTATION_STRATEGIES ? state.strategyMode : "passive";
    const cadence = mode === "passive" ? "passive" : state.activeCadence;
    // Manual basket — fixed at upload, same across all 3 cadences. Resolved
    // via the same picksByMonth / picks fallback the History tab uses.
    const lkpResolved = lkpOverride() || lkp;
    const anchorMonth = anchorDate?.slice(0, 7) || null;
    const mostRecentMonth = snapshots[snapshots.length - 1].date.slice(0, 7);
    const manualPicks = lkpResolved
      ? lkpPicksForMonth(lkpResolved, anchorMonth, mostRecentMonth) || lkpResolved.picks || []
      : [];

    // Quotes are fetched in the BACKGROUND, not awaited. Blocking the first
    // paint on them left the tab reading "Loading active strategy…" for
    // seven seconds whenever the API was slow -- trading a stale number for
    // a blank screen, which is the worse of the two. Draw with what we have,
    // then redraw when the quotes land. The TTL check inside stops the
    // second render from starting another fetch.
    try {
      const anchorSnap = snapshots.find((s) => s.date === anchorDate) || snapshots[snapshots.length - 1];
      const want = (strat5.strategyById(strat5.primaryId())?.basketSize || 7) + 3;
      const cohort = (anchorSnap?.stocks || [])
        .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed)
        .sort((a, b) => b.composite - a.composite).slice(0, want).map((s) => s.ticker);
      refreshLiveQuotes([...cohort, ...Object.keys(state.cache.history.livePrices || {})])
        .then((changed) => { if (changed && state.activeTab === "active") renderActive(); })
        .catch(() => {});
    } catch (e) { console.warn("live quote refresh skipped:", e?.message); }

    const view = buildActiveView(viewSnaps, anchorDate, todayDate, cadence, niftyOn, manualPicks, nifty500On);

    // Alerts section (lives inside the Strategy tab, not a separate tab).
    if (!customTechByTicker) {
      try { const tj = await fetch("data/technicals.json").then((r) => r.json()); customTechByTicker = new Map((tj.companies || []).map((c) => [c.ticker, c])); } catch { customTechByTicker = new Map(); }
    }
    let alertsHtml = "";
    try {
      const strategyHits = view ? collectStrategyHits(view, todayDate) : { todayHits: [] };
      const alerts = evaluateAlerts(snapshots, customTechByTicker, strategyHits, alertPrefs);
      alertsHtml = renderAlertsSection(alerts, alertPrefs);
    } catch (e) { console.error("alerts section failed:", e); }

    host.innerHTML = renderActiveShell(view, cadence, anchorDate, todayDate, mode, alertsHtml, monthSelectorHtml, holdNote);
    wireStrategyModeToggle();
    wireManualReturnToggle();
    wireStrategySubNav();
    wireManualMonthPills();
    // Client-basket upload (Excel / CSV) — surfaced here on the Strategy tab.
    $("#lkp-upload-btn")?.addEventListener("click", () => $("#lkp-file-input")?.click());
    $("#lkp-file-input")?.addEventListener("change", (e) => { const f = e.target.files?.[0]; if (f) handleLkpExcelUpload(f); e.target.value = ""; });
    $("#lkp-reset-upload")?.addEventListener("click", () => { clearLkpOverride(); renderActive(); });
    wireActiveCadenceToggle();
    wireStrategySegmentPills();
    wireSectorTimingToggle("#active-content", renderActive);
    wireStrategyBalanceToggle();
    wireStrategyAlertsDropdown();
    wireAlertsInputs("#active-content", renderActive);
    // Cohort-style row clicks open the same drill modal everywhere else uses.
    $$("#active-content [data-cohort-row]").forEach((el) => el.addEventListener("click", () => {
      const ticker = el.dataset.ticker;
      const side = el.dataset.cohortSide || "ai";
      const segAnchor = el.dataset.segAnchor || null;
      if (!ticker) return;
      const pick = buildCohortClickPick(ticker, side, segAnchor);
      if (pick) openHistoryDrill(pick);
    }));
    // "+ N more" — expand / collapse the per-pick column preview in place.
    $$("#active-content [data-pick-toggle]").forEach((btn) => btn.addEventListener("click", () => {
      const side = btn.dataset.pickToggle;
      const list = $(`#active-content [data-pick-list="${side}"]`);
      if (!list) return;
      const extra = list.querySelectorAll(".pick-extra-row");
      const expand = btn.dataset.expanded !== "1";
      extra.forEach((el) => { el.style.display = expand ? "" : "none"; });
      btn.dataset.expanded = expand ? "1" : "0";
      btn.textContent = expand ? "Show less" : `+ ${extra.length} more`;
    }));
    // Capital & charges panel — any change recomputes the whole strategy.
    $$("#active-content [data-sim-field]").forEach((inp) => inp.addEventListener("change", () => {
      const key = inp.dataset.simField;
      const num = parseFloat(inp.value);
      if (!Number.isFinite(num) || num < 0) { inp.value = simPrefs[key]; return; }
      simPrefs = { ...simPrefs, [key]: num };
      saveSimPrefs(simPrefs);
      renderActive();
    }));
    const simReset = $("#sim-reset");
    if (simReset) simReset.addEventListener("click", () => {
      simPrefs = { ...SIM_DEFAULTS };
      saveSimPrefs(simPrefs);
      renderActive();
    });
    wireStrategyChartMode();
    wireStrategyStockBasket();
    // Trade Plan: switch which strategy's plan you are looking at.
    $$("#active-content [data-plan-strategy]").forEach((b) => b.addEventListener("click", () => {
      state.planStrategyId = b.dataset.planStrategy;
      try { localStorage.setItem("vn-plan-strategy-v1", state.planStrategyId); } catch {}
      renderActive();
    }));
    // Compare: view switch + "make this my #1".
    $$("#active-content [data-compare-source]").forEach((b) => b.addEventListener("click", () => {
      state.compareSource = b.dataset.compareSource;
      try { localStorage.setItem("vn-compare-source-v1", state.compareSource); } catch {}
      renderActive();
    }));
    $$("#active-content [data-compare-view]").forEach((b) => b.addEventListener("click", () => {
      state.compareView = b.dataset.compareView;
      try { localStorage.setItem("vn-compare-view-v1", state.compareView); } catch {}
      renderActive();
    }));
    $$("#active-content [data-set-primary]").forEach((b) => b.addEventListener("click", () => {
      const id = b.dataset.setPrimary;
      strat5.setPrimaryId(id);
      state.planStrategyId = id;
      try {
        localStorage.setItem("vn-plan-strategy-v1", id);
        localStorage.setItem("vn-primary-override-v1", id);
      } catch {}
      openPrimaryCommitModal(id);
      renderActive();
    }));
    // Chart hover crosshair + tooltip — basket mode only (the per-stock
    // "Stocks" view draws static lines with its own legend).
    if (state.strategySubTab === "overview" && (state.strategyChartMode || "basket") !== "stocks") setupActiveChartHover(view);
  } catch (e) {
    console.error("renderActive failed:", e);
    host.innerHTML = `
      <div class="bg-white rounded-2xl shadow-sm ring-1 ring-rose-200 p-6">
        <div class="flex items-start gap-3">
          <div class="text-rose-500 text-2xl">⚠</div>
          <div class="flex-1">
            <h2 class="font-bold text-slate-900 text-base">Strategy tab failed to render</h2>
            <p class="text-sm text-slate-600 mt-1">${escapeHtml(e?.message || String(e))}</p>
            <pre class="text-[10px] text-slate-500 mt-2 whitespace-pre-wrap overflow-x-auto bg-slate-50 rounded p-2 max-h-48">${escapeHtml(e?.stack || "")}</pre>
            <p class="text-xs text-slate-500 mt-2">Check browser DevTools console for full trace.</p>
          </div>
        </div>
      </div>`;
  }
}

// Build everything the Active shell needs for a given cadence. Returns
// { kind, sim?, segments?, picks, hitSummary, equityCurve, niftyCurve,
//   manualCurve, manualPicks, manualSummary, periodLabel, ... }.
function buildActiveView(snapshots, anchorDate, todayDate, cadence, niftyOn, manualPicks, nifty500On = () => null) {
  if (cadence === "daily") {
    const sim = simulateActiveBasket(snapshots, anchorDate, simPrefs);
    if (!sim || !sim.equity.length) return null;
    // Frictionless twin — same allocation, zero charges — so we can show
    // how much the charges actually cost (gross vs net).
    const simGross = simulateActiveBasket(snapshots, anchorDate, ZERO_CHARGER);
    const picks = buildActiveDailyPicks(sim, snapshots, todayDate);
    const hitSummary = computeOverallHitSummary(picks);
    const start = sim.equity[0];
    const end = sim.equity[sim.equity.length - 1];
    const equityCurve = sim.equity.map((e) => ({ date: e.date, retPct: (e.value / sim.startCapital - 1) * 100 }));
    const finalReturn = equityCurve[equityCurve.length - 1].retPct;
    const grossEnd = simGross.equity[simGross.equity.length - 1];
    const grossFinalReturn = (grossEnd.value / simGross.startCapital - 1) * 100;
    const nStart = niftyOn(start.date), nEnd = niftyOn(end.date);
    const niftyRet = (nStart && nEnd) ? (nEnd / nStart - 1) * 100 : null;
    const dates = equityCurve.map((e) => e.date);
    const niftyCurve = buildNiftyCurve(dates, niftyOn);
    const { manualRows, manualSummary, manualCurve, manualBooked } = buildManualBundle(manualPicks, snapshots, anchorDate, todayDate, dates);
    // Per-day segment chain so the basket roster panel can show each
    // day's top 7 (founder ask — same affordance the History tab used
    // to give for weekly baskets, just at 1-day granularity here).
    const segments = buildActiveSegmentChain(snapshots, anchorDate, 1);
    const nifty500Curve = buildNiftyCurve(dates, nifty500On);
    const nifty500Ret = nifty500Curve.length ? (nifty500Curve[nifty500Curve.length - 1].retPct ?? null) : null;
    const { aiStockCurves, manualStockCurves } = buildBasketStockCurves(segments, manualPicks, snapshots, anchorDate, dates, picks, manualRows);
    return {
      kind: "daily",
      sim, picks, hitSummary, segments,
      equityCurve, niftyCurve, manualCurve,
      nifty500Curve, nifty500Ret, aiStockCurves, manualStockCurves,
      manualPicks: manualRows, manualSummary, manualBooked,
      periodLabel: `Daily rebalance from ${fmtDateDMY(anchorDate)}`,
      finalReturn, niftyRet,
      alpha: niftyRet != null ? finalReturn - niftyRet : null,
      manualFinalReturn: manualCurve.length ? (manualCurve[manualCurve.length - 1].retPct ?? null) : null,
      startDate: start.date, endDate: end.date,
      finalValue: end.value, startCapital: sim.startCapital,
      grossFinalReturn, totalCharges: sim.totalCharges,
      tradeCount: sim.trades.length,
      liveHoldings: sim.holdings.size,
    };
  }

  // Weekly / Monthly / Passive — all segmented chains. Weekly + Monthly
  // re-pick top 7 every 7 / 30 days from anchor. Passive uses a single
  // segment that runs upload → today (no re-locking — basket frozen).
  // Returns the same { segments, equityCurve, picks, ... } contract so
  // the chart / accuracy / per-pick rows render uniformly.
  const periodDays = cadence === "weekly" ? 7 : cadence === "monthly" ? 30 : 99999;
  const segments = buildActiveSegmentChain(snapshots, anchorDate, periodDays);
  if (!segments.length) return null;
  const picks = buildActiveSegmentedPicks(segments, snapshots, todayDate);
  const hitSummary = computeOverallHitSummary(picks);
  const chg = makeCharger(simPrefs);
  // Same booked/held convention as the manual basket, so AI and Manual are a
  // like-to-like comparison: booked freezes each name at its first target / SL.
  // Trading calendar straight off the benchmark: if the index has no close
  // for a date, the market was shut and nothing may be plotted there.
  const tradingDays = new Set(Object.keys(state.cache.history?.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || {}));
  const curveOpts = { booked: state.manualReturnMode !== "held", todayDate, livePrices: state.cache.history?.livePrices || {}, tradingDays };
  const equityCurve = buildSegmentedEquityCurve(segments, snapshots, anchorDate, chg, { ...curveOpts, capital: simPrefs.capital });
  const grossCurve = buildSegmentedEquityCurve(segments, snapshots, anchorDate, ZERO_CHARGER, curveOpts);
  const dates = equityCurve.map((e) => e.date);
  const niftyCurve = buildNiftyCurve(dates, niftyOn);
  const { manualRows, manualSummary, manualCurve, manualBooked } = buildManualBundle(manualPicks, snapshots, anchorDate, todayDate, dates);
  const finalReturn = equityCurve.length ? equityCurve[equityCurve.length - 1].retPct : 0;
  const grossFinalReturn = grossCurve.length ? grossCurve[grossCurve.length - 1].retPct : finalReturn;
  const niftyRet = niftyCurve.length ? niftyCurve[niftyCurve.length - 1].retPct : null;
  const nifty500Curve = buildNiftyCurve(dates, nifty500On);
  const nifty500Ret = nifty500Curve.length ? (nifty500Curve[nifty500Curve.length - 1].retPct ?? null) : null;
  const { aiStockCurves, manualStockCurves } = buildBasketStockCurves(segments, manualPicks, snapshots, anchorDate, dates, picks, manualRows);
  const capital = simPrefs.capital ?? ACTIVE_INITIAL_CAPITAL;
  const finalValue = capital * (1 + finalReturn / 100);
  // Charges in ₹ ≈ the gap the costs opened between gross and net pots.
  const totalCharges = capital * (grossFinalReturn - finalReturn) / 100;
  const cadenceLabel = cadence === "weekly" ? "Weekly re-lock" : cadence === "monthly" ? "Monthly re-lock" : "Passive (basket frozen)";
  return {
    kind: cadence,
    entryCostPct: equityCurve.entryCostPct ?? 0,
    liveMark: !!equityCurve.liveMark,
    tradedDays: Math.max(0, equityCurve.length - 1),
    segments, picks, hitSummary,
    equityCurve, niftyCurve, manualCurve,
    nifty500Curve, nifty500Ret, aiStockCurves, manualStockCurves,
    manualPicks: manualRows, manualSummary, manualBooked,
    periodLabel: `${cadenceLabel} from ${fmtDateDMY(anchorDate)}${cadence === "passive" ? "" : ` · ${segments.length} segment${segments.length === 1 ? "" : "s"}`}`,
    finalReturn, niftyRet,
    alpha: niftyRet != null ? finalReturn - niftyRet : null,
    manualFinalReturn: manualCurve.length ? (manualCurve[manualCurve.length - 1].retPct ?? null) : null,
    startDate: equityCurve.length ? equityCurve[0].date : anchorDate,
    endDate: equityCurve.length ? equityCurve[equityCurve.length - 1].date : todayDate,
    finalValue, startCapital: capital, grossFinalReturn, totalCharges,
  };
}

// Manual basket cumulative return per day. Anchored at the snapshot
// close on the upload date (in-universe picks only — out-of-coverage
// picks are skipped). Forward-fills missing closes so the basket size
// stays at its locked count even on illiquid days.
//
// bookingMap (optional): ticker -> { booked, bookDate, bookPrice }. When
// supplied (Booked mode), a pick's price freezes at bookPrice from its
// bookDate onward — modelling the real exit at the first target (T1) / SL so
// a winner that kept running past target doesn't keep inflating the basket.
function buildActiveManualCurve(snapshots, anchorDate, manualPicks, dates, bookingMap) {
  if (!manualPicks?.length || !dates?.length) return [];
  const anchorSnap = snapshots.find((s) => s.date >= anchorDate);
  if (!anchorSnap) return [];
  const inUni = manualPicks.filter((p) => p && p.in_universe && p.ticker);
  if (!inUni.length) return [];
  const entryCloses = {};
  for (const p of inUni) {
    const s = anchorSnap.stocks.find((x) => x.ticker === p.ticker);
    if (s && typeof s.close === "number") entryCloses[p.ticker] = s.close;
  }
  if (!Object.keys(entryCloses).length) return [];
  const lastClose = { ...entryCloses };
  const snapByDate = new Map(snapshots.map((s) => [s.date, s]));
  return dates.map((date) => {
    const snap = snapByDate.get(date);
    if (!snap) return { date, retPct: null };
    let sum = 0, n = 0;
    for (const ticker of Object.keys(entryCloses)) {
      const bk = bookingMap?.get(ticker);
      const frozen = bk && bk.booked && date >= bk.bookDate;
      let close;
      if (frozen) {
        close = bk.bookPrice;               // sold here — price locked
      } else {
        const s = snap.stocks.find((x) => x.ticker === ticker);
        close = (s && typeof s.close === "number") ? s.close : lastClose[ticker];
      }
      if (typeof close === "number") {
        if (!frozen) lastClose[ticker] = close;
        sum += close / entryCloses[ticker];
        n++;
      }
    }
    if (n === 0) return { date, retPct: null };
    return { date, retPct: (sum / n - 1) * 100 };
  });
}

// Manual basket per-pick accuracy. Each in-universe pick = a row with
// the client's TGT1 (first target — the exit / freeze level) + SL, not the
// AI's uniform bands. Out-of-coverage picks appear as Not Covered rows. Each
// covered row also carries its booking (frozen at Target 1 / SL when hit),
// its booked return and its mark-to-market "if held" return so the UI can
// flip between the two without recomputing.
function buildActiveManualData(manualPicks, snapshots, anchorDate, todayDate) {
  const result = { manualPicks: [], manualSummary: null };
  if (!manualPicks?.length) return result;
  const livePrices = state.cache.history?.livePrices || {};
  const anchorSnap = snapshots.find((s) => s.date >= anchorDate);
  const lastSnap = snapshots[snapshots.length - 1];
  const rows = [];
  for (const p of manualPicks) {
    if (!p.in_universe || !p.ticker) {
      rows.push({
        ticker: p.ticker || null,
        name: p.selection || p.ticker || "—",
        notCovered: true,
        outReason: p.out_reason || "Below market-cap floor / not matched",
      });
      continue;
    }
    const entryFromSnap = anchorSnap?.stocks?.find((x) => x.ticker === p.ticker)?.close;
    const entryPrice = entryFromSnap ?? p.entry ?? null;
    const entryDate = anchorSnap?.date || anchorDate;
    const target = p.tgt1 ?? null;   // book/close at Target 1 — the desk exits at the first target (T2 is "if held" territory)
    const sl = p.sl ?? null;
    if (entryPrice == null || target == null || sl == null) {
      rows.push({
        ticker: p.ticker, name: p.selection || p.ticker,
        entryDate, entryPrice, target, sl,
        targetPct: target != null && entryPrice ? (target / entryPrice - 1) * 100 : null,
        slPct: sl != null && entryPrice ? (sl / entryPrice - 1) * 100 : null,
        status: "OPEN", currentClose: null,
      });
      continue;
    }
    const live = livePrices[p.ticker] || livePrices[(p.selection || "").toUpperCase()] || null;
    const status = computeManualMilestones(p.ticker, entryDate, entryPrice, target, sl, snapshots, todayDate, live);
    const peak = computePeakStats(p.ticker, entryDate, entryPrice, snapshots, todayDate);
    // Booking = the real exit at the first target (T1) or SL. Freeze at the
    // LEVEL, not the overshooting close — a pick that ran to +22% but had
    // T1 = +12% books at +12%. Held return marks to market, so the UI can
    // show both ("if held" is where it went after we'd already closed).
    const todayCloseVal = (live && typeof live.current === "number")
      ? live.current
      : (lastSnap?.stocks?.find((x) => x.ticker === p.ticker)?.close ?? entryPrice);
    const heldRet = entryPrice ? (todayCloseVal / entryPrice - 1) * 100 : null;
    const isBooked = status.status === "TARGET_HIT" || status.status === "SL_HIT";
    const bookPrice = status.status === "TARGET_HIT" ? target : status.status === "SL_HIT" ? sl : null;
    const booking = isBooked
      ? { booked: true, reason: status.status === "TARGET_HIT" ? "TARGET" : "SL", bookDate: status.hitDate, bookPrice, daysToBook: status.daysToHit }
      : { booked: false };
    const bookedRet = (isBooked && entryPrice) ? (bookPrice / entryPrice - 1) * 100 : heldRet;
    rows.push({
      ticker: p.ticker,
      name: p.selection || p.ticker,
      entryDate, entryPrice, target, sl,
      targetPct: (target / entryPrice - 1) * 100,
      slPct: (sl / entryPrice - 1) * 100,
      ...status,
      peak,
      todayClose: todayCloseVal, heldRet, bookedRet, booking,
    });
  }
  const enriched = enrichAndSortPicks(rows.filter((r) => !r.notCovered));
  const trackable = enriched.filter((r) => r.entryPrice != null && r.target != null && r.sl != null);
  const summary = computeOverallHitSummary(trackable);
  // Re-attach not-covered rows at the bottom for display.
  const notCovered = rows.filter((r) => r.notCovered);
  result.manualPicks = [...enriched, ...notCovered];
  result.manualSummary = summary;
  return result;
}

// Resolve the manual basket's curve + rows + summary honouring the current
// return mode. Booked (default) freezes each pick at its first target / SL
// the day it hits; Held marks to market. Shared by the Active, Passive and
// Custom views so every manual figure agrees. Returns manualBooked so the
// render layer knows which convention produced the numbers.
function buildManualBundle(manualPicks, snapshots, anchorDate, todayDate, dates) {
  const manualBooked = state.manualReturnMode !== "held";
  const { manualPicks: manualRows, manualSummary } = buildActiveManualData(manualPicks, snapshots, anchorDate, todayDate);
  const bookingMap = manualBooked
    ? new Map(manualRows.filter((r) => r.booking?.booked && r.ticker).map((r) => [r.ticker, r.booking]))
    : null;
  const manualCurve = buildActiveManualCurve(snapshots, anchorDate, manualPicks, dates, bookingMap);
  return { manualRows, manualSummary, manualCurve, manualBooked };
}

// Re-pick top 7 every periodDays from anchorDate. Each segment locks at
// its start snapshot and tracks closes through the segment window.
// Returns [{ index, label, startDate, endDate, entrySnap, top7, tracking }].
function buildActiveSegmentChain(snapshots, anchorDate, periodDays) {
  if (!snapshots.length || !anchorDate) return [];
  const today = snapshots[snapshots.length - 1].date;
  const dateToMs = (d) => Date.UTC(Number(d.slice(0, 4)), Number(d.slice(5, 7)) - 1, Number(d.slice(8, 10)));
  const fmt = (ms) => new Date(ms).toISOString().slice(0, 10);
  const pickTop7 = (snap) => snap.stocks
    .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed)
    .sort((a, b) => b.composite - a.composite)
    .slice(0, 7);

  let cursorMs = dateToMs(anchorDate);
  const todayMs = dateToMs(today);
  const segments = [];
  while (cursorMs <= todayMs) {
    const winStart = fmt(cursorMs);
    const winEnd = fmt(cursorMs + (periodDays - 1) * 86400000);
    const entry = snapshots.find((s) => s.date >= winStart && s.date <= winEnd);
    if (!entry) { cursorMs += periodDays * 86400000; continue; }
    const tracking = snapshots.filter((s) => s.date >= entry.date && s.date <= winEnd);
    segments.push({
      index: segments.length,
      label: periodDays === 1
        ? `Day ${segments.length + 1}`
        : periodDays === 7
          ? `Week ${segments.length + 1}`
          : `Month ${segments.length + 1}`,
      startDate: entry.date,
      endDate: tracking[tracking.length - 1].date,
      entrySnap: entry,
      top7: pickTop7(entry),
      tracking,
    });
    cursorMs += periodDays * 86400000;
  }
  return segments;
}

// Composite per-day equity curve across all segments. Each segment's
// average basket return is multiplicatively chained to the previous
// segment's end-factor (so a +5% week followed by +3% week = +8.15%).
// The charger nets out transaction costs: each
// re-lock is a round-trip (sell the old basket, buy the new one), and
// the very first segment is a buy only. Pass 0 for the gross curve.
// opts.booked (default off): once a name reaches its +5% target or −20% SL,
// its contribution freezes at that LEVEL for the rest of the segment — the
// real exit, so a runaway winner can't keep inflating the basket. Mirrors
// the manual basket's book-at-first-target rule so the AI and Manual curves
// are a true like-to-like comparison. Detection is on the daily close; for
// today it widens to the live intraday high/low (client's "hit = touched").
function buildSegmentedEquityCurve(segments, snapshots, anchorDate, chg = ZERO_CHARGER, opts = {}) {
  const { booked = false, todayDate = null, livePrices = {}, capital = 0, tradingDays = null } = opts;
  // Only real trading days may draw a point. Snapshots are written every
  // calendar day, so without this a Saturday and a Sunday both land on the
  // chart carrying Friday's closes -- and the entry charge debited on day
  // one then reads as a price move on a day the market was shut.
  const isTradingDay = (d) => !tradingDays || tradingDays.size === 0 || tradingDays.has(d);
  // Charges are expressed as a fraction of the basket so they can be
  // folded into the return factor. The sell leg carries the flat DP fee,
  // so it depends on position size -- at a small account that fee is the
  // larger half of the round trip and cannot be treated as a rate.
  const perPos = capital > 0 && segments[0]?.top7?.length ? capital / segments[0].top7.length : 0;
  const buyFrac  = perPos > 0 ? chg.buy(perPos) / perPos : 0;
  const sellFrac = perPos > 0 ? chg.sell(perPos) / perPos : 0;
  if (!segments.length) return [];
  const curve = [];
  let prevFactor = 1.0, segStartFactor = 1.0;
  curve.push({ date: anchorDate, retPct: 0 });
  const TGT = 1 + AI_TARGET_PCT, SLF = 1 - AI_SL_PCT;
  segments.forEach((seg, si) => {
    // Charge to (re)lock this basket: a buy on the first segment, a full
    // round-trip (sell previous + buy current) on every re-lock after.
    prevFactor *= (1 - (si === 0 ? buyFrac : buyFrac + sellFrac));
    // Remembered so the intraday mark below can compound today's live move
    // onto the book as it stood BEFORE today, without having to reconstruct
    // that state by division.
    segStartFactor = prevFactor;
    const entryCloses = {};
    for (const s of seg.top7) entryCloses[s.ticker] = s.close;
    const frozenFactor = {};   // ticker -> locked factor (target / SL level)
    let lastFactor = 1.0;
    for (const day of seg.tracking) {
      if (day.date === anchorDate) continue;
      if (!isTradingDay(day.date)) continue;
      let sum = 0, n = 0;
      for (const ticker of Object.keys(entryCloses)) {
        if (booked && frozenFactor[ticker] != null) { sum += frozenFactor[ticker]; n++; continue; }
        const s = day.stocks.find((x) => x.ticker === ticker);
        if (!s || s.close == null) continue;
        let f = s.close / entryCloses[ticker];
        if (booked) {
          let hiF = f, loF = f;
          if (todayDate && day.date === todayDate) {
            const live = livePrices[ticker];
            if (live) {
              if (live.dayHigh != null) hiF = Math.max(hiF, live.dayHigh / entryCloses[ticker]);
              if (live.dayLow  != null) loF = Math.min(loF, live.dayLow  / entryCloses[ticker]);
              if (live.current != null) { const cf = live.current / entryCloses[ticker]; hiF = Math.max(hiF, cf); loF = Math.min(loF, cf); }
            }
          }
          if (hiF >= TGT) { frozenFactor[ticker] = TGT; f = TGT; }
          else if (loF <= SLF) { frozenFactor[ticker] = SLF; f = SLF; }
        }
        sum += f; n++;
      }
      if (n === 0) continue;
      const dayFactor = sum / n;
      lastFactor = dayFactor;
      const cum = (prevFactor * dayFactor - 1) * 100;
      // Replace if the day already exists (segment boundaries).
      const existing = curve.findIndex((p) => p.date === day.date);
      if (existing >= 0) curve[existing] = { date: day.date, retPct: cum };
      else curve.push({ date: day.date, retPct: cum });
    }
    prevFactor *= lastFactor;
  });

  // Today's intraday mark. The last snapshot was written pre-market and
  // carries the previous close, so during a session the curve would
  // otherwise sit still while the basket is plainly moving.
  const q = liveQuotesToday();
  const lastSeg = segments[segments.length - 1];
  if (q && lastSeg?.top7?.length) {
    let sum = 0, n = 0;
    for (const s of lastSeg.top7) {
      const px = q.prices[s.ticker];
      if (px == null || !(s.close > 0)) continue;
      let f = px / s.close;
      // Book at the target / stop exactly as the daily walk does. Without
      // this the chart marked a name at its full live gain while the picks
      // list beside it showed the same name frozen at +5%, and the two
      // headline numbers disagreed by more than half a percent.
      if (booked) {
        const live = livePrices[s.ticker] || {};
        let hiF = f, loF = f;
        if (live.dayHigh != null) hiF = Math.max(hiF, live.dayHigh / s.close);
        if (live.dayLow != null) loF = Math.min(loF, live.dayLow / s.close);
        if (hiF >= TGT) f = TGT;
        else if (loF <= SLF) f = SLF;
      }
      sum += f; n++;
    }
    if (n) {
      const cum = (segStartFactor * (sum / n) - 1) * 100;
      const pt = { date: q.date, retPct: cum, live: true };
      const at = curve.findIndex((p) => p.date === q.date);
      if (at >= 0) curve[at] = pt; else curve.push(pt);
      curve.liveMark = true;
    }
  }

  // The entry charge is a cost, not a price move. Exposed here so the
  // chart can label it instead of letting it masquerade as day-one loss.
  curve.entryCostPct = buyFrac * 100;
  return curve;
}


function buildNiftyCurve(dates, niftyOn) {
  if (!dates.length) return [];
  const startClose = niftyOn(dates[0]);
  if (startClose == null) return [];
  return dates.map((date) => {
    const close = niftyOn(date);
    return { date, retPct: close != null ? (close / startClose - 1) * 100 : null };
  });
}

// Per-stock cumulative-return curves over `dates` for the chart's "Stocks"
// mode (client ask: toggle basket ↔ stocks → one line per holding, for both
// AI and Manual baskets). entryCloses = { ticker: entryClose }; forward-fills
// a missing daily close so an illiquid day doesn't punch a hole in the line.
function buildPerStockCurves(entryCloses, nameByTicker, snapshots, dates, bookingByTicker = null, booked = false) {
  if (!dates?.length) return [];
  const snapByDate = new Map(snapshots.map((s) => [s.date, s]));
  return Object.keys(entryCloses).map((ticker) => {
    let last = entryCloses[ticker];
    // Booked mode: once a name hits its target / SL it exits — freeze the line
    // at that exit level from the exit date so the per-stock lines decompose
    // the (also-booked) basket line instead of drifting past the real exit.
    const bk = (booked && bookingByTicker) ? bookingByTicker.get(ticker) : null;
    const curve = dates.map((date) => {
      let close;
      if (bk && bk.bookPrice != null && date >= bk.bookDate) {
        close = bk.bookPrice;
      } else {
        const snap = snapByDate.get(date);
        const s = snap ? snap.stocks.find((x) => x.ticker === ticker) : null;
        if (s && typeof s.close === "number") last = s.close;
        close = last;
      }
      return { date, retPct: (close / entryCloses[ticker] - 1) * 100 };
    });
    return { ticker, name: nameByTicker[ticker] || ticker, curve };
  });
}

// Both baskets' per-stock line curves for the chart's "Stocks" mode. AI names
// = the latest segment's top-7 (entry at that segment's close); Manual names =
// the in-coverage client picks (entry at the anchor snapshot close).
function buildBasketStockCurves(segments, manualPicks, snapshots, anchorDate, dates, aiPicks, manualRows) {
  const booked = state.manualReturnMode !== "held";
  const lastSeg = (segments || [])[(segments || []).length - 1];
  const aiEntry = {}, aiName = {};
  for (const s of (lastSeg?.top7 || [])) if (s.ticker && typeof s.close === "number") { aiEntry[s.ticker] = s.close; aiName[s.ticker] = s.name || s.ticker; }
  // AI names book at their uniform +5% / −20% level the day they first hit it —
  // mirror the basket's booked convention so the two curves stay compatible.
  const aiBooking = new Map();
  for (const p of (aiPicks || [])) {
    if (!p || !p.ticker || aiEntry[p.ticker] == null || aiBooking.has(p.ticker)) continue;
    if ((p.status === "TARGET_HIT" || p.status === "SL_HIT") && p.hitDate) {
      const bookPrice = p.status === "TARGET_HIT" ? p.target : p.sl;
      if (bookPrice != null) aiBooking.set(p.ticker, { bookDate: p.hitDate, bookPrice });
    }
  }
  const aiStockCurves = buildPerStockCurves(aiEntry, aiName, snapshots, dates, aiBooking, booked);
  const manualEntry = {}, manualName = {};
  const anchorSnap = snapshots.find((s) => s.date >= anchorDate);
  for (const p of (manualPicks || [])) {
    if (!p.in_universe || !p.ticker) continue;
    const s = anchorSnap?.stocks?.find((x) => x.ticker === p.ticker);
    if (s && typeof s.close === "number") { manualEntry[p.ticker] = s.close; manualName[p.ticker] = p.selection || p.ticker; }
  }
  // Manual names book at the client's Target 1 / SL price.
  const manualBooking = new Map();
  for (const r of (manualRows || [])) {
    if (r?.ticker && r.booking?.booked && r.booking.bookDate && r.booking.bookPrice != null) manualBooking.set(r.ticker, { bookDate: r.booking.bookDate, bookPrice: r.booking.bookPrice });
  }
  const manualStockCurves = buildPerStockCurves(manualEntry, manualName, snapshots, dates, manualBooking, booked);
  return { aiStockCurves, manualStockCurves };
}

// Daily picks: every BUY event in the simulator's trade log becomes a
// fresh accuracy-tracked pick. If the same stock re-enters multiple
// times, each entry is tracked separately (per founder confirmation —
// every entry is a fresh prediction with its own entry close).
// Targets follow the framework's uniform +5% / −20% bands.
// Map ticker → GICS-style industry (finer than sector). Industry is a
// stable attribute of a stock, so we read it from the snapshots (last-seen
// wins) rather than threading it through every trade/holding — accurate
// for grouping regardless of a pick's entry date.
function buildIndustryMap(snapshots) {
  const map = new Map();
  for (const snap of (snapshots || [])) {
    for (const s of (snap.stocks || [])) {
      if (s.ticker && s.industry) map.set(s.ticker, s.industry);
    }
  }
  return map;
}
function attachIndustry(picks, snapshots) {
  const indBy = buildIndustryMap(snapshots);
  for (const p of picks) p.industry = indBy.get(p.ticker) || null;
  return picks;
}

function buildActiveDailyPicks(sim, snapshots, todayDate) {
  const picks = [];
  for (const t of sim.trades) {
    if (t.action !== "BUY") continue;
    const entryPrice = t.price;
    const target = entryPrice * (1 + AI_TARGET_PCT);
    const sl = entryPrice * (1 - AI_SL_PCT);
    const status = computeHitStatus(t.ticker, t.date, entryPrice, target, sl, snapshots, todayDate);
    const peak = computePeakStats(t.ticker, t.date, entryPrice, snapshots, todayDate);
    picks.push({
      ticker: t.ticker,
      name: t.name || t.ticker,
      sector: t.sector || null,
      entryDate: t.date,
      entryPrice, target, sl,
      targetPct: AI_TARGET_PCT * 100,
      slPct: -AI_SL_PCT * 100,
      ...status,
      peak,
    });
  }
  return enrichAndSortPicks(attachIndustry(picks, snapshots));
}

// Weekly / Monthly picks: each segment's top 7 contribute 7 picks
// anchored at the segment's start date and close. Different segments
// can repeat the same ticker — each is a fresh pick.
function buildActiveSegmentedPicks(segments, snapshots, todayDate) {
  const picks = [];
  for (const seg of segments) {
    for (const s of seg.top7) {
      if (!s.ticker || s.close == null) continue;
      const target = s.close * (1 + AI_TARGET_PCT);
      const sl = s.close * (1 - AI_SL_PCT);
      const status = computeHitStatus(s.ticker, seg.startDate, s.close, target, sl, snapshots, todayDate);
      const peak = computePeakStats(s.ticker, seg.startDate, s.close, snapshots, todayDate);
      picks.push({
        ticker: s.ticker,
        name: s.name || s.ticker,
        sector: s.sector || null,
        entryDate: seg.startDate,
        entryPrice: s.close,
        target, sl,
        targetPct: AI_TARGET_PCT * 100,
        slPct: -AI_SL_PCT * 100,
        cohortLabel: segments.length > 1 ? seg.label : null,
        ...status,
        peak,
      });
    }
  }
  return enrichAndSortPicks(attachIndustry(picks, snapshots));
}

function enrichAndSortPicks(picks) {
  // Compute current return % + proximity score (1.0 = at target, 0 = at SL)
  // and sort: TARGET_HIT pinned top → OPEN by proximity desc → SL_HIT bottom.
  const enriched = picks.map((p) => {
    const range = p.targetPct - p.slPct;
    const curRet = p.currentClose != null && p.entryPrice
      ? (p.currentClose / p.entryPrice - 1) * 100
      : (p.exitPrice != null && p.entryPrice ? (p.exitPrice / p.entryPrice - 1) * 100 : null);
    let proximity = null;
    if (p.status === "TARGET_HIT") proximity = 1.5;
    else if (p.status === "SL_HIT") proximity = -0.5;
    else if (curRet != null && range > 0) proximity = (curRet - p.slPct) / range;
    return { ...p, currentReturnPct: curRet, proximity };
  });
  enriched.sort((a, b) => (b.proximity ?? -2) - (a.proximity ?? -2));
  return enriched;
}

function computeOverallHitSummary(picks) {
  const total = picks.length;
  const targetHits = picks.filter((p) => p.status === "TARGET_HIT").length;
  const slHits = picks.filter((p) => p.status === "SL_HIT").length;
  const open = total - targetHits - slHits;
  const closed = targetHits + slHits;
  const hitRate = total > 0 ? (targetHits / total) * 100 : null; // share of ALL picks that hit target (open ones count until they resolve)
  const targetHitDays = picks.filter((p) => p.status === "TARGET_HIT" && p.daysToHit != null);
  const slHitDays = picks.filter((p) => p.status === "SL_HIT" && p.daysToHit != null);
  const avgDaysToTarget = targetHitDays.length
    ? targetHitDays.reduce((a, p) => a + p.daysToHit, 0) / targetHitDays.length : null;
  const avgDaysToSL = slHitDays.length
    ? slHitDays.reduce((a, p) => a + p.daysToHit, 0) / slHitDays.length : null;
  return { total, targetHits, slHits, open, closed, hitRate, avgDaysToTarget, avgDaysToSL };
}

function renderActiveShell(view, cadence, anchorDate, todayDate, mode, alertsHtml = "", monthSelectorHtml = "", holdNote = "") {
  const isPassive = mode === "passive";
  const cadenceBar = isPassive ? "" : renderActiveCadencePills(cadence);
  if (!view) {
    return `
      <div class="space-y-4">
        ${monthSelectorHtml}
        ${holdNote}
        ${renderStrategyModeToggle(mode, null)}
        ${cadenceBar}
        ${renderHistoryEmpty(isPassive ? "No passive picks yet — upload a client basket and wait for a snapshot." : "No active picks yet — snapshot trail too short for this cadence.")}
      </div>
    `;
  }
  const hits = collectStrategyHits(view, todayDate);
  const sub = STRATEGY_SUBTABS.includes(state.strategySubTab) ? state.strategySubTab : "plan";
  // Tight top: ONE command bar (title + returns + upload/alerts) then the
  // sub-tab bar sharing a row with the Booked/If-held toggle — so the chart
  // clears the fold instead of sitting under stacked header cards (founder ask).
  return `
    <div id="active-strategy" class="space-y-3">
      ${monthSelectorHtml}
      ${holdNote}
      ${renderDataFreshness()}
      ${renderStrategyCommandBar(view, cadence, mode, hits)}
      ${cadenceBar}
      <div class="flex flex-col sm:flex-row gap-2 sm:items-stretch">
        <div class="flex-1 min-w-0">${renderStrategySubNav(sub)}</div>
        ${renderManualReturnToggle(view)}
      </div>
      ${renderStrategySubPanel(view, sub, mode, anchorDate)}
    </div>
  `;
}

// ============================================================
// TRADE PLAN — turns this month's basket into sized orders.
//
// Everything here is derived, nothing is stored: change the capital box and
// the whole sizing table recomputes. The inputs are the current month's
// cohort (same pickTop7 rule as everywhere else), each stock's ATR / RSI /
// beta / 52-week distance from the snapshot's techVals, live ADTV from
// technicals.json, and the index vs its own 200 DMA for the regime call.
//
// Sizing is inverse-ATR: weight proportional to 1/ATR%, so a stock that
// swings 4.5% a day gets a smaller position than one that swings 2.6%, and
// every holding puts roughly the same rupees at risk. Equal rupees would
// mean the wildest name dominates the month's P&L.
// ============================================================

// Regime: is the benchmark above its own long moving average? Long-only
// momentum bleeds when the index itself is in a downtrend, so this gates
// everything below it. Falls back to the longest MA the history supports
// and says so, rather than silently comparing against a shorter window.
function planRegime(benchmark) {
  const closes = benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes;
  if (!closes) return { ok: null, note: "Benchmark history not loaded." };
  const dates = Object.keys(closes).sort();
  const vals = dates.map((d) => closes[d]);
  const want = 200;
  const n = Math.min(want, vals.length);
  if (n < 50) return { ok: null, note: `Only ${vals.length} days of index history — need ~50 minimum.` };
  const ma = vals.slice(-n).reduce((a, b) => a + b, 0) / n;
  const last = vals[vals.length - 1];
  return {
    ok: last > ma, last, ma, days: n, exact: n === want,
    asOf: dates[dates.length - 1],
  };
}

// How far above yesterday's close we will still pay, in percent. A limit
// pinned exactly at the previous close throws away good names on any
// morning the whole market gaps up — which is a market move, not the
// stock running away from us. So the tolerance follows the overnight
// signal, and is capped at half the stock's own ATR either way: we accept
// the market's move, never the stock's.
function entryTolerancePct() {
  const mc = state.cache.marketContext;
  const t = mc?.signal?.tolerancePct;
  return Number.isFinite(t) ? t : 0.5;
}

// The basket for a given strategy, with everything needed to place an
// order. Returns `core` (what you buy) and `buffer` (the bench), because
// on any given morning some core name will have gapped away from us and
// the substitute has to be pre-sized and ready, not calculated at 9:20.
function planRows(capital, strategy) {
  const cache = state.cache.history;
  const snapshots = cache?.snapshots || [];
  if (!snapshots.length) return null;
  const ym = snapshots[snapshots.length - 1].date.slice(0, 7);
  const first = snapshots.find((s) => s.date.slice(0, 7) === ym);
  if (!first) return null;

  const st = strategy || { basketSize: 7, bufferSize: 3, stopAtr: PLAN_STOP_ATR, invAtrSizing: true, gates: {} };
  const adtvBy = state.cache.planAdtv || {};
  const live = cache.todayClose || {};

  // Rank once, then let the strategy's own gates decide what it will hold.
  const ranked = first.stocks
    .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed)
    .sort((a, b) => b.composite - a.composite)
    .map((s) => {
      const tv = s.techVals || {};
      return {
        ticker: s.ticker, name: s.name, sector: s.sector || "—", score: s.composite,
        ind: { rsi: tv.rsi, atr: tv.atr, d52: tv.d52, adtv: adtvBy[s.ticker] ?? null,
               aboveMa50: tv.e50, aboveMa200: tv.d200, macdPositive: tv.macd },
        techVals: tv, close: s.close,
      };
    });
  if (!ranked.length) return null;

  // Ask for a deeper bench than we intend to show. Some core names turn out
  // to be unbuildable at this capital and get replaced from the bench, and
  // the bench still has to have three names left afterwards -- its real job
  // is covering entries we miss on the morning, not patching sizing.
  const wantBuffer = st.bufferSize || 3;
  const sel = strat5.selectFromRanked(ranked, { ...st, bufferSize: wantBuffer + 4 }, {});
  if (!sel.core.length) return null;

  const decorate = (c) => {
    const tv = c.techVals || {};
    // Anchor every price to the MONTH-START close, exactly as the founder
    // asked ("1 tarik ka closing ko entry limit bana raha hai"). The buy
    // limit, the stop and the entry are all decided off that one number,
    // so the price shown here is the same one the Overview tracks the
    // position from. Letting today's live price in here produced a "buy up
    // to ₹578" for a stock already held from ₹528 -- three prices for one
    // holding. The live price belongs to the Entry Monitor and the
    // Overview, not to the plan's reference price.
    const price = c.close;
    return {
      ticker: c.ticker, name: c.name, sector: c.sector, composite: c.score,
      price, entryClose: c.close, ind: c.ind,
      atr: tv.atr, rsi: tv.rsi, beta: tv.beta, d52: tv.d52,
      adtv: adtvBy[c.ticker] ?? null,
    };
  };
  const rows = sel.core.map(decorate);
  const buffer = sel.buffer.map(decorate);

  // Sizing comes from the strategy itself (risk parity with a per-name cap,
  // or flat slots), so the plan you trade is the plan that was backtested.
  const weightOf = strat5.invAtrWeights(sel.core, st);
  const atrs = rows.map((r) => r.atr).filter((v) => Number.isFinite(v) && v > 0).sort((a, b) => a - b);
  const medAtr = atrs.length ? atrs[Math.floor(atrs.length / 2)] : 3.5;
  const stopAtr = st.stopAtr ?? PLAN_STOP_ATR;
  const tol = entryTolerancePct() / 100;

  const size = (r, coreIdx) => {
    r.atrUsed = Number.isFinite(r.atr) && r.atr > 0 ? r.atr : medAtr;
    r.weight = coreIdx >= 0 ? weightOf(sel.core[coreIdx]) : 1 / (st.basketSize || 7);
    r.target = capital * r.weight;
    r.shares = r.price > 0 ? Math.floor(r.target / r.price) : 0;
    r.cost = r.shares * r.price;
    r.stopPct = stopAtr * r.atrUsed / 100;
    r.stop = r.price * (1 - r.stopPct);
    r.risk = r.cost * r.stopPct;
    // Three prices, decided before the bell: what we'd like to pay, the most
    // we'll accept if the market itself gapped, and the level above which the
    // trade is simply a different trade.
    r.limitIdeal = r.price;
    r.limitMax = r.price * (1 + Math.min(tol, r.atrUsed / 200));
    r.skipAbove = r.price * (1 + r.atrUsed / 100);
    // Can this position even be built? One share of a ₹2,400 stock inside a
    // ₹3,500 slot is not a position, it is a rounding error wearing a hat.
    r.tradable = r.shares >= 3;
    r.oneShare = r.shares > 0 && r.shares < 3;
    r.unaffordable = r.shares === 0;
    return r;
  };
  rows.forEach((r, i) => size(r, i));
  buffer.forEach((r) => size(r, -1));

  // ---- unbuildable names ----------------------------------------------
  // A ₹5,686 share does not fit in a ₹3,679 slot, so that position is zero
  // and its capital sits idle -- at ₹50k across 10 names this was leaving a
  // fifth of the account undeployed. Swap in the best bench name we CAN
  // afford, and if none is affordable, drop the slot and spread its money
  // over the rest. Either way the book gets fully invested and the risk
  // weighting still holds across whatever actually got bought.
  const skipped = [];
  let coreSel = [...sel.core];
  let finalRows = [...rows];
  const benchPool = [...buffer];
  for (let i = finalRows.length - 1; i >= 0; i--) {
    if (finalRows[i].shares > 0) continue;
    skipped.push({ name: finalRows[i].name, price: finalRows[i].price, target: finalRows[i].target });
    const bi = benchPool.findIndex((b) => b.price > 0 && finalRows[i].target / b.price >= 1);
    if (bi >= 0) {
      const promoted = benchPool.splice(bi, 1)[0];
      promoted.promoted = true;
      promoted.replaces = finalRows[i].name;
      finalRows[i] = promoted;
      coreSel[i] = sel.buffer.find((c) => c.ticker === promoted.ticker) || coreSel[i];
    } else {
      finalRows.splice(i, 1);
      coreSel.splice(i, 1);
    }
  }
  // Re-weight over what survived so nothing is left on the table.
  if (skipped.length) {
    const w2 = strat5.invAtrWeights(coreSel, { ...st, basketSize: coreSel.length });
    finalRows.forEach((r, i) => { r.weight = w2(coreSel[i]); size(r, -2); r.weight = w2(coreSel[i]);
      r.target = capital * r.weight;
      r.shares = r.price > 0 ? Math.floor(r.target / r.price) : 0;
      r.cost = r.shares * r.price;
      r.risk = r.cost * r.stopPct;
      r.tradable = r.shares >= 3;
    });
  }
  rows.length = 0;
  rows.push(...finalRows);
  const bench = benchPool.slice(0, wantBuffer);

  // ---- leftover cash sweep --------------------------------------------
  // Share counts round DOWN, not to nearest: rounding up overspends the
  // slot, and on a ₹4,849 share inside a ₹3,832 slot "nearest" buys one
  // anyway and blows the position 27% over target. Rounding down alone
  // leaves a few thousand rupees idle, so the change is spent back here --
  // one share at a time, always to whichever holding sits furthest below
  // its target. Same result as rounding to nearest when that was safe,
  // without ever overshooting a slot.
  let leftover = capital - rows.reduce((a, r) => a + r.cost, 0);
  for (let guard = 0; guard < 50; guard++) {
    const candidates = rows
      .filter((r) => r.price > 0 && r.price <= leftover && r.cost + r.price <= r.target * 1.15)
      .sort((a, b) => (a.cost - a.target) - (b.cost - b.target));
    if (!candidates.length) break;
    const r = candidates[0];
    r.shares += 1;
    r.cost = r.shares * r.price;
    r.risk = r.cost * r.stopPct;
    r.tradable = r.shares >= 3;
    leftover -= r.price;
  }

  let totalCost = 0, totalRisk = 0;
  rows.forEach((r) => { totalCost += r.cost; totalRisk += r.risk; });

  // Flags — each one is a number against a threshold, nothing subjective.
  const secCount = {};
  rows.forEach((r) => { secCount[r.sector] = (secCount[r.sector] || 0) + 1; });
  const heavySectors = Object.entries(secCount).filter(([, n]) => n > PLAN_MAX_PER_SECTOR);
  [...rows, ...bench].forEach((r) => {
    r.hot = Number.isFinite(r.rsi) && r.rsi > PLAN_RSI_HOT;
    r.thin = Number.isFinite(r.adtv) && r.adtv < 2;
    r.atHigh = Number.isFinite(r.d52) && r.d52 <= 2;
  });

  const roundTripPct = rows.length ? roundTrip(capital / rows.length, simPrefs).pct : 0;
  return {
    rows, buffer: bench, totalCost, totalRisk, capital, heavySectors, secCount,
    month: ym, anchorDate: first.date, strategy: st, tolerancePct: tol * 100, roundTripPct,
    tooSmall: rows.filter((r) => !r.tradable), skipped,
  };
}

const inr = (n) => "₹" + Math.round(n).toLocaleString("en-IN");

// ---- Entry monitor -------------------------------------------------
// On entry morning the question is not "what should I buy" — that was
// settled last night — but "what did I actually get, and what do I do
// about what I missed". This reads the live quote for every core name
// and classifies it against the limit that was decided before the open.
//
// A miss is only useful if it says WHY. "Gapped 3.2% above your limit" is
// actionable; "missed" is not.
function entryStatus(row, live) {
  if (!live || live.current == null) return { code: "unknown", label: "No live price", tone: "slate" };
  const { current, open, dayLow, dayHigh, prevClose } = live;
  const limit = row.limitMax;
  const awayPct = ((current - limit) / limit) * 100;

  // Filled: price is at or under the limit right now.
  if (current <= limit) return { code: "filled", label: "At or below your limit", tone: "emerald", awayPct };

  // The low of the day reached the limit — the order should have filled
  // even though it has since moved up.
  if (dayLow != null && dayLow <= limit) {
    return { code: "filled_earlier", label: "Traded at your limit earlier today", tone: "emerald", awayPct };
  }

  // Never came back: distinguish a gap at the open from an intraday run,
  // because they call for different responses.
  const gapPct = open != null && prevClose ? ((open - prevClose) / prevClose) * 100 : null;
  if (open != null && open > limit) {
    return {
      code: "missed_gap", tone: "rose", awayPct, gapPct,
      label: gapPct != null
        ? `Opened ${gapPct >= 0 ? "+" : ""}${gapPct.toFixed(1)}% — above your limit from the first tick`
        : "Opened above your limit",
    };
  }
  return { code: "missed_ran", label: `Ran ${awayPct.toFixed(1)}% past your limit intraday`, tone: "amber", awayPct };
}

// Pair each missed core name with the best available bench name, and
// re-size the substitute at ITS own price and ATR so the rupees at risk
// stay where they were. Simply copying the quantity across would change
// the risk of the book, which is the one thing the sizing rule exists to
// hold constant.
function planSubstitutions(plan, livePrices) {
  const out = [];
  const used = new Set();
  for (const r of plan.rows) {
    const st = entryStatus(r, livePrices[r.ticker]);
    if (st.code !== "missed_gap" && st.code !== "missed_ran") continue;
    const sub = plan.buffer.find((b) => {
      if (used.has(b.ticker)) return false;
      const bl = livePrices[b.ticker];
      // Only offer a substitute we could actually buy right now.
      return !bl || bl.current == null || bl.current <= b.skipAbove;
    });
    if (sub) {
      used.add(sub.ticker);
      const live = livePrices[sub.ticker];
      const px = live?.current ?? sub.price;
      const shares = px > 0 ? Math.floor(r.target / px) : 0;
      out.push({ missed: r, status: st, sub, subPrice: px, shares, cost: shares * px,
                 stop: px * (1 - sub.stopPct), risk: shares * px * sub.stopPct });
    } else {
      out.push({ missed: r, status: st, sub: null });
    }
  }
  return out;
}

function renderEntryMonitor(plan) {
  const livePrices = state.cache.history?.livePrices || {};
  if (!Object.keys(livePrices).length) return "";
  // The Entry Monitor answers "did my orders fill this morning" — a
  // question that only exists on entry day. Past the first couple of
  // trading days the basket is bought and held, so comparing today's live
  // price against a day-one limit just flags every winner as "missed".
  // Count trading days since the basket was locked, and hide it after that.
  const cal = state.cache.history?.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || {};
  const anchor = plan.anchorDate;
  const today = istTodayDate();
  let tradedSince = 0;
  for (const d of Object.keys(cal)) if (d > anchor && d <= today) tradedSince++;
  // Once a full trading day has closed since the basket locked, entry is
  // over — you either filled or you didn't, and the Overview shows what you
  // now hold. Only show this on entry morning itself (no completed session
  // yet) and only when we actually have live prices to compare against.
  if (tradedSince >= 1 || !liveQuotesToday()) return "";
  const rows = plan.rows.map((r) => ({ r, st: entryStatus(r, livePrices[r.ticker]) }));
  const known = rows.filter((x) => x.st.code !== "unknown");
  if (!known.length) return "";
  const missed = rows.filter((x) => x.st.code.startsWith("missed"));
  const subs = planSubstitutions(plan, livePrices);
  const tone = { emerald: "text-emerald-700", rose: "text-rose-700", amber: "text-amber-700", slate: "text-slate-400" };
  const dot = { emerald: "bg-emerald-500", rose: "bg-rose-500", amber: "bg-amber-500", slate: "bg-slate-300" };

  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <div class="px-4 py-3 border-b border-slate-100 flex items-center justify-between gap-2 flex-wrap">
        <div>
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Entry monitor · live</div>
          <div class="text-xs text-slate-500 mt-0.5">${missed.length ? `${missed.length} of ${plan.rows.length} moved past your limit` : "Every name is still inside your limit"}</div>
        </div>
        ${missed.length ? `<span class="text-[10px] font-bold uppercase tracking-wider px-2 py-1 rounded-full bg-rose-50 text-rose-700 ring-1 ring-rose-200">${missed.length} to replace</span>` : ""}
      </div>
      <div class="divide-y divide-slate-100">
        ${rows.map(({ r, st }) => `
          <div class="px-4 py-2 flex items-center gap-3 text-sm">
            <span class="w-1.5 h-1.5 rounded-full flex-shrink-0 ${dot[st.tone]}"></span>
            <span class="font-semibold text-slate-900 truncate flex-1 min-w-0">${escapeHtml(r.name)}</span>
            <span class="text-xs ${tone[st.tone]} text-right">${escapeHtml(st.label)}</span>
            <span class="text-xs tabular-nums text-slate-400 w-24 text-right">limit ${inr(r.limitMax)}</span>
          </div>`).join("")}
      </div>
      ${subs.length ? `
        <div class="px-4 py-3 bg-slate-50/70 border-t border-slate-100">
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-2">Buy instead</div>
          <div class="space-y-1.5">
            ${subs.map((s) => s.sub ? `
              <div class="text-[13px] text-slate-700 flex flex-wrap items-baseline gap-x-1.5">
                <span class="text-slate-400">${escapeHtml(s.missed.name)} →</span>
                <span class="font-semibold text-slate-900">${escapeHtml(s.sub.name)}</span>
                <span class="tabular-nums">· <b>${s.shares}</b> shares at ${inr(s.subPrice)} = ${inr(s.cost)}</span>
                <span class="tabular-nums text-rose-700">· stop ${inr(s.stop)}</span>
              </div>` : `
              <div class="text-[13px] text-slate-500">${escapeHtml(s.missed.name)} → no bench name available under its skip price. Hold the cash.</div>`).join("")}
          </div>
        </div>` : ""}
    </div>`;
}

// Strategy picker — the five plans, one click apart. The crown marks the
// one committed as #1 in data/strategies.json; each pill carries its
// backtested CAGR so the choice is made against evidence, not vibes.
function renderPlanStrategyPills() {
  const defs = strat5.allStrategies();
  if (!defs.length) return "";
  const bt = state.cache.strategyBacktest;
  const primary = strat5.primaryId();
  const active = state.planStrategyId || primary;
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-1.5 flex flex-wrap gap-1">
      ${defs.map((d) => {
        const on = d.id === active;
        const m = bt?.strategies?.find((x) => x.id === d.id)?.metrics;
        const cagr = m?.cagr;
        return `<button type="button" data-plan-strategy="${d.id}"
          class="flex-1 min-w-[104px] px-2.5 py-2 rounded-xl text-left transition ${on ? "bg-indigo-600 text-white shadow-sm" : "hover:bg-slate-50"}">
          <div class="flex items-center gap-1">
            <span class="text-[13px] font-bold leading-tight ${on ? "text-white" : "text-slate-800"}">${escapeHtml(d.name)}</span>
            ${d.id === primary ? `<span title="Your #1 strategy">★</span>` : ""}
            ${d.isOriginal ? `<span class="text-[8px] font-bold uppercase tracking-wider px-1 py-0.5 rounded ${on ? "bg-white/25 text-white" : "bg-slate-100 text-slate-500"}">current</span>` : ""}
          </div>
          <div class="text-[10px] leading-tight mt-0.5 ${on ? "text-indigo-100" : "text-slate-400"}">
            ${cagr != null ? `${cagr >= 0 ? "+" : ""}${cagr.toFixed(1)}% a year` : escapeHtml(d.tagline || "")}
          </div>
        </button>`;
      }).join("")}
    </div>`;
}

// Pre-open context: what happened overnight, and how jumpy the market is.
// Advisory only — it sets how far above yesterday's close we will pay, and
// warns when limits are unlikely to fill. It never picks a stock.
function renderMarketContext() {
  const mc = state.cache.marketContext;
  if (!mc?.signal) return "";
  const s = mc.signal;
  const f = mc.feeds || {};
  const tone = s.dir === "up" ? "emerald" : s.dir === "down" ? "rose" : "slate";
  const chip = (label, v) => v?.changePct == null ? "" :
    `<span class="text-[11px] text-slate-500">${label} <b class="tabular-nums ${v.changePct >= 0 ? "text-emerald-700" : "text-rose-700"}">${v.changePct >= 0 ? "+" : ""}${v.changePct.toFixed(2)}%</b></span>`;
  return `
    <div class="bg-${tone}-50 ring-1 ring-${tone}-200 rounded-2xl px-4 py-3 flex flex-wrap items-center gap-x-4 gap-y-1.5">
      <span class="text-sm font-semibold text-${tone}-900">${escapeHtml(s.headline)}</span>
      ${chip("US futures", f.spFutures)}
      ${chip("Nikkei", f.nikkei)}
      ${chip("Hang Seng", f.hangSeng)}
      ${s.vix != null ? `<span class="text-[11px] text-slate-500">India VIX <b class="tabular-nums text-slate-700">${s.vix.toFixed(1)}</b> <span class="text-slate-400">${escapeHtml(s.vixBand)}</span></span>` : ""}
      <span class="text-[11px] text-slate-500 ml-auto">Paying up to <b class="text-slate-700">+${s.tolerancePct}%</b> over yesterday\u2019s close today</span>
    </div>`;
}

function renderTradePlan() {
  const activeId = state.planStrategyId || strat5.primaryId();
  const strategy = strat5.strategyById(activeId);
  const plan = planRows(state.planCapital, strategy);
  const regime = planRegime(state.cache.history?.benchmark);
  const pills = renderPlanStrategyPills();
  const context = renderMarketContext();

  const capBox = `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 flex flex-wrap items-end gap-4">
      <div>
        <label for="plan-capital" class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Capital to deploy</label>
        <div class="flex items-center gap-2">
          <span class="text-slate-400 text-lg font-semibold">₹</span>
          <input id="plan-capital" type="number" min="1000" step="1000" value="${state.planCapital}"
            class="w-40 px-3 py-2 text-lg font-bold tabular-nums bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white">
        </div>
      </div>
      <div class="flex items-center gap-1.5">
        ${[25000, 50000, 60000, 100000].map((v) => `
          <button type="button" data-plan-preset="${v}" class="px-2.5 py-1.5 text-xs font-semibold rounded-lg ring-1 transition ${state.planCapital === v ? "bg-indigo-600 text-white ring-indigo-600" : "bg-white text-slate-600 ring-slate-200 hover:bg-slate-50"}">${v >= 100000 ? (v / 100000) + "L" : (v / 1000) + "k"}</button>`).join("")}
      </div>
      <div class="text-[11px] text-slate-500 leading-snug max-w-xs">${plan ? `Sizing follows <span class="font-semibold">${escapeHtml(strategy?.name || "the strategy")}</span>. Buying and selling one position costs <span class="font-semibold">${plan.roundTripPct.toFixed(2)}%</span> at this size.` : "Sizing recomputes instantly."}</div>
    </div>`;

  if (!plan) {
    return `<div class="space-y-4">${pills}${capBox}
      <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-8 text-center text-sm text-slate-500">
        No basket for this month yet — the plan appears once the first snapshot of the month exists.
      </div></div>`;
  }

  // ---- one line for market conditions -------------------------------
  // The overnight strip and the regime check were two separate cards
  // saying the same kind of thing. Merged: one row, one verdict.
  const rOk = regime.ok;
  const rTone = rOk === null ? "slate" : rOk ? "emerald" : "rose";
  const regimeChip = `
    <span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg bg-${rTone}-50 ring-1 ring-${rTone}-200 text-[11px] font-semibold text-${rTone}-800">
      ${rOk === null ? "?" : rOk ? "✓" : "✕"} ${rOk === null ? "Regime unknown" : rOk ? "Risk-on" : "Risk-off — half size"}
    </span>`;

  // ---- warnings, as short chips ---------------------------------------
  const warns = [];
  plan.rows.filter((r) => r.hot).forEach((r) => warns.push(`${escapeHtml(r.name)} · RSI ${r.rsi}, already run`));
  plan.heavySectors.forEach(([sec, n]) => warns.push(`${n} of ${plan.rows.length} in ${escapeHtml(sec)}`));
  plan.rows.filter((r) => r.thin).forEach((r) => warns.push(`${escapeHtml(r.name)} · thin, ₹${r.adtv} Cr a day`));
  plan.skipped.forEach((k) => {
    const swap = plan.rows.find((r) => r.replaces === k.name);
    warns.push(`${escapeHtml(k.name)} too pricey for its slot${swap ? ` → ${escapeHtml(swap.name)}` : " → dropped"}`);
  });
  const warnRow = warns.length ? `
    <div class="flex flex-wrap items-center gap-1.5">
      ${warns.map((w) => `<span class="text-[11px] px-2 py-1 rounded-lg bg-amber-50 ring-1 ring-amber-200 text-amber-900">${w}</span>`).join("")}
    </div>` : "";

  const cell = (v, tone) => `<td class="px-3 py-2 text-right tabular-nums ${tone || ""}">${v}</td>`;

  // ---- the one table you trade from -----------------------------------
  // Was two tables of eight and nine columns -- what the screener says,
  // then what to buy -- read together on every single order. Merged into
  // one, with the diagnostic columns (ATR, RSI, beta, ADTV, off-high)
  // moved behind an expander since they inform the flags above rather
  // than the order itself.
  const deployPct = plan.capital > 0 ? (plan.totalCost / plan.capital) * 100 : 0;
  const riskPct = plan.capital > 0 ? (plan.totalRisk / plan.capital) * 100 : 0;
  const orderTable = `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <div class="px-4 py-3 border-b border-slate-100 flex flex-wrap items-baseline justify-between gap-2">
        <div class="font-semibold text-slate-900 text-sm">Buy these ${plan.rows.length}</div>
        <div class="text-[11px] text-slate-500">Limit orders only · stop = ${plan.strategy.stopAtr ?? PLAN_STOP_ATR} × ATR below entry</div>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-sm">
          <thead class="bg-slate-50/70 text-[10px] uppercase tracking-wider text-slate-500">
            <tr><th class="px-3 py-2 text-left">Stock</th>
            <th class="px-3 py-2 text-right" title="The most to pay — allows for the market's own gap">Buy up to</th>
            <th class="px-3 py-2 text-right" title="Above this, take a bench name instead">Skip above</th>
            <th class="px-3 py-2 text-right">Shares</th><th class="px-3 py-2 text-right">Cost</th>
            <th class="px-3 py-2 text-right">Stop</th><th class="px-3 py-2 text-right">Risk</th></tr>
          </thead>
          <tbody>
            ${plan.rows.map((r) => `
              <tr class="border-t border-slate-100">
                <td class="px-3 py-2">
                  <span class="font-semibold text-slate-900">${escapeHtml(r.name)}</span>
                  ${r.promoted ? `<span class="ml-1.5 text-[9px] font-bold uppercase px-1 py-0.5 rounded bg-sky-100 text-sky-700">bench</span>` : ""}
                  <div class="text-[10px] text-slate-400">${escapeHtml(String(r.sector).slice(0, 24))}</div>
                </td>
                ${cell(inr(r.limitMax), "font-semibold text-slate-900")}
                ${cell(inr(r.skipAbove), "text-slate-400")}
                ${cell(`<span class="text-base font-bold ${r.tradable ? "text-slate-900" : "text-amber-700"}">${r.shares}</span>`)}
                ${cell(inr(r.cost), "font-semibold")}
                ${cell(inr(r.stop), "text-rose-700 font-bold")}
                ${cell(inr(r.risk), "text-slate-500")}
              </tr>`).join("")}
          </tbody>
          <tfoot class="bg-slate-50 border-t-2 border-slate-200 font-bold">
            <tr><td class="px-3 py-2.5">Total</td><td></td><td></td><td></td>
              ${cell(inr(plan.totalCost))}
              <td class="px-3 py-2.5 text-right text-[10px] uppercase tracking-wider text-slate-500">at risk</td>
              ${cell(inr(plan.totalRisk), "text-rose-700")}</tr>
          </tfoot>
        </table>
      </div>
      <div class="px-4 py-2.5 bg-slate-50/60 border-t border-slate-100 text-xs text-slate-600">
        <span class="font-bold text-slate-900">${deployPct.toFixed(0)}%</span> deployed ·
        <span class="font-bold ${riskPct > 12 ? "text-rose-700" : "text-slate-900"}">${riskPct.toFixed(1)}%</span> at risk if every stop hits ·
        <span class="text-slate-400">${plan.roundTripPct.toFixed(2)}% charges a round trip</span>
      </div>
      <details class="border-t border-slate-100">
        <summary class="cursor-pointer px-4 py-2 text-[11px] font-semibold text-slate-500 hover:text-slate-700 select-none">Why these names ▾</summary>
        <div class="overflow-x-auto pb-2">
          <table class="w-full text-sm">
            <thead class="bg-slate-50/70 text-[10px] uppercase tracking-wider text-slate-500">
              <tr><th class="px-3 py-2 text-left">Stock</th><th class="px-3 py-2 text-right">Score</th>
              <th class="px-3 py-2 text-right">Price</th><th class="px-3 py-2 text-right">ATR</th>
              <th class="px-3 py-2 text-right">RSI</th><th class="px-3 py-2 text-right">Beta</th>
              <th class="px-3 py-2 text-right">ADTV</th><th class="px-3 py-2 text-right">Off high</th></tr>
            </thead>
            <tbody>
              ${plan.rows.map((r) => `
                <tr class="border-t border-slate-100 ${r.hot ? "bg-amber-50/40" : ""}">
                  <td class="px-3 py-2 text-slate-700">${escapeHtml(r.name)}</td>
                  ${cell(r.composite, "font-bold text-indigo-700")}
                  ${cell(inr(r.price))}
                  ${cell(r.atr != null ? r.atr + "%" : "—", r.atr > 4 ? "text-rose-600" : "")}
                  ${cell(r.rsi ?? "—", r.hot ? "text-rose-600 font-bold" : "")}
                  ${cell(r.beta ?? "—")}
                  ${cell(r.adtv != null ? "₹" + r.adtv + " Cr" : "—", r.thin ? "text-amber-700 font-bold" : "")}
                  ${cell(r.d52 != null ? r.d52 + "%" : "—", r.atHigh ? "text-amber-700" : "")}
                </tr>`).join("")}
            </tbody>
          </table>
        </div>
      </details>
    </div>`;

  // ---- the bench, four columns ----------------------------------------
  const bufferTable = plan.buffer.length ? `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <div class="px-4 py-2.5 border-b border-slate-100 flex flex-wrap items-baseline justify-between gap-2">
        <div class="font-semibold text-slate-900 text-sm">Bench · ${plan.buffer.length}</div>
        <div class="text-[11px] text-slate-500">If one above gaps past its skip price, buy the top bench name instead</div>
      </div>
      <table class="w-full text-sm">
        <tbody>
          ${plan.buffer.map((r) => `
            <tr class="border-t border-slate-100 first:border-t-0">
              <td class="px-3 py-2"><span class="font-semibold text-slate-800">${escapeHtml(r.name)}</span></td>
              ${cell(`<span class="text-slate-400 text-[11px]">buy up to</span> ${inr(r.limitMax)}`)}
              ${cell(`<span class="font-bold">${r.shares}</span> <span class="text-slate-400 text-[11px]">shares</span>`)}
              ${cell(`<span class="text-slate-400 text-[11px]">stop</span> <span class="text-rose-700 font-semibold">${inr(r.stop)}</span>`)}
            </tr>`).join("")}
        </tbody>
      </table>
    </div>` : "";

  const entryMonitor = renderEntryMonitor(plan);

  // ---- upside, folded away ---------------------------------------------
  // Reference levels, consulted rarely. It was a seven-column table plus
  // three tiles plus two paragraphs, sitting open above the paper results.
  const r2Total = plan.rows.reduce((a, r) => a + r.cost * r.stopPct * 2, 0);
  const upCard = `
    <details class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <summary class="cursor-pointer px-4 py-3 flex flex-wrap items-baseline justify-between gap-2 hover:bg-slate-50 select-none">
        <span class="font-semibold text-slate-900 text-sm">Upside reference</span>
        <span class="text-[11px] text-slate-500">Every stop hits <b class="text-rose-700">−${inr(plan.totalRisk)}</b> · every name reaches 2R <b class="text-emerald-700">+${inr(r2Total)}</b></span>
      </summary>
      <div class="overflow-x-auto border-t border-slate-100">
        <table class="w-full text-sm">
          <thead class="bg-slate-50/70 text-[10px] uppercase tracking-wider text-slate-500">
            <tr><th class="px-3 py-2 text-left">Stock</th><th class="px-3 py-2 text-right">Stop</th>
            <th class="px-3 py-2 text-right">Entry</th><th class="px-3 py-2 text-right">+1R</th>
            <th class="px-3 py-2 text-right">+2R</th><th class="px-3 py-2 text-right">+3R</th></tr>
          </thead>
          <tbody>
            ${plan.rows.map((r) => `
              <tr class="border-t border-slate-100">
                <td class="px-3 py-2 text-slate-700">${escapeHtml(r.name)}</td>
                ${cell(inr(r.stop), "text-rose-700")}
                ${cell(inr(r.price), "text-slate-500")}
                ${cell(inr(r.price * (1 + r.stopPct)), "text-emerald-700")}
                ${cell(inr(r.price * (1 + r.stopPct * 2)), "text-emerald-700 font-bold")}
                ${cell(inr(r.price * (1 + r.stopPct * 3)), "text-emerald-700")}
              </tr>`).join("")}
          </tbody>
        </table>
      </div>
      <div class="px-4 py-2.5 text-[11px] text-slate-500 border-t border-slate-100">R = the stop distance. Reference levels, not predictions — momentum pays through a few large winners, so there is no fixed sell target here.</div>
    </details>`;

  const conditions = `
    <div class="flex flex-wrap items-center gap-2">
      ${regimeChip}
      ${regime.last ? `<span class="text-[11px] text-slate-500 tabular-nums">Smallcap 250 ${Math.round(regime.last).toLocaleString("en-IN")} vs ${regime.days}d avg ${Math.round(regime.ma).toLocaleString("en-IN")}</span>` : ""}
    </div>`;

  // One line so the two tabs never blur: this tab is the shopping list,
  // Overview is the scoreboard. Prices here are the month-start entry
  // basis, so a stock's price here matches its "bought at" in Overview.
  const intro = `
    <div class="text-[12px] text-slate-500 leading-snug px-1">
      <b class="text-slate-700">What to buy this month</b> — the list, the price to buy at, and where the stop sits. All prices are this month's entry (${fmtDateDMY(plan.anchorDate)}). For how the basket is doing since then, see <b class="text-slate-600">Overview</b>.
    </div>`;
  return `<div class="space-y-3">${pills}${intro}${context}${conditions}${capBox}${warnRow}${entryMonitor}${orderTable}${bufferTable}${upCard}${renderPaperResults()}</div>`;
}

// ============================================================
// PAPER RESULTS — how the plan's basket has actually done.
//
// Nothing is stored. Every month's basket is re-derived from the snapshot
// trail on the fly: the cohort is the top 7 of that month's FIRST snapshot,
// and the daily prices come from every snapshot after it. Because snapshots
// are permanent files, past months stay viewable forever without any
// separate archive to keep in sync.
//
// The stop rule is applied faithfully: walking forward day by day, the first
// close at or below entry - 2.5 x ATR ends that position and freezes its
// return there. Note this checks CLOSES only, because that is all a snapshot
// carries -- a real intraday stop would have triggered earlier and at a
// worse price, so these figures are mildly optimistic.
// ============================================================

// Months that have at least one snapshot, newest first.
function paperMonths() {
  const snaps = state.cache.history?.snapshots || [];
  const set = [];
  for (const s of snaps) { const ym = s.date.slice(0, 7); if (!set.includes(ym)) set.push(ym); }
  return set.reverse();
}

function paperTrack(ym) {
  const cache = state.cache.history;
  const snaps = (cache?.snapshots || []).filter((s) => s.date.slice(0, 7) === ym);
  if (!snaps.length) return null;
  const anchor = snaps[0];

  const top = anchor.stocks
    .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed)
    .sort((a, b) => b.composite - a.composite)
    .slice(0, 7);
  if (!top.length) return null;

  // Same inverse-ATR weights the plan told you to buy with.
  const atrOf = (s) => (Number.isFinite(s.techVals?.atr) && s.techVals.atr > 0 ? s.techVals.atr : 3.5);
  const invSum = top.reduce((a, s) => a + 1 / atrOf(s), 0);

  // Price per ticker per date across this month's snapshots.
  const px = {};
  for (const snap of snaps) {
    for (const s of snap.stocks) {
      if (!s.ticker || typeof s.close !== "number") continue;
      (px[s.ticker] || (px[s.ticker] = {}))[snap.date] = s.close;
    }
  }
  const dates = snaps.map((s) => s.date);

  const positions = top.map((s) => {
    const atr = atrOf(s), weight = (1 / atr) / invSum;
    const entry = px[s.ticker]?.[anchor.date] ?? s.close;
    const stopPct = PLAN_STOP_ATR * atr / 100;
    const stop = entry * (1 - stopPct);
    let exit = null, exitDate = null, status = "open";
    for (const d of dates.slice(1)) {
      const c = px[s.ticker]?.[d];
      if (c == null) continue;
      if (c <= stop) { exit = stop; exitDate = d; status = "stopped"; break; }
    }
    let last = entry;
    for (let i = dates.length - 1; i >= 0; i--) { const c = px[s.ticker]?.[dates[i]]; if (c != null) { last = c; break; } }
    const close = exit ?? last;
    return {
      ticker: s.ticker, name: s.name, sector: s.sector, atr, weight,
      entry, stop, stopPct, status, exitDate, current: close,
      ret: entry > 0 ? (close / entry - 1) * 100 : 0,
    };
  });

  // Daily weighted basket curve. A stopped position is frozen at its stop
  // from the day it triggered — it stops contributing further moves.
  const curve = dates.map((d) => {
    let tot = 0;
    for (const p of positions) {
      const stoppedBy = p.exitDate && d >= p.exitDate;
      const c = stoppedBy ? p.stop : (px[p.ticker]?.[d] ?? p.entry);
      tot += p.weight * ((c / p.entry) - 1);
    }
    return { date: d, ret: tot * 100 };
  });

  // Benchmark over the identical window.
  const bcl = cache?.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || null;
  let bench = null;
  if (bcl) {
    const at = (d) => {
      if (bcl[d] != null) return bcl[d];
      let last = null;
      for (const k of Object.keys(bcl).sort()) { if (k <= d) last = bcl[k]; else break; }
      return last;
    };
    const b0 = at(anchor.date);
    if (b0) bench = dates.map((d) => ({ date: d, ret: ((at(d) ?? b0) / b0 - 1) * 100 }));
  }

  const basketRet = curve[curve.length - 1]?.ret ?? 0;
  const benchRet = bench ? bench[bench.length - 1].ret : null;
  return {
    ym, anchorDate: anchor.date, lastDate: dates[dates.length - 1], days: dates.length,
    positions, curve, bench, basketRet, benchRet,
    alpha: benchRet == null ? null : basketRet - benchRet,
    stops: positions.filter((p) => p.status === "stopped").length,
    winners: positions.filter((p) => p.ret > 0).length,
    peak: Math.max(...curve.map((c) => c.ret)),
    trough: Math.min(...curve.map((c) => c.ret)),
  };
}

// Two-line chart: basket vs benchmark, both in % from the anchor day.
function paperChart(t) {
  if (t.curve.length < 2) {
    return `<div class="px-4 py-8 text-center text-xs text-slate-400">The line appears from the second day of the month — one day is a single point.</div>`;
  }
  const W = 780, H = 150, M = { t: 12, r: 12, b: 20, l: 44 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const all = [...t.curve.map((c) => c.ret), ...(t.bench ? t.bench.map((c) => c.ret) : []), 0];
  let lo = Math.min(...all), hi = Math.max(...all);
  const pad = Math.max((hi - lo) * 0.15, 0.6); lo -= pad; hi += pad;
  const x = (i) => M.l + (i / (t.curve.length - 1)) * iw;
  const y = (v) => M.t + ih - ((v - lo) / (hi - lo)) * ih;
  const path = (arr) => arr.map((c, i) => `${i ? "L" : "M"} ${x(i).toFixed(1)} ${y(c.ret).toFixed(1)}`).join(" ");
  const zeroY = y(0);
  return `
    <div class="px-2 pt-2 pb-1 overflow-x-auto">
      <svg viewBox="0 0 ${W} ${H}" class="w-full" style="min-width:420px;max-height:170px" role="img" aria-label="Basket versus benchmark return">
        <line x1="${M.l}" y1="${zeroY}" x2="${W - M.r}" y2="${zeroY}" stroke="currentColor" class="text-slate-300" stroke-width="1" stroke-dasharray="3 3"/>
        <text x="${M.l - 6}" y="${zeroY + 3.5}" text-anchor="end" class="fill-slate-400" style="font-size:9px">0%</text>
        <text x="${M.l - 6}" y="${y(hi) + 8}" text-anchor="end" class="fill-slate-400" style="font-size:9px">${hi.toFixed(1)}%</text>
        <text x="${M.l - 6}" y="${y(lo)}" text-anchor="end" class="fill-slate-400" style="font-size:9px">${lo.toFixed(1)}%</text>
        ${t.bench ? `<path d="${path(t.bench)}" fill="none" stroke="#94a3b8" stroke-width="1.6" stroke-dasharray="4 3"/>` : ""}
        <path d="${path(t.curve)}" fill="none" stroke="#4f46e5" stroke-width="2.2" stroke-linejoin="round"/>
        <circle cx="${x(t.curve.length - 1)}" cy="${y(t.basketRet)}" r="3.5" fill="#4f46e5"/>
        <text x="${M.l}" y="${H - 5}" class="fill-slate-400" style="font-size:9px">${fmtDateDMY(t.anchorDate)}</text>
        <text x="${W - M.r}" y="${H - 5}" text-anchor="end" class="fill-slate-400" style="font-size:9px">${fmtDateDMY(t.lastDate)}</text>
      </svg>
      <div class="flex items-center gap-4 px-3 pb-1 text-[10px] text-slate-500">
        <span class="inline-flex items-center gap-1.5"><span class="w-3 h-0.5 bg-indigo-600"></span>AI basket</span>
        ${t.bench ? `<span class="inline-flex items-center gap-1.5"><span class="w-3 h-0 border-t border-dashed border-slate-400"></span>Smallcap 250</span>` : ""}
      </div>
    </div>`;
}

function renderPaperResults() {
  const months = paperMonths();
  if (!months.length) return "";
  const sel = months.includes(state.paperMonth) ? state.paperMonth : months[0];
  const t = paperTrack(sel);
  if (!t) return "";

  const pct = (v, big) => {
    if (v == null) return `<span class="text-slate-400">—</span>`;
    const c = v > 0.005 ? "text-emerald-700" : v < -0.005 ? "text-rose-700" : "text-slate-600";
    return `<span class="${c} ${big ? "text-xl font-bold" : "font-semibold"} tabular-nums">${v > 0 ? "+" : ""}${v.toFixed(2)}%</span>`;
  };
  const monthName = (ym) => {
    const [y, m] = ym.split("-");
    return new Date(Date.UTC(+y, +m - 1, 1)).toLocaleDateString("en-IN", { month: "short", year: "numeric", timeZone: "UTC" });
  };

  const pills = months.map((m) => `
    <button type="button" data-paper-month="${m}" class="px-2.5 py-1.5 text-xs font-semibold rounded-lg ring-1 transition ${m === sel ? "bg-slate-900 text-white ring-slate-900" : "bg-white text-slate-600 ring-slate-200 hover:bg-slate-50"}">${monthName(m)}</button>`).join("");

  const kpi = (label, val, sub) => `
    <div class="px-4 py-3">
      <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">${label}</div>
      <div class="mt-0.5">${val}</div>
      ${sub ? `<div class="text-[10px] text-slate-400 mt-0.5">${sub}</div>` : ""}
    </div>`;

  const rows = t.positions.map((p) => `
    <tr class="border-t border-slate-100 ${p.status === "stopped" ? "bg-rose-50/40" : ""}">
      <td class="px-3 py-2 font-semibold text-slate-900">${escapeHtml(p.name)}<div class="text-[10px] font-mono text-slate-400">${escapeHtml(p.ticker)}</div></td>
      <td class="px-3 py-2 text-right tabular-nums text-slate-500">${inr(p.entry)}</td>
      <td class="px-3 py-2 text-right tabular-nums font-semibold">${inr(p.current)}</td>
      <td class="px-3 py-2 text-right">${pct(p.ret)}</td>
      <td class="px-3 py-2 text-right tabular-nums text-rose-700">${inr(p.stop)}</td>
      <td class="px-3 py-2 text-right">
        ${p.status === "stopped"
          ? `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-rose-100 text-rose-700 ring-1 ring-rose-200">Stopped ${fmtDateDMY(p.exitDate)}</span>`
          : `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-emerald-50 text-emerald-700 ring-1 ring-emerald-100">Open</span>`}
      </td>
      <td class="px-3 py-2 text-right tabular-nums text-slate-500">${(p.weight * 100).toFixed(1)}%</td>
    </tr>`).join("");

  const oneDay = t.days < 2;
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <div class="px-4 py-3 border-b border-slate-100 flex flex-wrap items-center justify-between gap-3">
        <div>
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Paper results · ${monthName(sel)}</div>
          <div class="text-xs text-slate-500 mt-0.5">Held from ${fmtDateDMY(t.anchorDate)} · ${t.days} trading day${t.days === 1 ? "" : "s"} recorded${oneDay ? " · results start building tomorrow" : ""}</div>
        </div>
        <div class="flex flex-wrap items-center gap-1.5">${pills}</div>
      </div>

      <div class="grid grid-cols-2 sm:grid-cols-4 divide-x divide-y sm:divide-y-0 divide-slate-100 border-b border-slate-100">
        ${kpi("AI basket", pct(t.basketRet, true), "weighted, stops applied")}
        ${kpi("Smallcap 250", pct(t.benchRet, true), "same window")}
        ${kpi("Alpha", pct(t.alpha, true), "basket − benchmark")}
        ${kpi("Stops hit", `<span class="text-xl font-bold tabular-nums ${t.stops ? "text-rose-700" : "text-slate-700"}">${t.stops} <span class="text-sm font-normal text-slate-400">of ${t.positions.length}</span></span>`, `${t.winners} in profit`)}
      </div>

      ${paperChart(t)}

      <div class="overflow-x-auto border-t border-slate-100">
        <table class="w-full text-sm">
          <thead class="bg-slate-50/70 text-[10px] uppercase tracking-wider text-slate-500">
            <tr><th class="px-3 py-2 text-left">Stock</th><th class="px-3 py-2 text-right">Entry</th><th class="px-3 py-2 text-right">Now</th>
            <th class="px-3 py-2 text-right">Return</th><th class="px-3 py-2 text-right">Stop</th><th class="px-3 py-2 text-right">Status</th><th class="px-3 py-2 text-right">Weight</th></tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>

      <div class="px-4 py-3 bg-slate-50/60 border-t border-slate-100 flex flex-wrap gap-x-6 gap-y-1 text-xs text-slate-600">
        <span>Best day <span class="font-bold text-emerald-700 tabular-nums">${t.peak > 0 ? "+" : ""}${t.peak.toFixed(2)}%</span></span>
        <span>Worst day <span class="font-bold text-rose-700 tabular-nums">${t.trough.toFixed(2)}%</span></span>
        <span class="text-slate-400">Stops are checked on closing prices, so a real intraday stop would trigger earlier and slightly worse.</span>
      </div>
    </div>`;
}

// Strategy-tab sub-tab bar. Same pattern as the top-level tabs the founder
// likes on the AI Basket page — one click, no scroll.
function renderStrategySubNav(sub) {
  const tabs = [
    { k: "overview",  icon: "📈", label: "Overview" },
    { k: "plan",      icon: "🧾", label: "Trade Plan" },
    { k: "compare",   icon: "⚖️", label: "Compare" },
    { k: "accuracy",  icon: "🎯", label: "Accuracy" },
    { k: "balancing", icon: "🧭", label: "Balancing" },
  ];
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-1 flex flex-wrap gap-1">
      ${tabs.map((t) => {
        const on = sub === t.k;
        return `<button type="button" data-strategy-subtab="${t.k}" class="flex-1 sm:flex-initial px-3 py-2 rounded-xl text-sm font-semibold transition ${on ? "bg-indigo-600 text-white shadow-sm" : "text-slate-600 hover:text-slate-900 hover:bg-slate-50"}"><span class="mr-1">${t.icon}</span>${t.label}</button>`;
      }).join("")}
    </div>`;
}

// Body for the active sub-tab. Overview = the main view (chart + alpha KPIs
// + AI/Manual picks with targets); the rest are the founder's named sub-tabs.
// Everything from the old single-scroll layout still lives here — just gated.
function renderStrategySubPanel(view, sub, mode, anchorDate) {
  switch (sub) {
    case "plan":
      return renderTradePlan();
    case "compare":
      return renderStrategyCompare();
    case "accuracy":
      return `<div class="space-y-4">${renderActiveOverallHitsSplit(view)}${renderActivePickRowsSplit(view)}</div>`;
    case "balancing":
      return renderStrategyBalancing(view);
    case "capital":
      return `<div class="space-y-4">${renderSimPanel(view)}${renderActiveBetaCaveat(view, anchorDate)}</div>`;
    case "overview":
    default:
      return `<div class="space-y-4">${renderActiveCumulativeChart(view)}${renderStrategyKpis(view)}${renderActiveSegmentedBaskets(view, mode)}</div>`;
  }
}

// ============================================================
// COMPARE — five strategies, one table, no favourites.
//
// The founder's question is "which of these actually makes the most money
// for the least risk", and the honest answer needs three views: month by
// month (is it consistent, or was it one lucky month), a summary of every
// risk and cost measure, and the trades themselves.
//
// Numbers come from data/strategy-backtest.json, precomputed by
// screener-test/backtest-strategies.mjs. Replaying 504 days x 620 tickers
// x 5 strategies in the browser took seconds and blocked the first paint.
// ============================================================

// Picking a #1 has to survive the browser it was picked in — otherwise the
// two founders each see their own answer and the comparison is pointless.
// A static site cannot write to its own repo, so the honest version is:
// apply it locally straight away, and hand over the exact file to commit.
function openPrimaryCommitModal(id) {
  const def = strat5.strategyById(id);
  const json = strat5.strategiesFileJson();
  const overlay = $("#modal-overlay");
  const content = $("#modal-content");
  if (!overlay || !content) return;
  content.innerHTML = `
    <div class="p-6">
      <div class="flex items-start gap-3">
        <span class="text-2xl">★</span>
        <div class="flex-1 min-w-0">
          <h3 class="font-display font-bold text-slate-900 text-lg">${escapeHtml(def?.name || id)} is now your #1</h3>
          <p class="text-sm text-slate-600 mt-1">Applied here already — the Trade Plan follows it from now on. To make it stick for everyone, commit this file.</p>
        </div>
        <button type="button" data-close-modal class="text-slate-400 hover:text-slate-600 text-xl leading-none">×</button>
      </div>
      <div class="mt-4">
        <div class="flex items-center justify-between mb-1.5">
          <code class="text-[11px] font-semibold text-slate-500">public/data/strategies.json</code>
          <button type="button" id="copy-strategies-json" class="px-3 py-1.5 rounded-lg bg-indigo-600 text-white text-xs font-semibold hover:bg-indigo-700">Copy file</button>
        </div>
        <pre class="text-[10px] bg-slate-900 text-slate-200 rounded-xl p-3 max-h-56 overflow-auto scrollbar-thin">${escapeHtml(json)}</pre>
      </div>
    </div>`;
  overlay.classList.add("is-open");
  overlay.classList.remove("hidden");
  const close = () => { overlay.classList.remove("is-open"); overlay.classList.add("hidden"); };
  content.querySelector("[data-close-modal]")?.addEventListener("click", close);
  content.querySelector("#copy-strategies-json")?.addEventListener("click", (e) => {
    navigator.clipboard?.writeText(json);
    e.target.textContent = "Copied ✓";
  });
}

// ============================================================
// LIVE STRATEGY TRACKING
//
// The backtest answers "what would these five have done over the last two
// years". It cannot answer "which one is actually winning", because it
// never happened. This does: it replays all five over the REAL snapshot
// trail, with the closes actually recorded each morning and the charges
// actually payable, and it extends by one point every trading day without
// anyone running anything -- the snapshots already refresh at 06:23 IST.
//
// It will look thin for the first few weeks. That is honest: a comparison
// between five strategies picking from one ranking needs a couple of
// months before sector caps and basket width separate them.
// ============================================================

// ---- Live quotes, fetched by the browser on every open ----------------
//
// live-prices.json is written by a workflow and committed, so the page was
// only ever as current as the last successful run -- and when that run was
// cancelled on 3 Aug the dashboard served Friday's closes through a live
// session without saying so. The quote API allows any origin
// (access-control-allow-origin: *), so the page can simply ask for itself.
//
// The committed file stays as the fallback: it covers a blocked network, a
// rate limit, or the API hanging as it briefly did. So the worst case is
// what we had before, and the normal case is quotes seconds old.
const QUOTE_API = "https://fastapi.muns.io/stock-data";
const QUOTE_TTL_MS = 60_000;   // don't refetch on every sub-tab click

function parseQuoteLine(str) {
  const kv = {};
  for (const part of String(str).split(",")) {
    const i = part.indexOf("=");
    if (i > 0) kv[part.slice(0, i).trim()] = part.slice(i + 1).trim();
  }
  const num = (v) => { const n = Number(String(v).replace(/[^0-9.\-]/g, "")); return Number.isFinite(n) ? n : null; };
  const range = (v) => {
    const m = String(v || "").split("-").map(num);
    return m.length === 2 && m[0] != null && m[1] != null ? { lo: Math.min(...m), hi: Math.max(...m) } : null;
  };
  const current = num(kv["Current Price"]);
  if (current == null) return null;
  const day = range(kv["Day Range"]), wk = range(kv["52-Week Range"]);
  return {
    current,
    open: num(kv["Opening Price"]),
    prevClose: num(kv["Previous Close"]),
    dayHigh: day?.hi ?? null, dayLow: day?.lo ?? null,
    week52High: wk?.hi ?? null, week52Low: wk?.lo ?? null,
    ma50: num(kv["50-Day Moving Average"]), ma200: num(kv["200-Day Moving Average"]),
  };
}

async function fetchQuoteDirect(ticker) {
  try {
    const r = await fetch(QUOTE_API, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ticker_symbol: ticker, type: "stockquote", country: "india" }),
      signal: AbortSignal.timeout(7000),
    });
    if (!r.ok) return null;
    const body = await r.json();
    return typeof body === "string" ? parseQuoteLine(body) : null;
  } catch { return null; }
}

// Refresh quotes for the basket in place. Anything that fails keeps whatever
// the committed file had, so a partial outage degrades one row rather than
// blanking the basket.
async function refreshLiveQuotes(tickers) {
  const h = state.cache.history;
  if (!h || !tickers?.length) return;
  const now = Date.now();
  if (state.cache.quotesFetchedAt && now - state.cache.quotesFetchedAt < QUOTE_TTL_MS) return false;
  state.cache.quotesFetchedAt = now;

  const uniq = [...new Set(tickers.filter(Boolean))].slice(0, 20);
  const results = await Promise.all(uniq.map(async (t) => [t, await fetchQuoteDirect(t)]));
  let ok = 0;
  for (const [t, q] of results) if (q) { h.livePrices[t] = q; ok++; }
  if (!ok) return false;                 // API unreachable — committed file stands
  h.livePricesGeneratedAt = new Date().toISOString();
  h.livePricesSource = `browser · ${ok}/${uniq.length} quotes`;
  // Display prices follow the fresh quotes too, not just the hit detection.
  for (const [t, q] of results) if (q?.current != null) h.todayClose[t] = q.current;
  return true;
}

// Is live-prices.json actually from today? The file is committed by a
// workflow, so a stale one looks exactly like a fresh one unless the
// timestamp is checked. Returns null when it is not today's, so callers
// fall back to closes rather than quietly plotting two-day-old prices as
// "live".
// Which trading session do the quotes we are holding belong to? NOT the
// wall-clock date: at 00:53 IST a quote is Monday's closing price, and
// stamping it "Tuesday" invented a data point for a day that had not
// traded -- which is exactly what the chart did. Before the 09:15 open the
// session is the previous trading day, and either way we walk back to a
// date the exchange actually opened.
function quoteSessionDate() {
  const istMs = Date.now() + 5.5 * 3600 * 1000;
  const ist = new Date(istMs);
  const minutes = ist.getUTCHours() * 60 + ist.getUTCMinutes();
  let d = ist.toISOString().slice(0, 10);
  if (minutes < 9 * 60 + 15) d = shiftDateStr(d, -1);   // pre-open: still yesterday's session
  const cal = state.cache.history?.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes;
  if (cal) {
    for (let i = 0; i < 10 && cal[d] == null; i++) d = shiftDateStr(d, -1);
    if (cal[d] == null) return null;
  }
  return d;
}

// Quotes usable as the current session's mark, or null. Stale quotes --
// generated before the session they would be plotted on -- are refused, so
// a feed that stopped updating can never masquerade as live.
function liveQuotesToday() {
  const lp = state.cache.history?.livePrices;
  const gen = state.cache.history?.livePricesGeneratedAt;
  if (!lp || !gen) return null;
  const istDay = (d) => new Date(new Date(d).getTime() + 5.5 * 3600 * 1000).toISOString().slice(0, 10);
  const session = quoteSessionDate();
  if (!session || istDay(gen) < session) return null;
  const prices = {};
  for (const [t, q] of Object.entries(lp)) if (q?.current != null) prices[t] = q.current;
  return Object.keys(prices).length ? { date: session, prices, asOf: gen } : null;
}

// Age of a timestamp in hours, for the freshness strip.
function hoursSince(iso) {
  if (!iso) return null;
  return (Date.now() - new Date(iso).getTime()) / 3600000;
}

// One honest line per feed: what it is, when it last landed, and whether
// that is a problem. The dashboard used to show +0.00% across the board
// when a scraper had not run, which reads as "nothing moved" rather than
// "nothing arrived" -- the two look identical and mean opposite things.
function renderDataFreshness() {
  const h = state.cache.history;
  if (!h) return "";
  const snaps = h.snapshots || [];
  const lastSnap = snaps.length ? snaps[snaps.length - 1].date : null;
  const cal = h.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || {};
  const lastClose = Object.keys(cal).sort().pop() || null;
  const lpAge = hoursSince(h.livePricesGeneratedAt);
  const mcAge = hoursSince(state.cache.marketContext?.generated_at);
  const fresh = liveQuotesToday();

  const liveFrom = state.cache.liveFrom;
  const recon = liveFrom && snaps.length && snaps[0].date < liveFrom;
  const rows = [
    { label: "Daily snapshot", value: lastSnap ? fmtDateDMY(lastSnap) : "none",
      bad: !lastSnap, warn: lastSnap && lastSnap < istTodayDate() },
    { label: "Prices through", value: lastClose ? fmtDateDMY(lastClose) : "—", bad: !lastClose },
    { label: "Live quotes", value: lpAge == null ? "never"
        : fresh ? (lpAge < 0.05 ? "just now" : `${Math.round(lpAge * 60)} min ago`)
        : `${Math.floor(lpAge / 24)}d old`,
      bad: !fresh, warn: false, note: h.livePricesSource || null },
    { label: "Market context", value: mcAge == null ? "never" : mcAge < 24 ? `${Math.round(mcAge)}h ago` : `${Math.floor(mcAge / 24)}d old`,
      warn: mcAge != null && mcAge > 24, bad: mcAge == null },
  ];
  const anyBad = rows.some((r) => r.bad);
  const tone = anyBad ? "amber" : "slate";
  return `
    <div class="bg-${tone}-50 ring-1 ring-${tone}-200 rounded-2xl px-4 py-2.5 flex flex-wrap items-center gap-x-5 gap-y-1">
      <span class="text-[10px] font-bold uppercase tracking-wider text-${tone}-700">Data freshness</span>
      ${rows.map((r) => `<span class="text-[11px] text-slate-600"${r.note ? ` title="${escapeHtml(r.note)}"` : ""}>${r.label}
        <b class="${r.bad ? "text-amber-800" : r.warn ? "text-slate-700" : "text-emerald-700"}">${r.value}</b></span>`).join("")}
      ${recon ? `<span class="text-[11px] text-slate-500">History before <b class="text-slate-700">${fmtDateDMY(liveFrom)}</b> is rebuilt from past prices, not a live record</span>` : ""}
      ${anyBad ? `<span class="text-[11px] text-amber-900 ml-auto">A feed hasn't landed — figures below are the last good data, not today's.</span>` : ""}
    </div>`;
}

function liveTradingDays(snapshots) {
  // The benchmark is the trading calendar. Snapshots are written every
  // calendar day, so without this a Saturday carrying Friday's closes
  // becomes a data point.
  const cal = state.cache.history?.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes;
  if (!cal) return snapshots;
  return snapshots.filter((s) => cal[s.date] != null);
}

// One snapshot -> ranked candidates in the shape the selection helpers want.
function liveCandidates(snap) {
  const adtvBy = state.cache.planAdtv || {};
  return snap.stocks
    .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed && typeof s.close === "number")
    .sort((a, b) => b.composite - a.composite)
    .map((s) => {
      const tv = s.techVals || {};
      return {
        ticker: s.ticker, name: s.name || s.ticker, sector: s.sector || "—", score: s.composite,
        close: s.close,
        ind: { rsi: tv.rsi, atr: tv.atr, d52: tv.d52, adtv: adtvBy[s.ticker] ?? null,
               aboveMa50: tv.e50, aboveMa200: tv.d200, macdPositive: tv.macd },
      };
    });
}

// Replay one strategy over the real snapshots. Rebalances on the first
// trading day of each month -- the product is a monthly basket, so the
// live track has to be the monthly basket, not a daily re-pick.
function runLiveStrategy(days, strategy, chg, capital) {
  const want = strategy.basketSize || 7;
  const stopAtr = strategy.stopAtr ?? 2.5;
  const trailAtr = strategy.trailAtr ?? 3;
  let cash = capital, charges = 0, lastMonth = null, daysInvested = 0;
  const holdings = new Map();
  const trades = [];
  const curve = [];
  const baseDefs = strategy.consensus ? strat5.allStrategies().filter((s) => !s.consensus) : null;

  days.forEach((snap, i) => {
    const closeBy = new Map();
    for (const s of snap.stocks) if (typeof s.close === "number") closeBy.set(s.ticker, s.close);
    const ym = snap.date.slice(0, 7);
    const isRebal = lastMonth === null || ym !== lastMonth;

    let sel = null;
    if (isRebal) {
      lastMonth = ym;
      const ranked = liveCandidates(snap);
      sel = strategy.consensus && baseDefs?.length
        ? strat5.consensusPicks(baseDefs.map((d) => strat5.selectFromRanked(ranked, d, {})), strategy)
        : strat5.selectFromRanked(ranked, strategy, {});
    }
    const inBasket = sel ? new Set(sel.core.map((p) => p.ticker)) : null;

    // Exits first, so freed cash can fund the same day's entries.
    for (const [t, pos] of [...holdings]) {
      const px = closeBy.get(t);
      if (px == null) continue;
      if (px > pos.peak) pos.peak = px;
      let reason = null;
      if (px <= pos.entry * (1 - (stopAtr * pos.atr) / 100)) reason = "SL";
      else if (strategy.trailStop && pos.peak >= pos.entry * (1 + ((strategy.trailArmAtr ?? 1.5) * pos.atr) / 100)
               && px <= pos.peak * (1 - (trailAtr * pos.atr) / 100)) reason = "TRAIL";
      else if (isRebal && inBasket && !inBasket.has(t)) reason = "REBAL";
      if (!reason) continue;
      const gross = pos.units * px, fee = chg.sell(gross);
      cash += gross - fee; charges += fee;
      trades.push({ ticker: t, entryDate: pos.entryDate, exitDate: snap.date, entry: pos.entry, exit: px,
                    retPct: (px / pos.entry - 1) * 100, days: i - pos.dayIdx, reason });
      holdings.delete(t);
    }

    if (isRebal && sel) {
      let nav = cash;
      for (const [t, p] of holdings) nav += p.units * (closeBy.get(t) ?? p.entry);
      const weightOf = strat5.invAtrWeights(sel.core, strategy);
      for (const pick of sel.core) {
        if (holdings.size >= want || holdings.has(pick.ticker)) continue;
        const px = closeBy.get(pick.ticker);
        if (px == null || px <= 0) continue;
        const spend = Math.min(nav * weightOf(pick), Math.max(0, cash) * 0.98);
        const units = Math.floor(spend / px);
        if (units < 1) continue;
        const cost = units * px, fee = chg.buy(cost);
        if (cost + fee > cash) continue;
        cash -= cost + fee; charges += fee;
        holdings.set(pick.ticker, { units, entry: px, entryDate: snap.date, dayIdx: i, peak: px,
                                    atr: pick.ind?.atr ?? 3.5 });
      }
    }

    let value = cash;
    for (const [t, p] of holdings) value += p.units * (closeBy.get(t) ?? p.entry);
    if (holdings.size) daysInvested++;
    curve.push({ date: snap.date, retPct: (value / capital - 1) * 100, holdings: holdings.size });
  });

  // Today's intraday mark. Snapshots are written pre-market and carry the
  // PREVIOUS close, so without this the chart sits frozen all session while
  // the basket is visibly moving -- which reads as "flat", not "no data".
  const q = liveQuotesToday();
  if (q && holdings.size) {
    let value = cash;
    for (const [t, p] of holdings) value += p.units * (q.prices[t] ?? p.entry);
    const pt = { date: q.date, retPct: (value / capital - 1) * 100, holdings: holdings.size, live: true };
    if (curve.length && curve[curve.length - 1].date === q.date) curve[curve.length - 1] = pt;
    else curve.push(pt);
  }

  return { curve, trades, charges, startCapital: capital, daysInvested, strategy, liveMark: !!q };
}

// All five over the real trail, cached per snapshot count so switching
// sub-tabs does not replay everything.
function liveStrategyRuns() {
  const snaps = state.cache.history?.snapshots || [];
  const days = liveTradingDays(snaps);
  const defs = strat5.allStrategies();
  if (!days.length || !defs.length) return null;
  const key = `${days.length}:${days[days.length - 1]?.date}:${simPrefs.capital}`;
  if (state.cache.liveRunsKey === key && state.cache.liveRuns) return state.cache.liveRuns;

  const chg = makeCharger(simPrefs);
  const bench = buildNiftyCurve(days.map((d) => d.date), (date) => {
    const cl = state.cache.history?.benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes;
    return cl ? cl[date] ?? null : null;
  });
  const runs = defs.map((d) => {
    const result = runLiveStrategy(days, d, chg, simPrefs.capital);
    return { def: d, id: d.id, name: d.name, result, curve: result.curve,
             metrics: strat5.metrics(result, bench), trades: result.trades };
  });
  const curveDays = Math.max(...runs.map((r) => r.curve.length), 0);
  const out = { runs, bench, days, curveDays, start: days[0]?.date, end: days[days.length - 1]?.date };
  state.cache.liveRuns = out; state.cache.liveRunsKey = key;
  return out;
}

const COMPARE_VIEWS = ["monthly", "summary", "trades"];
const COMPARE_SOURCES = ["live", "backtest"];
function compareSource() {
  return COMPARE_SOURCES.includes(state.compareSource) ? state.compareSource : "live";
}

function compareView() {
  return COMPARE_VIEWS.includes(state.compareView) ? state.compareView : "summary";
}

function renderStrategyCompare() {
  const bt = state.cache.strategyBacktest;
  const defs = strat5.allStrategies();
  if (!defs.length) {
    return `<div class="bg-white rounded-2xl ring-1 ring-slate-100 p-8 text-center text-sm text-slate-500">Strategy definitions not loaded.</div>`;
  }
  const src = compareSource();
  const view = compareView();
  const primary = strat5.primaryId();
  const live = src === "live" ? liveStrategyRuns() : null;

  const srcBtn = (k, label, sub) => `
    <button type="button" data-compare-source="${k}" class="px-3 py-1.5 rounded-lg text-left transition ${src === k ? "bg-white shadow-sm ring-1 ring-slate-200" : "hover:bg-white/60"}">
      <div class="text-xs font-bold ${src === k ? "text-slate-900" : "text-slate-500"}">${label}</div>
      <div class="text-[9px] ${src === k ? "text-slate-500" : "text-slate-400"}">${sub}</div>
    </button>`;
  const tab = (k, label) => `<button type="button" data-compare-view="${k}" class="px-3 py-1.5 rounded-lg text-xs font-bold transition ${view === k ? "bg-indigo-600 text-white shadow-sm" : "text-slate-500 hover:text-slate-700"}">${label}</button>`;

  const range = src === "live"
    ? (live?.start ? `${fmtDateDMY(live.start)} → ${fmtDateDMY(live.end)} · ${live.days.length} trading day${live.days.length === 1 ? "" : "s"} of real prices` : "waiting for the first trading day")
    : bt?.window ? `${fmtDateDMY(bt.window.start)} → ${fmtDateDMY(bt.window.end)} · ${(bt.window.days / 252).toFixed(1)} years, simulated` : "not built";

  const header = `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4 flex flex-wrap items-center justify-between gap-3">
      <div class="min-w-0">
        <h3 class="font-display font-bold text-slate-900 text-base">Which strategy wins?</h3>
        <div class="text-xs text-slate-500 mt-0.5">${range} · net of charges at ${inr(src === "live" ? simPrefs.capital : (bt?.capital ?? simPrefs.capital))}</div>
      </div>
      <div class="flex items-center gap-2 flex-wrap">
        <div class="inline-flex items-center gap-0.5 bg-slate-100 rounded-lg p-1">
          ${srcBtn("live", "Live", "real money, grows daily")}
          ${srcBtn("backtest", "Backtest", "2 years, simulated")}
        </div>
        <div class="inline-flex items-center gap-0.5 bg-slate-100 rounded-lg p-0.5">${tab("summary", "Summary")}${tab("monthly", "Month by month")}${tab("trades", "Trades")}</div>
      </div>
    </div>`;

  // ---- live branch --------------------------------------------------
  if (src === "live") {
    if (!live || !live.days.length) {
      return `<div class="space-y-3">${header}
        <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-8 text-center">
          <div class="text-sm text-slate-600">No trading day has closed since tracking began.</div>
          <div class="text-xs text-slate-400 mt-1">This fills in by itself — one point per trading day, starting with the next market close.</div>
          <button type="button" data-compare-source="backtest" class="mt-4 px-4 py-2 rounded-xl bg-indigo-600 text-white text-sm font-semibold hover:bg-indigo-700">See the 2-year backtest instead →</button>
        </div></div>`;
    }
    const rows = live.runs.filter((r) => r.metrics).map((r) => ({ def: r.def, metrics: r.metrics, trades: r.trades }));
    // A chart needs two points to be a line; a table of returns is useful
    // from the first day. Gating both on two days hid today's numbers for
    // no reason.
    const enough = live.curveDays >= 2;
    const chart = enough ? renderMultiCurveChart(
      [...live.runs.map((r) => ({ label: r.def.name, color: r.def.color, curve: r.curve })),
       { label: "Smallcap 250", color: "#94a3b8", curve: live.bench, dash: "5 4" }],
      "All five, side by side", "Real prices from the daily snapshots. One new point every trading day.") : "";
    const lf = state.cache.liveFrom;
    const reconNote = lf && live.days.length && live.days[0].date < lf ? `
      <div class="bg-slate-50 ring-1 ring-slate-200 rounded-2xl px-4 py-2.5 text-[12px] text-slate-600 leading-snug">
        Everything before <b>${fmtDateDMY(lf)}</b> is <b>rebuilt</b> from real past prices and real technical scores — but the fundamentals 30% is held at today's value, because no historical record of it exists. Read it as a well-grounded rehearsal, not as picks we actually made. From ${fmtDateDMY(lf)} onward it is a live record.
      </div>` : "";
    const young = `
      <div class="bg-sky-50 ring-1 ring-sky-200 rounded-2xl px-4 py-3 text-[13px] text-sky-900 leading-snug">
        <b>${live.days.length} trading day${live.days.length === 1 ? "" : "s"} in.</b>
        All five pick from the same ranking, so early on their lines sit almost on top of each other — the sector cap and the wider basket only start to separate them over a couple of months.
        This chart extends itself every trading day; nothing needs running.
      </div>`;
    const body =
      view === "monthly" ? renderCompareMonthly(rows) :
      view === "trades"  ? renderCompareTrades(rows) :
      renderCompareSummary(rows, { capital: simPrefs.capital }, primary);
    return `<div class="space-y-3">${header}${reconNote}${live.days.length < 25 ? young : ""}${chart}${body}</div>`;
  }

  // ---- backtest branch ----------------------------------------------
  if (!bt?.strategies?.length) {
    return `<div class="space-y-3">${header}
      <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-8 text-center text-sm text-slate-500">
        Backtest not built yet — run <code class="bg-slate-100 px-1 rounded text-[11px]">node screener-test/backtest-strategies.mjs</code>.
      </div></div>`;
  }
  const rows = defs.map((d) => ({ def: d, ...(bt.strategies.find((x) => x.id === d.id) || {}) })).filter((r) => r.metrics);
  const btChart = renderMultiCurveChart(
    [...bt.strategies.map((s) => {
      const def = defs.find((d) => d.id === s.id);
      return { label: s.name, color: def?.color || "#94a3b8", curve: (s.curve || []).map((p) => ({ date: p.d, retPct: p.r })) };
    }), { label: "Universe", color: "#94a3b8", curve: (bt.universeCurve || []).map((p) => ({ date: p.d, retPct: p.r })), dash: "5 4" }],
    "All five over two years", "Simulated on real past prices. A rehearsal, not a record.");
  const body = view === "monthly" ? renderCompareMonthly(rows)
             : view === "trades"  ? renderCompareTrades(rows)
             : renderCompareSummary(rows, bt, primary);
  const caveats = `
    <details class="bg-white rounded-2xl ring-1 ring-slate-100 p-4">
      <summary class="cursor-pointer text-xs font-semibold text-slate-600 select-none">What these numbers are not ▾</summary>
      <ul class="mt-2 space-y-1 text-[11px] text-slate-500 leading-snug list-disc pl-4">
        ${(bt.caveats || []).map((c) => `<li>${escapeHtml(c)}</li>`).join("")}
        <li>A backtest is a rehearsal, not a record. Switch to <b>Live</b> for the only evidence that counts.</li>
      </ul>
    </details>`;
  return `<div class="space-y-3">${header}${btChart}${body}${caveats}</div>`;
}

const cmpPct = (v, d = 1) => v == null ? `<span class="text-slate-300">—</span>`
  : `<span class="tabular-nums font-semibold ${v >= 0 ? "text-emerald-700" : "text-rose-700"}">${v >= 0 ? "+" : ""}${v.toFixed(d)}%</span>`;

function renderCompareSummary(rows, bt, primary) {
  const best = (key, dir = 1) => {
    let win = null;
    for (const r of rows) { const v = r.metrics?.[key]; if (v == null) continue; if (win == null || v * dir > win * dir) win = v; }
    return win;
  };
  const bestCagr = best("cagr"), bestDD = best("maxDD"), bestRisk = best("riskAdj");
  const mark = (v, b) => v != null && b != null && Math.abs(v - b) < 1e-9 ? " ring-1 ring-emerald-300 bg-emerald-50/60 rounded" : "";

  const head = ["Strategy", "Return", "A year", "Worst fall", "Swing", "Return ÷ risk", "Best month", "Worst month", "Months up", "Win rate", "Trades", "Charges"];
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <div class="overflow-x-auto">
        <table class="w-full text-sm">
          <thead class="bg-slate-50/70 text-[10px] uppercase tracking-wider text-slate-500">
            <tr>${head.map((h, i) => `<th class="px-3 py-2 ${i === 0 ? "text-left" : "text-right"}">${h}</th>`).join("")}</tr>
          </thead>
          <tbody>
            ${rows.map((r) => {
              const m = r.metrics;
              return `
              <tr class="border-t border-slate-100 hover:bg-slate-50">
                <td class="px-3 py-2.5">
                  <div class="flex items-center gap-2">
                    <span class="w-2 h-2 rounded-full flex-shrink-0" style="background:${r.def.color}"></span>
                    <span class="font-semibold text-slate-900">${escapeHtml(r.def.name)}</span>
                    ${r.def.id === primary ? `<span class="text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 ring-1 ring-amber-200">★ #1</span>` : ""}
                    ${r.def.isOriginal ? `<span class="text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 ring-1 ring-slate-200">what you run today</span>` : ""}
                  </div>
                  <div class="text-[10px] text-slate-400 mt-0.5">${escapeHtml(r.def.tagline)}</div>
                </td>
                <td class="px-3 py-2.5 text-right">${cmpPct(m.netReturn)}</td>
                <td class="px-3 py-2.5 text-right${m.days >= 252 ? mark(m.cagr, bestCagr) : ""}">${m.days >= 252 ? cmpPct(m.cagr) : `<span class="text-slate-300" title="Needs a full year before annualising means anything">—</span>`}</td>
                <td class="px-3 py-2.5 text-right${mark(m.maxDD, bestDD)}"><span class="tabular-nums font-semibold text-rose-700">${m.maxDD.toFixed(1)}%</span></td>
                <td class="px-3 py-2.5 text-right"><span class="tabular-nums text-slate-500">${m.vol.toFixed(1)}%</span></td>
                <td class="px-3 py-2.5 text-right${mark(m.riskAdj, bestRisk)}"><span class="tabular-nums font-bold text-slate-800">${m.riskAdj == null ? "—" : m.riskAdj.toFixed(2)}</span></td>
                <td class="px-3 py-2.5 text-right">${cmpPct(m.bestMonth?.retPct)}</td>
                <td class="px-3 py-2.5 text-right">${cmpPct(m.worstMonth?.retPct)}</td>
                <td class="px-3 py-2.5 text-right"><span class="tabular-nums text-slate-600">${m.monthlyWinRate == null ? "—" : m.monthlyWinRate.toFixed(0) + "%"}</span></td>
                <td class="px-3 py-2.5 text-right"><span class="tabular-nums text-slate-600">${m.winRate == null ? "—" : m.winRate.toFixed(0) + "%"}</span></td>
                <td class="px-3 py-2.5 text-right"><span class="tabular-nums text-slate-500">${m.trades}</span></td>
                <td class="px-3 py-2.5 text-right"><span class="tabular-nums text-slate-500">${m.chargesPct.toFixed(1)}%</span></td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>
      <div class="px-4 py-3 bg-slate-50/60 border-t border-slate-100">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-2">Make one your #1</div>
        <div class="flex flex-wrap gap-1.5">
          ${rows.map((r) => `<button type="button" data-set-primary="${r.def.id}"
            class="px-3 py-1.5 rounded-lg text-xs font-semibold ring-1 transition ${r.def.id === primary ? "bg-amber-500 text-white ring-amber-500" : "bg-white text-slate-600 ring-slate-200 hover:bg-amber-50 hover:ring-amber-200"}">
            ${r.def.id === primary ? "★ " : ""}${escapeHtml(r.def.name)}</button>`).join("")}
        </div>
        <div class="text-[10px] text-slate-400 mt-2 leading-snug">Your #1 drives the Trade Plan. The dashboard is a static site, so the button cannot commit for you — it applies here immediately and hands you the updated <code class="bg-slate-100 px-1 rounded">strategies.json</code> to commit, so your co-founder sees the same pick.</div>
      </div>
      <div class="px-4 py-3 border-t border-slate-100 text-[11px] text-slate-500 leading-snug">
        ${rows.some((r) => r.metrics.days < 252) ? `<b>A year</b> is blank until there are twelve months to annualise from — stretching a few months into a yearly figure produces a number that looks impressive and means nothing. ` : ""}<b>Return ÷ risk</b> is the return divided by how much the value swings — the closest thing here to "most money for least worry". <b>Charges</b> is what the broker took, as a share of starting capital.
      </div>
    </div>`;
}

function renderCompareMonthly(rows) {
  // Union of every month any strategy traded, so a strategy that sat in
  // cash shows a gap rather than silently shifting its neighbours' columns.
  const months = [...new Set(rows.flatMap((r) => (r.metrics.months || []).map((m) => m.month)))].sort();
  const cellFor = (v) => {
    if (v == null) return `<td class="px-2 py-1.5 text-right text-slate-300">—</td>`;
    const a = Math.min(1, Math.abs(v) / 8);
    const bg = v >= 0 ? `rgba(16,185,129,${0.08 + a * 0.32})` : `rgba(244,63,94,${0.08 + a * 0.32})`;
    return `<td class="px-2 py-1.5 text-right tabular-nums text-[11px] font-semibold ${v >= 0 ? "text-emerald-900" : "text-rose-900"}" style="background:${bg}">${v >= 0 ? "+" : ""}${v.toFixed(1)}</td>`;
  };
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 overflow-hidden">
      <div class="px-4 py-3 border-b border-slate-100">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Month by month · % net</div>
        <div class="text-xs text-slate-500 mt-0.5">Consistency matters more than any single month. Look for a row with few deep red cells.</div>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-sm">
          <thead class="bg-slate-50/70 text-[9px] uppercase tracking-wider text-slate-500">
            <tr><th class="px-3 py-2 text-left sticky left-0 bg-slate-50">Strategy</th>
            ${months.map((m) => `<th class="px-2 py-2 text-right whitespace-nowrap">${m.slice(2).replace("-", "/")}</th>`).join("")}</tr>
          </thead>
          <tbody>
            ${rows.map((r) => {
              const by = new Map((r.metrics.months || []).map((m) => [m.month, m.retPct]));
              return `<tr class="border-t border-slate-100">
                <td class="px-3 py-1.5 font-semibold text-slate-800 whitespace-nowrap sticky left-0 bg-white">
                  <span class="inline-block w-2 h-2 rounded-full mr-1.5" style="background:${r.def.color}"></span>${escapeHtml(r.def.name)}
                </td>
                ${months.map((m) => cellFor(by.get(m))).join("")}
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>
    </div>`;
}

function renderCompareTrades(rows) {
  return `
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-3">
      ${rows.map((r) => {
        const m = r.metrics;
        const t = (r.trades || []).slice().sort((a, b) => b.retPct - a.retPct);
        const best = t.slice(0, 3), worst = t.slice(-3).reverse();
        const line = (x) => `<div class="flex items-center justify-between gap-2 text-[11px] py-0.5">
            <span class="font-mono text-slate-600 truncate">${escapeHtml(x.ticker)}</span>
            <span class="text-slate-400 tabular-nums">${x.days}d · ${escapeHtml(x.reason)}</span>
            <span class="tabular-nums font-bold ${x.retPct >= 0 ? "text-emerald-700" : "text-rose-700"}">${x.retPct >= 0 ? "+" : ""}${x.retPct.toFixed(1)}%</span>
          </div>`;
        return `
        <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-4">
          <div class="flex items-center gap-2 mb-2">
            <span class="w-2 h-2 rounded-full" style="background:${r.def.color}"></span>
            <span class="font-semibold text-slate-900 text-sm">${escapeHtml(r.def.name)}</span>
            <span class="text-[10px] text-slate-400 ml-auto">${m.trades} trades · avg ${m.avgHoldDays == null ? "—" : Math.round(m.avgHoldDays) + "d"} held</span>
          </div>
          <div class="grid grid-cols-3 gap-2 mb-3">
            ${["Avg win", "Avg loss", "Win ÷ loss"].map((lbl, i) => {
              const v = i === 0 ? m.avgWin : i === 1 ? m.avgLoss : m.profitFactor;
              const txt = v == null ? "—" : i === 2 ? v.toFixed(2) : `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
              const cls = i === 0 ? "text-emerald-700" : i === 1 ? "text-rose-700" : "text-slate-800";
              return `<div class="rounded-lg bg-slate-50 px-2 py-1.5">
                <div class="text-[9px] uppercase tracking-wider text-slate-500">${lbl}</div>
                <div class="text-sm font-bold tabular-nums ${cls}">${txt}</div></div>`;
            }).join("")}
          </div>
          <div class="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-1">Best</div>
          ${best.map(line).join("")}
          <div class="text-[9px] font-bold uppercase tracking-wider text-slate-400 mt-2 mb-1">Worst</div>
          ${worst.map(line).join("")}
        </div>`;
      }).join("")}
    </div>`;
}

// Sector / Industry timing sub-tab wrapper — shows a friendly empty state
// instead of a blank panel when there's no timing data for that grouping.
function renderStrategyTimingPanel(view, by) {
  const html = renderSectorTiming(view, by);
  if (html) return html;
  const label = by === "industry" ? "Industry" : "Sector";
  return `<div class="bg-white rounded-2xl ring-1 ring-slate-100 p-6 text-center text-sm text-slate-500">No ${label.toLowerCase()} timing data yet — needs a few AI picks carrying ${label.toLowerCase()} tags across this window.</div>`;
}

// Collects target/SL hits across AI + Manual baskets so the alerts bell
// and the optional banner can surface them at the top of the page.
// todayHits = hit happened on the latest snapshot date (gets the pulse).
function collectStrategyHits(view, todayDate) {
  const aiHits = (view.picks || [])
    .filter((p) => (p.status === "TARGET_HIT" || p.status === "SL_HIT") && p.hitDate)
    .map((p) => ({ ...p, basket: "AI" }));
  const manualHits = (view.manualPicks || [])
    .filter((p) => !p.notCovered && (p.status === "TARGET_HIT" || p.status === "SL_HIT") && p.hitDate)
    .map((p) => ({ ...p, basket: "Manual" }));
  const allHits = [...aiHits, ...manualHits].sort((a, b) => (b.hitDate || "").localeCompare(a.hitDate || ""));
  const todayHits = allHits.filter((h) => h.hitDate === todayDate);
  return { allHits, todayHits };
}


// Top-level Active vs Passive toggle. Sits above the cadence pills.
// Active = AI re-locks at chosen cadence; Passive = AI frozen at upload.
// Also hosts the alerts bell + dropdown — surfaces target/SL hits
// across both AI and Manual baskets without the analyst having to
// scroll through every per-pick row.
function renderStrategyModeToggle(mode, hits) {
  const pill = (k, label, sub) => {
    const isActive = mode === k;
    return `
      <button type="button" data-strategy-mode="${k}" class="flex-1 sm:flex-initial relative px-5 py-3 text-sm font-semibold transition rounded-xl ${isActive ? "bg-indigo-600 text-white shadow-sm" : "text-slate-600 hover:text-slate-900 hover:bg-slate-50"}">
        <span class="block">${label}</span>
        <span class="block text-[10px] font-normal mt-0.5 ${isActive ? "text-indigo-100" : "text-slate-400"}">${sub}</span>
      </button>`;
  };
  const modePills = SHOW_ROTATION_STRATEGIES
    ? `<div class="flex gap-1 flex-1 sm:flex-initial">
        ${pill("active", "Active strategy", "AI re-locks at cadence")}
        ${pill("passive", "Passive strategy", "AI basket frozen at month start")}
      </div>`
    : `<div class="flex items-center gap-2 flex-1">
        <span class="text-base">🎯</span>
        <div><div class="text-sm font-semibold text-slate-900">AI basket vs benchmarks</div></div>
      </div>`;
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 p-2 sm:p-3 flex flex-wrap items-center justify-between gap-2">
      ${modePills}
      <div class="flex items-center gap-2 pr-1 sm:pr-2">
        <input id="lkp-file-input" type="file" accept=".xlsx,.xls,.csv" class="hidden" />
        <button id="lkp-upload-btn" type="button" title="Upload the month's client basket (Excel / CSV)" class="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg ring-1 ring-slate-200 bg-white hover:bg-indigo-50 hover:ring-indigo-200 text-xs font-semibold text-slate-700">⬆ <span class="hidden sm:inline">Upload basket</span></button>
        ${lkpOverride() ? `<button id="lkp-reset-upload" type="button" title="Discard the uploaded basket and use the published one" class="text-[11px] font-semibold text-slate-500 hover:text-rose-600">Reset</button>` : ""}
        ${hits ? renderStrategyAlertsBell(hits) : ""}
      </div>
    </div>
  `;
}

function renderStrategyAlertsBell(hits) {
  const todayCount = hits.todayHits.length;
  const totalCount = hits.allHits.length;
  const dropdownBody = totalCount === 0
    ? `<div class="px-3 py-4 text-[11px] text-slate-500 text-center">No hits yet — once a pick crosses target or SL, it'll land here.</div>`
    : `
      <div class="sticky top-0 px-3 py-2 bg-slate-50 border-b border-slate-100 text-[11px] font-bold uppercase tracking-wider text-slate-600 flex items-center justify-between">
        <span>Recent hits · ${totalCount} total</span>
        ${todayCount > 0 ? `<span class="inline-flex items-center px-1.5 py-0 rounded bg-amber-500 text-white text-[9px] font-bold uppercase tracking-wider animate-pulse">${todayCount} today</span>` : ""}
      </div>
      <div class="py-1">
        ${hits.allHits.map((h) => `
          <div class="px-3 py-1.5 hover:bg-slate-50 text-xs ${h.hitDate === hits.allHits[0]?.hitDate && hits.todayHits.length ? "" : ""}${(hits.todayHits.find((t) => t.ticker === h.ticker && t.hitDate === h.hitDate) ? "bg-amber-50/40" : "")}">
            <div class="flex items-center justify-between gap-2">
              <span class="font-semibold text-slate-900 truncate">${escapeHtml(h.name || h.ticker)}</span>
              <span class="text-[10px] tabular-nums text-slate-400 whitespace-nowrap">${fmtDateDMY(h.hitDate)}</span>
            </div>
            <div class="text-[10px] text-slate-500 mt-0.5">
              <span class="inline-flex items-center gap-1">
                <span class="w-1.5 h-1.5 rounded-full" style="background:${h.basket === "AI" ? "#6366f1" : "#f59e0b"}"></span>
                ${escapeHtml(h.basket)}
              </span>
              ·
              <span class="${h.status === "TARGET_HIT" ? "text-emerald-700 font-bold" : "text-rose-700 font-bold"}">${h.status === "TARGET_HIT" ? "🎯 Target" : "⚠ SL"}</span>
              at ₹${formatPrice(h.exitPrice)} · ${h.daysToHit}d
              ${hits.todayHits.find((t) => t.ticker === h.ticker && t.hitDate === h.hitDate) ? `<span class="ml-1 inline-flex items-center px-1 py-0 rounded bg-amber-500 text-white text-[8px] font-bold uppercase">just hit</span>` : ""}
            </div>
          </div>`).join("")}
      </div>`;
  const badge = todayCount > 0
    ? `<span class="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-amber-500 text-white text-[9px] font-bold tabular-nums">${todayCount}</span>`
    : totalCount > 0 ? `<span class="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-slate-200 text-slate-600 text-[9px] font-bold tabular-nums">${totalCount}</span>` : "";
  return `
    <div class="relative">
      <button id="strategy-alerts-btn" type="button" class="relative inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg ring-1 ring-slate-200 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-700">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>
        <span class="hidden sm:inline">Alerts</span>
        ${badge}
      </button>
      <div id="strategy-alerts-dropdown" class="hidden absolute right-0 top-full mt-1.5 w-80 bg-white rounded-xl ring-1 ring-slate-200 shadow-2xl z-50 max-h-96 overflow-y-auto">
        ${dropdownBody}
      </div>
    </div>
  `;
}

function renderActiveCadencePills(cadence) {
  const pill = (k, label, sub) => {
    const active = cadence === k;
    return `
      <button type="button" data-cadence="${k}" class="relative px-4 py-2.5 text-sm font-semibold transition ${active ? "text-indigo-700" : "text-slate-500 hover:text-slate-900"}">
        <span>${label}</span>
        <span class="block text-[10px] font-normal text-slate-400 mt-0.5">${sub}</span>
        ${active ? `<span class="absolute bottom-0 left-2 right-2 h-0.5 bg-indigo-600 rounded-full"></span>` : ""}
      </button>`;
  };
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 px-2 sm:px-4 py-1 flex flex-wrap items-baseline gap-1 justify-between">
      <div class="flex items-baseline gap-0">
        ${pill("weekly", "Weekly", "re-lock every 7 days")}
        ${pill("monthly", "Monthly", "re-lock every 30 days")}
      </div>    </div>
  `;
}

// One command bar for the whole top of the Strategy tab: title + context on
// the left, the AI / Manual / Nifty returns in the middle, upload + alerts on
// the right. Replaces the old duplicated (mode-header + scoreboard) cards so
// the chart clears the fold. Carries the upload/alerts IDs the wiring expects.
function renderStrategyCommandBar(view, cadence, mode, hits) {
  const isPassive = mode === "passive";
  const modeLabel = isPassive ? "Passive · frozen at month start" : `${cadence} re-lock`;
  const aiDD = curveMaxDrawdown(view.equityCurve || []);
  const manualDD = view.manualFinalReturn != null ? curveMaxDrawdown(view.manualCurve || []) : null;
  const niftyDD = (view.niftyCurve && view.niftyCurve.length) ? curveMaxDrawdown(view.niftyCurve) : null;
  const nifty500DD = (view.nifty500Curve && view.nifty500Curve.length) ? curveMaxDrawdown(view.nifty500Curve) : null;
  const stat = (label, ret, dd, labelCls) => {
    if (ret == null) return "";
    const rc = ret >= 0 ? "text-emerald-600" : "text-rose-600";
    return `
      <div class="px-3 py-1 text-center leading-none">
        <div class="text-[9px] font-bold uppercase tracking-wider ${labelCls}">${label}</div>
        <div class="${rc} text-lg font-extrabold tabular-nums mt-1">${ret >= 0 ? "+" : ""}${ret.toFixed(2)}%</div>
        <div class="text-[8px] text-slate-400 tabular-nums mt-0.5" title="Worst peak-to-trough decline (max drawdown)">${dd != null ? `DD ${dd.toFixed(2)}%` : "—"}</div>
      </div>`;
  };
  const convoChip = view.manualFinalReturn != null
    ? `<span class="inline-flex items-center gap-1 text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded ${view.manualBooked ? "bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200" : "bg-slate-100 text-slate-600 ring-1 ring-slate-200"}" title="How returns are booked — flip with the Booked / If held toggle below">${view.manualBooked ? "🔒 Booked · first target / SL" : "Held · mark-to-market"}</span>`
    : "";
  return `
    <div id="strategy-command-bar" class="bg-white rounded-2xl ring-1 ring-slate-100 px-3 sm:px-4 py-2.5 flex flex-wrap items-center gap-x-4 gap-y-3">
      <div class="flex items-center gap-2.5 min-w-0">
        <span class="grid place-items-center w-8 h-8 rounded-xl bg-indigo-50 text-base ring-1 ring-indigo-100 shrink-0">🎯</span>
        <div class="min-w-0">
          <div class="flex items-center gap-1.5 flex-wrap">
            <span class="text-sm font-bold text-slate-900">AI basket vs benchmarks</span>
            <span class="text-[8px] font-bold uppercase tracking-wider px-1 py-0.5 rounded bg-amber-100 text-amber-700 ring-1 ring-amber-200">Beta</span>
          </div>
          <div class="flex items-center gap-1.5 flex-wrap mt-1">
            <span class="text-[10px] text-slate-500 tabular-nums">${escapeHtml(modeLabel)} · ${fmtDateDMY(view.startDate)} → ${fmtDateDMY(view.endDate)} · ${Math.max(0, view.equityCurve.length - 1)} trading day${view.equityCurve.length === 2 ? "" : "s"}</span>
            ${convoChip}
          </div>
        </div>
      </div>
      <div class="flex items-center gap-3 ml-auto flex-wrap">
        <div class="flex items-stretch rounded-xl ring-1 ring-slate-200 bg-slate-50/40 divide-x divide-slate-200/70">
          ${stat("AI", view.finalReturn, aiDD, "text-indigo-700")}
          
          ${stat("Smallcap 250", view.niftyRet, niftyDD, "text-slate-500")}
          ${stat("Midcap 150", view.nifty500Ret, nifty500DD, "text-sky-600")}
        </div>
        <div class="h-9 w-px bg-slate-200 hidden sm:block"></div>
        <div class="flex items-center gap-2">
          <input id="lkp-file-input" type="file" accept=".xlsx,.xls,.csv" class="hidden" />
          <button id="lkp-upload-btn" type="button" title="Upload the month's client basket (Excel / CSV)" class="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg ring-1 ring-slate-200 bg-white hover:bg-indigo-50 hover:ring-indigo-200 text-xs font-semibold text-slate-700">⬆ <span class="hidden sm:inline">Upload</span></button>
          ${lkpOverride() ? `<button id="lkp-reset-upload" type="button" title="Discard the uploaded basket and use the published one" class="text-[11px] font-semibold text-slate-500 hover:text-rose-600">Reset</button>` : ""}
          ${hits ? renderStrategyAlertsBell(hits) : ""}
        </div>
      </div>
    </div>
  `;
}

// Return convention toggle — governs BOTH baskets (AI + manual) so the
// comparison stays like-to-like. Booked (default) freezes each pick at its
// first target / SL the day it hits — the desk's real exit, so a runaway
// winner can't flatter the basket. If held marks to market ("what if we
// never sold"). Flips the headline, chart lines and every basket row.
function renderManualReturnToggle(view) {
  if (view.manualFinalReturn == null) return "";
  const mode = state.manualReturnMode === "held" ? "held" : "booked";
  const pill = (k, label) => {
    const on = mode === k;
    return `<button type="button" data-manual-return="${k}" class="px-3 py-1.5 rounded-lg text-xs font-semibold transition ${on ? "bg-amber-500 text-white shadow-sm" : "text-slate-600 hover:text-slate-900 hover:bg-slate-50"}">${label}</button>`;
  };
  return `
    <div class="bg-white rounded-2xl ring-1 ring-slate-100 px-2 py-1.5 flex items-center gap-1.5 shrink-0" title='Booked closes each pick (AI + manual) the day it hits its first target or SL — the real exit. "If held" is where it went after.'>
      ${pill("booked", "Booked")}
      ${pill("held", "If held")}
    </div>
  `;
}

// "Capital & charges" — the real-world money layer. Shows ₹ capital
// growth net of transaction costs and exposes every input (capital,
// cash buffer, charge rates) as a knob. Founder ask: "if a user comes
// with ₹5L, how does it grow net of charges." All cadences supported.
function moneySimTile(label, value, sub, valueCls = "text-slate-900") {
  return `
    <div class="rounded-xl ring-1 ring-slate-200 bg-slate-50/60 px-3 py-2">
      <div class="text-[9px] font-bold uppercase tracking-wider text-slate-500">${label}</div>
      <div class="text-base font-display font-bold ${valueCls} tabular-nums mt-0.5 truncate">${value}</div>
      <div class="text-[10px] text-slate-400 mt-0.5 truncate">${sub}</div>
    </div>`;
}

function renderSimPanel(view) {
  const p = simPrefs;
  const money = (n) => `₹${Math.round(n).toLocaleString("en-IN")}`;
  const pct = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const net = view.finalReturn;
  const gross = view.grossFinalReturn;
  const drag = (gross != null && net != null) ? gross - net : null;
  const cap = view.startCapital ?? p.capital;
  const val = view.finalValue ?? cap;
  const charges = view.totalCharges ?? 0;
  const netCls = net == null ? "text-slate-900" : net >= 0 ? "text-emerald-700" : "text-rose-700";

  const inputs = SIM_FIELDS.map((f) => `
    <label class="flex flex-col gap-1">
      <span class="text-[10px] font-semibold uppercase tracking-wider text-slate-500">${f.label}${f.suffix ? ` <span class="text-slate-300 normal-case">${f.suffix}</span>` : ""}</span>
      <input type="number" data-sim-field="${f.key}" value="${p[f.key]}" step="${f.step}" min="0"
             class="w-full rounded-lg ring-1 ring-slate-200 focus:ring-2 focus:ring-indigo-300 px-2 py-1.5 text-sm tabular-nums outline-none" />
      ${f.note ? `<span class="text-[9px] text-slate-400">${f.note}</span>` : ""}
    </label>`).join("");

  // What the charges actually cost at this ticket size — the number that
  // decides whether a strategy's edge survives contact with the broker.
  const basket = 7;
  const cs = chargeSummary(p.capital, basket, p);
  const costCard = `
    <div class="rounded-xl ring-1 ring-amber-200 bg-amber-50/60 px-3 py-2.5 mt-3">
      <div class="text-[11px] text-amber-900 leading-snug">
        At <b>${money(p.capital)}</b> across ${basket} stocks that is <b>${money(cs.perPosition)}</b> a position —
        <b>${cs.roundTripPct.toFixed(2)}%</b> to buy and sell it once
        (the flat ₹${p.dpFee} DP fee is <b>${cs.dpSharePct.toFixed(0)}%</b> of that).
        Replacing the whole basket every month costs <b>${cs.monthlyRotationAnnualPct.toFixed(1)}% a year</b>.
      </div>
    </div>`;

  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-center justify-between gap-3 flex-wrap mb-3">
        <div class="flex items-center gap-2">
          <span class="text-base">💰</span>
          <h3 class="font-display font-bold text-slate-900 text-sm">Capital &amp; charges</h3>
          <span class="text-[10px] text-slate-400 hidden sm:inline">net of transaction costs</span>
        </div>
        <button type="button" id="sim-reset" class="text-[11px] font-semibold text-slate-500 hover:text-indigo-600">Reset defaults</button>
      </div>
      <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-3">
        ${moneySimTile("Capital deployed", money(cap), "starting pot")}
        ${moneySimTile("Value now · net", money(val), pct(net) + " net of charges", netCls)}
        ${moneySimTile("Charges paid", money(charges), drag != null ? `${drag >= 0 ? "−" : "+"}${Math.abs(drag).toFixed(2)}% drag` : "")}
        ${moneySimTile("Gross · no charges", pct(gross), "frictionless model")}
      </div>
      <details class="group">
        <summary class="cursor-pointer text-[11px] font-semibold text-indigo-600 hover:text-indigo-700 select-none">Adjust capital, buffer &amp; charge rates ▾</summary>
        <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2.5 mt-3">${inputs}</div>
        ${costCard}
        <div class="text-[10px] text-slate-400 mt-2 leading-snug">Zerodha NSE delivery rates. STT and the exchange/SEBI fees hit both sides, stamp duty only the buy, and the DP fee is a flat charge per stock sold — which is why it costs far more, in percentage terms, on a small position than a large one.</div>
      </details>
    </div>`;
}

// Chart geometry kept in module-level constants so the hover setup
// reads the exact same numbers used to draw the chart — otherwise the
// crosshair drifts off the line.
const ACTIVE_CHART = {
  W: 820, H: 240,
  M: { left: 50, right: 18, top: 16, bottom: 32 },
  color: { ai: "#6366f1", manual: "#f59e0b", nifty: "#94a3b8", nifty500: "#0ea5e9" },
};
// Categorical palette for the per-stock ("Stocks") chart mode — AI names get
// indigo/violet family tones, Manual names get amber/orange, so the two
// baskets stay visually grouped even with 14 lines on screen.
const STOCK_PALETTE_AI = ["#6366f1", "#8b5cf6", "#4f46e5", "#7c3aed", "#3b82f6", "#6d28d9", "#2563eb"];
const STOCK_PALETTE_MANUAL = ["#f59e0b", "#ea580c", "#d97706", "#f97316", "#b45309", "#fb923c", "#c2410c"];

// Max-upside on a cumulative-return curve = the highest retPct point
// ever reached during the window. Tells the analyst "how far did this
// strategy run up at its peak". Null when the curve has no data.
function curveMaxUpside(curve) {
  if (!curve?.length) return null;
  let max = -Infinity;
  for (const p of curve) if (p?.retPct != null && p.retPct > max) max = p.retPct;
  return max === -Infinity ? null : max;
}

// Max-drawdown = worst peak-to-trough decline as an equity-factor
// ratio (current_factor / peak_factor − 1), expressed as a percentage.
// Subtracting cumulative retPct values directly would give percentage
// points, not the drawdown the analyst expects: peak +100% → trough
// +50% must report −25%, not −50%, since the portfolio went from 2× to
// 1.5× its starting value. 0 if the curve only ever climbed.
function curveMaxDrawdown(curve) {
  if (!curve?.length) return null;
  let peakFactor = -Infinity, maxDD = 0;
  for (const p of curve) {
    if (p?.retPct == null) continue;
    const factor = 1 + p.retPct / 100;
    if (factor > peakFactor) peakFactor = factor;
    const dd = (factor / peakFactor - 1) * 100;
    if (dd < maxDD) maxDD = dd;
  }
  return maxDD;
}

// KPI grid surfaced below the cumulative-return chart. Three cards:
//   1. Max upside   — peak return % for Manual / AI / Nifty
//   2. Max drawdown — worst peak-to-trough for the same three
//   3. Alpha matrix — pairwise outperformance (Manual−AI, AI−Nifty, Manual−Nifty)
// The "AI vs Nifty" alpha used to live in the hero card; moved here so
// every relative comparison is in one place.
function renderStrategyKpis(view) {
  const aiCurve = view.equityCurve || [];
  const manualCurve = view.manualCurve || [];
  const niftyCurve = view.niftyCurve || [];
  const nifty500Curve = view.nifty500Curve || [];
  const aiUp = curveMaxUpside(aiCurve);
  const manualUp = curveMaxUpside(manualCurve);
  const niftyUp = curveMaxUpside(niftyCurve);
  const nifty500Up = curveMaxUpside(nifty500Curve);
  const aiDD = curveMaxDrawdown(aiCurve);
  const manualDD = curveMaxDrawdown(manualCurve);
  const niftyDD = curveMaxDrawdown(niftyCurve);
  const nifty500DD = curveMaxDrawdown(nifty500Curve);
  const aiFinal = view.finalReturn;
  const manualFinal = view.manualFinalReturn;
  const niftyFinal = view.niftyRet;
  const nifty500Final = view.nifty500Ret;
  const manualVsAi = (manualFinal != null && aiFinal != null) ? manualFinal - aiFinal : null;
  const aiVsNifty = (aiFinal != null && niftyFinal != null) ? aiFinal - niftyFinal : null;
  const aiVsNifty500 = (aiFinal != null && nifty500Final != null) ? aiFinal - nifty500Final : null;

  const fmtPct = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const cls = (v) => v == null ? "text-slate-500" : v >= 0 ? "text-emerald-700" : "text-rose-700";
  const row = (dotCls, label, value) => `
    <div class="flex items-center gap-2 text-sm">
      <span class="inline-block w-2 h-2 rounded-full ${dotCls}"></span>
      <span class="text-slate-600">${label}</span>
      <span class="ml-auto font-bold tabular-nums ${cls(value)}">${fmtPct(value)}</span>
    </div>`;

  // Net performance as of today — the headline "where do we stand" the client
  // asked to see up front: both baskets against both benchmarks.
  const cardNet = `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-indigo-200 p-4">
      <div class="flex items-center justify-between mb-2">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Net performance · as of ${view.endDate ? fmtDateDMY(view.endDate) : "today"}</div>
        <div class="text-indigo-500 text-base">≡</div>
      </div>
      <div class="space-y-1.5">
        <div class="flex items-center gap-2 text-sm">
          <span class="inline-block w-2 h-2 rounded-full bg-indigo-500"></span>
          <span class="text-slate-600">AI basket</span>
          <span class="ml-auto text-right">
            <span class="font-bold tabular-nums ${cls(aiFinal)}">${fmtPct(aiFinal)}</span>
            <span class="text-[10px] text-slate-400">net</span>
            ${view.grossFinalReturn != null ? `<div class="text-[10px] text-slate-400 tabular-nums">${fmtPct(view.grossFinalReturn)} before charges</div>` : ""}
          </span>
        </div>
        ${row("bg-slate-400", "Smallcap 250", niftyFinal)}
        ${row("bg-sky-500", "Midcap 150", nifty500Final)}
      </div>    </div>`;

  const cardUpside = `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
      <div class="flex items-center justify-between mb-2">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Max upside</div>
        <div class="text-emerald-500 text-base">▲</div>
      </div>
      <div class="space-y-1.5">
        ${row("bg-indigo-500", "AI", aiUp)}
        ${row("bg-slate-400", "Smallcap 250", niftyUp)}
        ${row("bg-sky-500", "Midcap 150", nifty500Up)}
      </div>    </div>`;

  const cardDrawdown = `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
      <div class="flex items-center justify-between mb-2">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Max drawdown</div>
        <div class="text-rose-500 text-base">▼</div>
      </div>
      <div class="space-y-1.5">
        ${row("bg-indigo-500", "AI", aiDD)}
        ${row("bg-slate-400", "Smallcap 250", niftyDD)}
        ${row("bg-sky-500", "Midcap 150", nifty500DD)}
      </div>    </div>`;

  const alphaRow = (dotCls, label, value) => `
    <div class="flex items-center gap-2 text-sm">
      <span class="inline-flex items-center gap-1.5 text-slate-600">
        <span class="inline-block w-2 h-2 rounded-full ${dotCls}"></span>
        ${label}
      </span>
      <span class="ml-auto font-bold tabular-nums ${cls(value)}">${fmtPct(value)}</span>
    </div>`;

  const cardAlpha = `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
      <div class="flex items-center justify-between mb-2">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Relative alpha</div>
        <div class="text-indigo-500 text-base">⇆</div>
      </div>
      <div class="space-y-1.5">
        ${alphaRow("bg-slate-400", "AI − Smallcap 250", aiVsNifty)}
        ${alphaRow("bg-sky-500", "AI − Midcap 150", aiVsNifty500)}
      </div>    </div>`;

  return `
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
      ${cardNet}
      ${cardUpside}
      ${cardDrawdown}
      ${cardAlpha}
    </div>
  `;
}

// Chart mode toggle — Basket line vs one line per stock (client ask).
function renderChartModeToggle(mode) {
  const btn = (k, label) => `<button type="button" data-chart-mode="${k}" class="px-2.5 py-1 rounded-lg text-[11px] font-bold transition ${mode === k ? "bg-indigo-600 text-white shadow-sm" : "text-slate-500 hover:text-slate-700"}">${label}</button>`;
  return `<div class="inline-flex items-center gap-0.5 bg-slate-100 rounded-lg p-0.5">${btn("basket", "Basket")}${btn("stocks", "Stocks")}</div>`;
}

function renderActiveCumulativeChart(view) {
  const { W, H, M, color } = ACTIVE_CHART;
  const innerW = W - M.left - M.right;
  const innerH = H - M.top - M.bottom;
  const mode = (state.strategyChartMode === "stocks") ? "stocks" : "basket";
  const toggle = renderChartModeToggle(mode);
  const shell = (legend, bodyHtml) => `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
      <div class="flex items-center justify-between mb-2 gap-2 flex-wrap">
        <div class="flex items-center gap-2">
          <h3 class="font-semibold text-slate-900 text-sm">Cumulative return</h3>
          ${toggle}
        </div>
        <div class="flex items-center gap-x-3 gap-y-1 text-[11px] text-slate-500 flex-wrap justify-end">${legend}</div>
      </div>
      ${bodyHtml}`;
  // "Stocks" mode — small multiples: one mini-graph per holding (the client's
  // literal "seven graphs"), shown a basket at a time so 14 overlaid lines
  // never fight for the same axis. "Both" still shows every holding — as 14
  // separate, individually-readable tiles, so nothing is lost.
  if (mode === "stocks") {
    const basket = ["ai", "manual", "both"].includes(state.strategyStockBasket) ? state.strategyStockBasket : "ai";
    const bBtn = (k, label) => `<button type="button" data-stock-basket="${k}" class="px-2.5 py-1 rounded-lg text-[11px] font-bold transition ${basket === k ? "bg-indigo-600 text-white shadow-sm" : "text-slate-500 hover:text-slate-700"}">${label}</button>`;
    const basketToggle = `<div class="inline-flex items-center gap-0.5 bg-slate-100 rounded-lg p-0.5">${bBtn("ai", "AI")}</div>`;
    return shell(basketToggle, renderPerStockSmallMultiples(view, basket)) + `</div>`;
  }

  const pts = view.equityCurve;
  if (pts.length < 2) {
    // Distinguish "the market has not opened since you locked the basket"
    // from "something is broken". The old code drew a line through weekend
    // snapshots here, so the entry charge looked like a price move.
    return shell("", `
      <div class="text-xs text-slate-500 py-6 text-center leading-relaxed">
        No trading day since the basket was locked — nothing to plot yet.<br>
        <span class="text-slate-400">The line starts on the next market open.</span>
      </div>`) + `</div>`;
  }
  const niftyByDate = new Map((view.niftyCurve || []).map((p) => [p.date, p.retPct]));
  const nifty500ByDate = new Map((view.nifty500Curve || []).map((p) => [p.date, p.retPct]));
  const manualByDate = new Map((view.manualCurve || []).map((p) => [p.date, p.retPct]));
  const hasManual = view.manualCurve && view.manualCurve.some((p) => p.retPct != null);
  const hasN500 = view.nifty500Curve && view.nifty500Curve.some((p) => p.retPct != null);
  const allVals = [];
  for (const p of pts) {
    allVals.push(p.retPct);
    const n = niftyByDate.get(p.date); if (n != null) allVals.push(n);
    const n5 = nifty500ByDate.get(p.date); if (n5 != null) allVals.push(n5);
    const m = manualByDate.get(p.date); if (m != null) allVals.push(m);
  }
  const yMin = Math.min(0, ...allVals);
  const yMax = Math.max(0, ...allVals);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.12, yHi = yMax + ySpan * 0.12;
  const xAt = (i) => M.left + (pts.length <= 1 ? 0 : (i / (pts.length - 1)) * innerW);
  const yAt = (v) => M.top + (1 - (v - yLo) / (yHi - yLo)) * innerH;

  const buildLine = (getter) => {
    const segs = []; let cur = [];
    pts.forEach((p, i) => {
      const v = getter(p);
      if (v == null) { if (cur.length) { segs.push(cur); cur = []; } return; }
      cur.push(`${cur.length === 0 ? "M" : "L"} ${xAt(i).toFixed(2)} ${yAt(v).toFixed(2)}`);
    });
    if (cur.length) segs.push(cur);
    return segs.map((s) => s.join(" ")).join(" ");
  };
  const activePath = buildLine((p) => p.retPct);
  const niftyPath = buildLine((p) => niftyByDate.get(p.date));
  const nifty500Path = buildLine((p) => nifty500ByDate.get(p.date));
  const manualPath = buildLine((p) => manualByDate.get(p.date));

  const first = [xAt(0), yAt(pts[0].retPct)];
  const last = [xAt(pts.length - 1), yAt(pts[pts.length - 1].retPct)];
  const baseY = yAt(0).toFixed(2);
  const areaPath = `${pts.map((p, i) => `${i === 0 ? "M" : "L"} ${xAt(i).toFixed(2)} ${yAt(p.retPct).toFixed(2)}`).join(" ")} L ${last[0].toFixed(2)} ${baseY} L ${first[0].toFixed(2)} ${baseY} Z`;

  const yTicks = [0, 1, 2, 3, 4].map((step) => {
    const v = yLo + (yHi - yLo) * (step / 4);
    const yy = (M.top + innerH - (step / 4) * innerH).toFixed(2);
    const isZero = Math.abs(v) < 0.05;
    return `
      <line x1="${M.left}" x2="${(W - M.right).toFixed(2)}" y1="${yy}" y2="${yy}" stroke="${isZero ? "#94a3b8" : "#e2e8f0"}" stroke-width="${isZero ? 0.9 : 0.6}" stroke-dasharray="${isZero ? "0" : "3 4"}" />
      <text x="${(M.left - 8).toFixed(2)}" y="${yy}" text-anchor="end" dominant-baseline="middle" font-size="10" font-weight="500" fill="#94a3b8">${v >= 0 ? "+" : ""}${v.toFixed(1)}%</text>`;
  }).join("");
  const tickEvery = Math.max(1, Math.ceil(pts.length / 6));
  const xTicks = pts.map((p, i) => (i % tickEvery !== 0 && i !== pts.length - 1)
    ? ""
    : `<text x="${xAt(i).toFixed(2)}" y="${(M.top + innerH + 16).toFixed(2)}" text-anchor="middle" font-size="10" fill="#64748b">${fmtDateDM(p.date)}</text>`
  ).join("");

  const legend = `
    <span class="inline-flex items-center gap-1.5"><span class="w-2.5 h-0.5" style="background:${color.ai}"></span>AI basket</span>
    ${hasManual ? `<span class="inline-flex items-center gap-1.5"><span class="w-2.5 h-0.5" style="background:${color.manual}"></span>Manual basket</span>` : ""}
    <span class="inline-flex items-center gap-1.5" title="Benchmark"><span class="w-2.5 h-0.5 border-t border-dashed" style="border-color:${color.nifty}"></span>Smallcap 250</span>
    ${hasN500 ? `<span class="inline-flex items-center gap-1.5" title="Benchmark — Nifty Midcap 150">​<span class="w-2.5 h-0.5 border-t border-dashed" style="border-color:${color.nifty500}"></span>Midcap 150</span>` : ""}`;

  const body = `
      <div id="active-chart-container" class="relative">
        <svg id="active-chart-svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" class="w-full select-none" style="max-height:280px">
          <defs>
            <linearGradient id="activeStrategyArea" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stop-color="${color.ai}" stop-opacity="0.22"/>
              <stop offset="100%" stop-color="${color.ai}" stop-opacity="0"/>
            </linearGradient>
          </defs>
          ${yTicks}
          <path d="${areaPath}" fill="url(#activeStrategyArea)" />
          ${niftyPath ? `<path d="${niftyPath}" fill="none" stroke="${color.nifty}" stroke-width="1.6" stroke-dasharray="4 4" stroke-linecap="round" stroke-linejoin="round" />` : ""}
          ${nifty500Path ? `<path d="${nifty500Path}" fill="none" stroke="${color.nifty500}" stroke-width="1.6" stroke-dasharray="4 4" stroke-linecap="round" stroke-linejoin="round" />` : ""}
          ${manualPath ? `<path d="${manualPath}" fill="none" stroke="${color.manual}" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round" />` : ""}
          <path d="${activePath}" fill="none" stroke="${color.ai}" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round" />
          ${xTicks}
          <line id="active-chart-guide" x1="0" y1="${M.top}" x2="0" y2="${(M.top + innerH).toFixed(2)}" stroke="#94a3b8" stroke-width="0.8" stroke-dasharray="2 3" opacity="0" />
          <circle id="active-chart-dot-ai" cx="0" cy="0" r="4.5" fill="#fff" stroke="${color.ai}" stroke-width="2.2" opacity="0" />
          <circle id="active-chart-dot-manual" cx="0" cy="0" r="3.5" fill="#fff" stroke="${color.manual}" stroke-width="2" opacity="0" />
          <circle id="active-chart-dot-nifty" cx="0" cy="0" r="3.5" fill="#fff" stroke="${color.nifty}" stroke-width="2" opacity="0" />
          <circle id="active-chart-dot-nifty500" cx="0" cy="0" r="3.5" fill="#fff" stroke="${color.nifty500}" stroke-width="2" opacity="0" />
          <rect id="active-chart-capture" x="0" y="0" width="${W}" height="${H}" fill="transparent" />
        </svg>
        <div id="active-chart-tooltip" class="hidden absolute z-10 pointer-events-none -translate-x-1/2 -translate-y-[calc(100%+10px)] bg-slate-900/95 backdrop-blur text-white text-[11px] rounded-xl shadow-2xl ring-1 ring-slate-700/60 px-3 py-2 whitespace-nowrap"></div>
      </div>
      ${view.entryCostPct > 0 || view.liveMark ? `<div class="text-[10px] text-slate-400 mt-1.5 leading-snug">${view.entryCostPct > 0 ? `The basket line starts <b>${view.entryCostPct.toFixed(2)}%</b> down: that is the buying charge, debited the moment the basket is locked, not a price move. Only trading days are plotted.` : ""}${view.liveMark ? ` The last point is <b>today, live</b> — it moves with the market and settles at the close.` : ""}</div>` : ""}`;
  return shell(legend, body) + `</div>`;
}

// Per-stock ("Stocks" mode) small multiples — one mini cumulative-return graph
// per holding (the client's "seven graphs"), instead of layering 14 near-
// identical lines on one axis. Color encodes OUTCOME (green up / red down,
// redundantly backed by the signed % and ▲/▼ so it reads under CVD); the ticker
// title carries identity; a shared vertical scale keeps tiles comparable; a
// faint Nifty 50 line sits behind each for a market reference. Booked/held aware
// (the curves already freeze at the exit level); each tile drills into the modal.
function renderPerStockSmallMultiples(view, basket) {
  const nifty = view.niftyCurve || [];
  const aiStocks = (view.aiStockCurves || []).map((s) => ({ ...s, side: "ai" }));
  const manualStocks = (view.manualStockCurves || []).map((s) => ({ ...s, side: "manual" }));
  const stocks = basket === "ai" ? aiStocks : basket === "manual" ? manualStocks : [...aiStocks, ...manualStocks];
  if (!stocks.length) {
    const which = basket === "manual" ? "manual" : basket === "ai" ? "AI" : "";
    return `<div class="text-xs text-slate-500 py-6 text-center">No ${which} holdings with tracking data yet.</div>`;
  }
  const dates = (stocks[0].curve || []).map((p) => p.date);
  if (dates.length < 2) return `<div class="text-xs text-slate-500 py-6 text-center">Not enough days to plot yet.</div>`;
  // One shared vertical scale across every tile (+ Nifty) so a +8% runner
  // visibly towers over a +1% name — panels are directly comparable.
  const allVals = [];
  for (const s of stocks) for (const p of s.curve) if (p.retPct != null) allVals.push(p.retPct);
  for (const p of nifty) if (p.retPct != null) allVals.push(p.retPct);
  const yMin = Math.min(0, ...allVals), yMax = Math.max(0, ...allVals);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.1, yHi = yMax + ySpan * 0.1;
  const W = 150, H = 56, PADX = 3, PADT = 4, PADB = 4;
  const innerW = W - 2 * PADX, innerH = H - PADT - PADB;
  const dateIdx = new Map(dates.map((d, i) => [d, i]));
  const xAt = (i) => PADX + (dates.length <= 1 ? 0 : (i / (dates.length - 1)) * innerW);
  const yAt = (v) => PADT + (1 - (v - yLo) / (yHi - yLo)) * innerH;
  const pathFor = (curve) => {
    const segs = []; let cur = [];
    (curve || []).forEach((p) => {
      const i = dateIdx.get(p.date);
      if (i == null || p.retPct == null) { if (cur.length) { segs.push(cur); cur = []; } return; }
      cur.push(`${cur.length === 0 ? "M" : "L"} ${xAt(i).toFixed(1)} ${yAt(p.retPct).toFixed(1)}`);
    });
    if (cur.length) segs.push(cur);
    return segs.map((s) => s.join(" ")).join(" ");
  };
  const niftyPath = pathFor(nifty);
  const zeroY = yAt(0).toFixed(1);
  const sign = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
  const showBadge = basket === "both";
  const UP = "#059669", DOWN = "#e11d48";
  const tiles = stocks.map((s) => {
    const finalRet = lastRet(s.curve);
    const up = curveMaxUpside(s.curve), dd = curveMaxDrawdown(s.curve);
    const col = (finalRet ?? 0) >= 0 ? UP : DOWN;
    const d = pathFor(s.curve);
    const badge = showBadge
      ? (s.side === "ai"
          ? `<span class="inline-flex items-center px-1 py-0 rounded bg-indigo-100 text-indigo-700 text-[8px] font-bold uppercase tracking-wider">AI</span>`
          : `<span class="inline-flex items-center px-1 py-0 rounded bg-amber-100 text-amber-700 text-[8px] font-bold uppercase tracking-wider">Man</span>`)
      : "";
    return `
      <button type="button" data-cohort-row data-cohort-side="${s.side}" data-ticker="${escapeHtml(s.ticker)}" data-seg-anchor="${escapeHtml(dates[0] || "")}" class="text-left rounded-xl ring-1 ring-slate-200 bg-white hover:ring-indigo-300 hover:shadow-sm transition p-2 focus:outline-none focus:ring-2 focus:ring-indigo-400">
        <div class="flex items-center justify-between gap-1">
          <span class="text-[11px] font-bold text-slate-800 truncate" title="${escapeHtml(s.name || s.ticker)}">${escapeHtml(s.ticker)}</span>
          <span class="flex items-center gap-1 flex-shrink-0">${badge}<span class="text-[11px] font-bold tabular-nums" style="color:${col}">${sign(finalRet)}</span></span>
        </div>
        <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" class="w-full mt-1 block" style="height:50px">
          <line x1="${PADX}" x2="${W - PADX}" y1="${zeroY}" y2="${zeroY}" stroke="#e2e8f0" stroke-width="1" stroke-dasharray="2 2" vector-effect="non-scaling-stroke" />
          ${niftyPath ? `<path d="${niftyPath}" fill="none" stroke="#cbd5e1" stroke-width="1" stroke-dasharray="3 2" vector-effect="non-scaling-stroke" />` : ""}
          ${d ? `<path d="${d}" fill="none" stroke="${col}" stroke-width="1.7" stroke-linejoin="round" stroke-linecap="round" vector-effect="non-scaling-stroke" />` : ""}
        </svg>
        <div class="flex items-center justify-between mt-1 text-[9px] tabular-nums">
          <span class="text-emerald-600" title="Max upside since entry">▲ ${sign(up)}</span>
          <span class="text-rose-600" title="Max drawdown since entry">▼ ${sign(dd)}</span>
        </div>
      </button>`;
  }).join("");
  return `
      <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-2">${tiles}</div>
      <div class="mt-3 flex items-center gap-x-3 gap-y-1 text-[10px] text-slate-400 flex-wrap">
        <span class="inline-flex items-center gap-1"><span class="w-2.5 h-0.5" style="background:${UP}"></span>up</span>
        <span class="inline-flex items-center gap-1"><span class="w-2.5 h-0.5" style="background:${DOWN}"></span>down</span>
        <span class="inline-flex items-center gap-1"><span class="w-2.5 h-0.5 border-t border-dashed" style="border-color:#cbd5e1"></span>Smallcap 250</span>
        <span>· shared scale · honours the Booked / If-held toggle · click a tile to drill in.</span>
      </div>`;
}

function wireStrategyChartMode() {
  $$("#active-content [data-chart-mode]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.chartMode === "stocks" ? "stocks" : "basket";
    if (v === (state.strategyChartMode || "basket")) return;
    state.strategyChartMode = v;
    rerenderKeepingScroll(renderActive);
  }));
}

function wireStrategyStockBasket() {
  $$("#active-content [data-stock-basket]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.stockBasket;
    if (!["ai", "manual", "both"].includes(v) || v === (state.strategyStockBasket || "ai")) return;
    state.strategyStockBasket = v;
    rerenderKeepingScroll(renderActive);
  }));
}

function setupActiveChartHover(view) {
  if (!view) return;
  const container = document.getElementById("active-chart-container");
  const svg = document.getElementById("active-chart-svg");
  const capture = document.getElementById("active-chart-capture");
  const guide = document.getElementById("active-chart-guide");
  const dotAi = document.getElementById("active-chart-dot-ai");
  const dotManual = document.getElementById("active-chart-dot-manual");
  const dotNifty = document.getElementById("active-chart-dot-nifty");
  const dotNifty500 = document.getElementById("active-chart-dot-nifty500");
  const tip = document.getElementById("active-chart-tooltip");
  if (!container || !svg || !capture || !tip) return;

  const { W, H, M, color } = ACTIVE_CHART;
  const innerW = W - M.left - M.right;
  const innerH = H - M.top - M.bottom;
  const pts = view.equityCurve;
  if (pts.length < 2) return;
  const niftyByDate = new Map((view.niftyCurve || []).map((p) => [p.date, p.retPct]));
  const nifty500ByDate = new Map((view.nifty500Curve || []).map((p) => [p.date, p.retPct]));
  const manualByDate = new Map((view.manualCurve || []).map((p) => [p.date, p.retPct]));
  const allVals = [];
  for (const p of pts) {
    allVals.push(p.retPct);
    const n = niftyByDate.get(p.date); if (n != null) allVals.push(n);
    const n5 = nifty500ByDate.get(p.date); if (n5 != null) allVals.push(n5);
    const m = manualByDate.get(p.date); if (m != null) allVals.push(m);
  }
  const yMin = Math.min(0, ...allVals);
  const yMax = Math.max(0, ...allVals);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.12, yHi = yMax + ySpan * 0.12;
  const xAt = (i) => M.left + (i / (pts.length - 1)) * innerW;
  const yAt = (v) => M.top + (1 - (v - yLo) / (yHi - yLo)) * innerH;

  function show(idx) {
    const p = pts[idx];
    const aiPx = xAt(idx);
    const aiPy = yAt(p.retPct);
    guide.setAttribute("x1", aiPx); guide.setAttribute("x2", aiPx); guide.setAttribute("opacity", "1");
    dotAi.setAttribute("cx", aiPx); dotAi.setAttribute("cy", aiPy); dotAi.setAttribute("opacity", "1");
    const mVal = manualByDate.get(p.date);
    if (mVal != null) { dotManual.setAttribute("cx", aiPx); dotManual.setAttribute("cy", yAt(mVal)); dotManual.setAttribute("opacity", "1"); }
    else dotManual.setAttribute("opacity", "0");
    const nVal = niftyByDate.get(p.date);
    if (nVal != null) { dotNifty.setAttribute("cx", aiPx); dotNifty.setAttribute("cy", yAt(nVal)); dotNifty.setAttribute("opacity", "1"); }
    else dotNifty.setAttribute("opacity", "0");
    const n5Val = nifty500ByDate.get(p.date);
    if (n5Val != null && dotNifty500) { dotNifty500.setAttribute("cx", aiPx); dotNifty500.setAttribute("cy", yAt(n5Val)); dotNifty500.setAttribute("opacity", "1"); }
    else if (dotNifty500) dotNifty500.setAttribute("opacity", "0");

    // Position the tooltip at the point's ACTUAL screen location using the
    // SVG's live coordinate matrix. Deriving it from rect.width / W breaks
    // when preserveAspectRatio letterboxes the chart (wide container →
    // height-capped → side padding): today's point, drawn at the chart's
    // right edge, would otherwise be mislocated out into the empty padding.
    const contRect = container.getBoundingClientRect();
    const ctm = svg.getScreenCTM();
    let tipX, tipY;
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = aiPx; sp.y = aiPy;
      const scr = sp.matrixTransform(ctm);
      tipX = scr.x - contRect.left; tipY = scr.y - contRect.top;
    } else {
      const rect = svg.getBoundingClientRect();
      tipX = aiPx * (rect.width / W); tipY = aiPy * (rect.height / H);
    }
    const fmt = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
    const cls = (v) => v == null ? "text-slate-300" : v >= 0 ? "text-emerald-300" : "text-rose-300";
    tip.innerHTML = `
      <div class="font-bold text-sm leading-tight">${fmtDateDMY(p.date)}</div>
      <div class="mt-1 flex items-center gap-2"><span class="w-1.5 h-1.5 rounded-full" style="background:${color.ai}"></span><span class="text-slate-300">AI active</span><span class="ml-auto font-bold tabular-nums ${cls(p.retPct)}">${fmt(p.retPct)}</span></div>
      ${mVal != null ? `<div class="mt-0.5 flex items-center gap-2"><span class="w-1.5 h-1.5 rounded-full" style="background:${color.manual}"></span><span class="text-slate-300">Manual</span><span class="ml-auto font-bold tabular-nums ${cls(mVal)}">${fmt(mVal)}</span></div>` : ""}
      ${nVal != null ? `<div class="mt-0.5 flex items-center gap-2"><span class="w-1.5 h-1.5 rounded-full" style="background:${color.nifty}"></span><span class="text-slate-300">Smallcap 250</span><span class="ml-auto font-bold tabular-nums ${cls(nVal)}">${fmt(nVal)}</span></div>` : ""}
      ${n5Val != null ? `<div class="mt-0.5 flex items-center gap-2"><span class="w-1.5 h-1.5 rounded-full" style="background:${color.nifty500}"></span><span class="text-slate-300">Midcap 150</span><span class="ml-auto font-bold tabular-nums ${cls(n5Val)}">${fmt(n5Val)}</span></div>` : ""}
    `;
    tip.classList.remove("hidden");
    tip.style.left = tipX + "px";
    tip.style.top = tipY + "px";
    // Flip / clamp horizontally so today's tooltip (pinned to the right
    // edge) stays inside the card instead of clipping off-screen.
    tip.style.transform = "translate(-50%, calc(-100% - 10px))";
    requestAnimationFrame(() => {
      const tr = tip.getBoundingClientRect();
      let dx = 0;
      if (tr.left < contRect.left + 6) dx = (contRect.left + 6) - tr.left;
      else if (tr.right > contRect.right - 6) dx = (contRect.right - 6) - tr.right;
      if (dx !== 0) tip.style.transform = `translate(calc(-50% + ${dx}px), calc(-100% - 10px))`;
    });
  }
  function hide() {
    guide.setAttribute("opacity", "0");
    dotAi.setAttribute("opacity", "0");
    dotManual.setAttribute("opacity", "0");
    dotNifty.setAttribute("opacity", "0");
    if (dotNifty500) dotNifty500.setAttribute("opacity", "0");
    tip.classList.add("hidden");
  }
  // Pointer x → nearest snapshot index, via the SVG's inverse CTM so
  // letterbox padding, page zoom and device-pixel-ratio can't shift it.
  function eventToIdx(e) {
    const t = e.touches ? e.touches[0] : e;
    let vx;
    const ctm = svg.getScreenCTM();
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = t.clientX; sp.y = t.clientY;
      vx = sp.matrixTransform(ctm.inverse()).x;
    } else {
      const rect = svg.getBoundingClientRect();
      vx = (t.clientX - rect.left) / rect.width * W;
    }
    const rel = (vx - M.left) / innerW;
    return Math.max(0, Math.min(pts.length - 1, Math.round(rel * (pts.length - 1))));
  }
  capture.addEventListener("mousemove", (e) => show(eventToIdx(e)));
  capture.addEventListener("mouseleave", hide);
  capture.addEventListener("touchstart", (e) => { show(eventToIdx(e)); e.preventDefault(); }, { passive: false });
  capture.addEventListener("touchmove",  (e) => { show(eventToIdx(e)); e.preventDefault(); }, { passive: false });
  capture.addEventListener("touchend", hide);
  // Surface today (last point) on render so it reads without hovering.
  show(pts.length - 1);
}

// Two summary cards side by side: AI strategy hits + Manual basket hits.
// Manual uses the client's per-stock TGT1 / SL from the LKP upload.
function renderActiveOverallHitsSplit(view) {
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
      <div class="flex items-baseline justify-between mb-3 flex-wrap gap-2">
        <h3 class="font-display font-bold text-slate-900 text-base">Overall accuracy</h3>
      </div>
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
        ${renderActiveSummaryCard("AI active", "indigo", view.hitSummary)}
        ${view.manualSummary
          ? renderActiveSummaryCard("Manual basket", "amber", view.manualSummary)
          : `<div class="rounded-xl bg-amber-50/40 ring-1 ring-amber-100 p-4 text-center">
              <div class="text-[11px] font-bold uppercase tracking-wider text-amber-700 mb-1">Manual basket</div>
              <div class="text-xs text-slate-500">No client basket uploaded for this anchor.</div>
            </div>`}
      </div>
    </div>
  `;
}

function renderActiveSummaryCard(label, palette, summary) {
  const dot = palette === "indigo" ? "bg-indigo-500" : "bg-amber-500";
  const hitRateStr = summary.hitRate == null ? "—" : `${summary.hitRate.toFixed(0)}%`;
  const avgT = summary.avgDaysToTarget == null ? "—" : `${summary.avgDaysToTarget.toFixed(1)}d`;
  const avgS = summary.avgDaysToSL == null ? "—" : `${summary.avgDaysToSL.toFixed(1)}d`;
  return `
    <div class="rounded-xl bg-slate-50 ring-1 ring-slate-100 p-3">
      <div class="flex items-center gap-2 mb-2">
        <span class="inline-block w-2 h-2 rounded-full ${dot}"></span>
        <div class="text-[11px] font-bold uppercase tracking-wider text-slate-600">${escapeHtml(label)}</div>
        <span class="ml-auto text-[10px] text-slate-400 tabular-nums">${summary.total} pick${summary.total === 1 ? "" : "s"}</span>
      </div>
      <div class="grid grid-cols-4 gap-2 text-center">
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-emerald-700">${summary.targetHits}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">🎯 Target</div>
          <div class="text-[10px] text-slate-400 tabular-nums">avg ${avgT}</div>
        </div>
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-rose-700">${summary.slHits}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">⚠ SL</div>
          <div class="text-[10px] text-slate-400 tabular-nums">avg ${avgS}</div>
        </div>
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-slate-700">${summary.open}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">Open</div>
          <div class="text-[10px] text-slate-400 tabular-nums">tracking</div>
        </div>
        <div>
          <div class="text-base font-display font-extrabold tabular-nums text-indigo-700">${hitRateStr}</div>
          <div class="text-[9px] uppercase tracking-wider text-slate-500 font-bold">Hit rate</div>
          <div class="text-[10px] text-slate-400 tabular-nums">${summary.targetHits} of ${summary.total}</div>
        </div>
      </div>
    </div>
  `;
}

// AI + Manual pick-row tables, side by side. Mirrors the History tab's
// Accuracy view layout so the analyst can compare per-pick outcomes.
// Per-segment basket roster (founder ask — match the old History tab
// affordance from SS2/SS3). For multi-segment cadences (Active Daily /
// Weekly / Monthly with 2+ segments) renders segment pills + the
// selected segment's AI top 7 alongside the Manual basket. For Passive
// or a single-segment chain, no pills — just the basket.
function renderActiveSegmentedBaskets(view, mode) {
  if (!view?.segments?.length) return "";
  const segCount = view.segments.length;
  const selectedIdx = clampStrategySegmentIdx(segCount);
  const selected = view.segments[selectedIdx];

  const headerLabel = mode === "passive"
    ? "Held basket"
    : view.kind === "daily"   ? "AI basket per day"
    : view.kind === "weekly"  ? "AI basket per week"
    : view.kind === "monthly" ? "AI basket per month"
    : "AI basket";
  const pillsHtml = segCount > 1 ? renderActiveSegmentPills(view.segments, selectedIdx, view.kind) : "";

  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      ${segCount > 1 ? `
        <div class="flex items-baseline gap-3 mb-3 flex-wrap">
          <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500">${escapeHtml(headerLabel)}</div>
          ${pillsHtml}
        </div>` : ""}
      <div class="grid grid-cols-1 ${view.manualPicks?.length ? "lg:grid-cols-2" : ""} gap-3">
        ${renderAiBasketTable(selected, view, mode)}
        ${renderManualBasketTable(view.manualPicks)}
      </div>
    </div>
  `;
}

function clampStrategySegmentIdx(segCount) {
  if (state.strategySegmentIdx == null) return segCount - 1;
  return Math.max(0, Math.min(segCount - 1, state.strategySegmentIdx));
}

function renderActiveSegmentPills(segments, selectedIdx, kind) {
  const pillLabel = (seg) => {
    if (kind === "daily") return fmtDateDM(seg.startDate);
    return `${seg.label} · ${fmtDateDM(seg.startDate)}–${fmtDateDM(seg.endDate)}`;
  };
  return `
    <div class="flex items-center gap-1 overflow-x-auto -my-1 py-1 max-w-full">
      ${segments.map((seg, i) => `
        <button type="button" data-strategy-seg="${i}" class="px-2.5 py-1 rounded-lg text-[11px] font-semibold whitespace-nowrap transition flex-shrink-0 ${i === selectedIdx ? "bg-indigo-600 text-white shadow-sm" : "bg-slate-50 text-slate-600 ring-1 ring-slate-200 hover:ring-indigo-300 hover:text-indigo-700"}">
          ${escapeHtml(pillLabel(seg))}
        </button>
      `).join("")}
    </div>
  `;
}

// Open / Closed pill shared by the AI + Manual Overview rosters. Every pick
// is a top pick, so the rating (STRONG BUY etc.) is redundant here — it lives
// in the click-through drill modal. Closed is coloured by outcome.
function rosterStatusBadge(isClosed, reason) {
  if (!isClosed) {
    return `<span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap bg-slate-100 text-slate-600 ring-slate-200" title="Position open (tap for rating & detail)">Open</span>`;
  }
  const green = reason === "TARGET";
  return `<span class="inline-flex items-center gap-0.5 px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap ${green ? "bg-emerald-50 text-emerald-700 ring-emerald-200" : "bg-rose-50 text-rose-700 ring-rose-200"}" title="Position closed — ${green ? "target hit" : "stop-loss hit"} (tap for rating & detail)">🔒 Closed</span>`;
}

// AI top 7 for the selected segment — name, entry → exit, return %, and an
// Open/Closed status. Honours the Booked/If-held toggle (booked freezes a hit
// pick at its target/SL level, if-held marks to market). Row click opens the
// drill modal (with the rating) anchored at the segment's start date so the
// chart overlay levels match the row's entry close.
function renderAiBasketTable(segment, view, mode) {
  if (!segment) {
    return `<div><div class="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-2">AI picks</div><div class="text-xs text-slate-500">No AI basket for this segment.</div></div>`;
  }
  const cache = state.cache.history || {};
  const todayClose = cache.todayClose || {};
  const booked = state.manualReturnMode !== "held";
  // Match each roster stock to its tracked pick (status + target/SL levels),
  // keyed by ticker + entry date. Passive is a single held segment, so a hit
  // in view.picks belongs to this roster; the return convention mirrors the
  // Booked/If-held toggle exactly as the manual roster does.
  const pickByKey = new Map((view.picks || []).map((p) => [`${p.ticker}|${p.entryDate}`, p]));

  const subLabel = mode === "passive"
    ? `Held since ${fmtDateDMY(segment.startDate)}`
    : view.kind === "daily"   ? `Day · ${fmtDateDMY(segment.startDate)}`
    : view.kind === "weekly"  ? `${segment.label} · ${fmtDateDM(segment.startDate)} → ${fmtDateDM(segment.endDate)}`
    : view.kind === "monthly" ? `${segment.label} · ${fmtDateDM(segment.startDate)} → ${fmtDateDM(segment.endDate)}`
    : `Held since ${fmtDateDMY(segment.startDate)}`;

  let retSum = 0, retN = 0;
  const rows = segment.top7.map((s) => {
    const today = (typeof todayClose[s.ticker] === "number") ? todayClose[s.ticker] : s.close;
    const heldRet = ((today / s.close) - 1) * 100;
    const p = pickByKey.get(`${s.ticker}|${segment.startDate}`);
    const hit = p && (p.status === "TARGET_HIT" || p.status === "SL_HIT");
    const isBooked = booked && hit;
    const reason = (p && p.status === "SL_HIT") ? "SL" : "TARGET";
    // Booked freezes at the target/SL LEVEL — the real exit — not the live mark.
    const exitLevel = isBooked ? (p.status === "TARGET_HIT" ? p.target : p.sl) : null;
    const displayRet = isBooked ? ((exitLevel / s.close) - 1) * 100 : heldRet;
    const nowPx = today, exitPx = isBooked ? exitLevel : today;
    retSum += displayRet; retN++;
    const retCls = displayRet >= 0 ? "text-emerald-700" : "text-rose-700";
    // For a booked pick the "return" is the captured level and "now" is where
    // the stock actually sits — shown small so the two are never conflated.
    const heldNote = isBooked
      ? `<div class="text-[9px] text-slate-400 tabular-nums">now ₹${formatPrice(nowPx)} · if held ${heldRet >= 0 ? "+" : ""}${heldRet.toFixed(1)}%</div>`
      : "";
    return `
      <tr data-cohort-row data-cohort-side="ai" data-ticker="${escapeHtml(s.ticker)}" data-seg-anchor="${escapeHtml(segment.startDate)}" class="border-t border-slate-100 cursor-pointer transition hover:bg-indigo-50/40">
        <td class="py-2 pl-3 pr-2">
          <div class="font-semibold text-slate-900 text-sm truncate max-w-[180px]" title="${escapeHtml(s.name || s.ticker)}">${escapeHtml(s.name || s.ticker)}</div>
          ${heldNote}
        </td>
        <td class="py-2 px-2 text-right tabular-nums text-slate-500 text-sm">₹${formatPrice(s.close)}</td>
        <td class="py-2 px-2 text-right tabular-nums text-slate-700 text-sm">₹${formatPrice(exitPx)}${isBooked ? ` <span class="text-slate-300">🔒</span>` : ""}</td>
        <td class="py-2 px-2 text-right"><span class="tabular-nums text-sm font-bold ${retCls}">${displayRet >= 0 ? "+" : ""}${displayRet.toFixed(2)}%</span></td>
        <td class="py-2 pl-2 pr-3 text-right">${rosterStatusBadge(isBooked, reason)}</td>
      </tr>`;
  }).join("");

  // Basket average = mean of the equal-weighted per-row returns above, which
  // IS the simple average of the percentages shown. Net applies the buy-side
  // charge, so the two figures reconcile on screen instead of the headline
  // looking pulled from nowhere.
  const grossBasket = retN ? retSum / retN : 0;
  const netBasket = ((1 + grossBasket / 100) * (1 - buyRate(simPrefs)) - 1) * 100;
  const gCls = grossBasket >= 0 ? "text-emerald-700" : "text-rose-700";
  const nCls = netBasket >= 0 ? "text-emerald-700" : "text-rose-700";

  return `
    <div>
      <div class="flex items-center gap-1.5 mb-2 flex-wrap">
        <span class="inline-block w-2 h-2 rounded-full bg-indigo-500"></span>
        <div class="text-[11px] font-bold uppercase tracking-wider text-slate-700">AI picks</div>
        <span class="text-[10px] text-slate-400 truncate">· ${escapeHtml(subLabel)} · ${segment.top7.length} stocks</span>
      </div>
      <div class="rounded-lg ring-1 ring-slate-100 overflow-hidden">
        <table class="w-full">
          <thead class="bg-slate-50/70 text-[10px] uppercase tracking-wider text-slate-400">
            <tr>
              <th class="py-1.5 pl-3 pr-2 text-left font-semibold">Stock</th>
              <th class="py-1.5 px-2 text-right font-semibold">Bought at</th>
              <th class="py-1.5 px-2 text-right font-semibold">Now / exit</th>
              <th class="py-1.5 px-2 text-right font-semibold">Return</th>
              <th class="py-1.5 pl-2 pr-3 text-right font-semibold">Status</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
          <tfoot class="bg-slate-50 border-t-2 border-slate-200">
            <tr>
              <td class="py-2.5 pl-3 pr-2 text-[11px] font-bold uppercase tracking-wider text-slate-500" colspan="3">Basket average</td>
              <td class="py-2.5 px-2 text-right" colspan="2">
                <span class="tabular-nums text-base font-bold ${nCls}">${netBasket >= 0 ? "+" : ""}${netBasket.toFixed(2)}%</span>
                <span class="text-[10px] text-slate-400 ml-1">net</span>
                <div class="text-[10px] text-slate-400 tabular-nums">${grossBasket >= 0 ? "+" : ""}${grossBasket.toFixed(2)}% before charges</div>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  `;
}

// Manual basket roster — same shape as the AI table. Out-of-coverage
// picks render greyed with a "Not Covered" badge. The basket is
// always anchored at upload regardless of which segment is selected.
function renderManualBasketTable(manualPicks) {
  // No manual basket on this dashboard — render nothing rather than an
  // empty panel advertising a feature that does not apply.
  if (!manualPicks?.length) return "";
  const todayClose = state.cache.history?.todayClose || {};
  const inCoverage = manualPicks.filter((p) => !p.notCovered).length;
  const booked = state.manualReturnMode !== "held";

  const rows = manualPicks.map((r) => {
    if (r.notCovered) {
      return `
        <div class="grid grid-cols-12 items-center gap-2 py-2 px-2 rounded-lg bg-slate-50/40">
          <div class="col-span-8 min-w-0">
            <div class="font-semibold text-slate-500 text-sm truncate" title="${escapeHtml(r.outReason || "")}">${escapeHtml(r.name)}</div>
            <div class="text-[10px] text-slate-400 truncate">${escapeHtml(r.outReason || "")}</div>
          </div>
          <div class="col-span-4 text-right">
            <span class="inline-flex items-center px-1.5 py-0 rounded bg-slate-200 text-slate-600 ring-1 ring-slate-300 text-[9px] font-bold uppercase tracking-wider">Not Covered</span>
          </div>
        </div>`;
    }
    const today = (typeof todayClose[r.ticker] === "number") ? todayClose[r.ticker] : r.entryPrice;
    const heldRet = (r.heldRet != null) ? r.heldRet : (r.entryPrice ? ((today / r.entryPrice) - 1) * 100 : null);
    const isBooked = booked && r.booking?.booked;
    const displayRet = isBooked ? r.bookedRet : heldRet;
    const exitPx = isBooked ? r.booking.bookPrice : today;
    const retCls = displayRet == null ? "text-slate-500" : displayRet >= 0 ? "text-emerald-700" : "text-rose-700";
    // Booked picks show the exit price in the sub-line ("→ ₹xxx · locked") and
    // the if-held mark separately; Open/Closed status sits on the right, rating
    // moves to the drill modal — same shape as the AI roster.
    const heldNote = (isBooked && heldRet != null)
      ? `<div class="text-[9px] text-slate-400 tabular-nums" title="Where the stock is now — booked return is what was actually captured at the exit level">if held ${heldRet >= 0 ? "+" : ""}${heldRet.toFixed(1)}%</div>`
      : "";
    return `
      <button type="button" data-cohort-row data-cohort-side="manual" data-ticker="${escapeHtml(r.ticker)}" data-seg-anchor="${escapeHtml(r.entryDate || "")}" class="w-full text-left grid grid-cols-12 items-center gap-2 py-2 px-2 rounded-lg cursor-pointer transition hover:bg-amber-50/40 hover:ring-1 hover:ring-amber-200">
        <div class="col-span-6 sm:col-span-5 min-w-0">
          <div class="font-semibold text-slate-900 text-sm truncate" title="${escapeHtml(r.name)}">${escapeHtml(r.name)}</div>
          <div class="text-[10px] text-slate-500 tabular-nums">₹${formatPrice(r.entryPrice)} → ₹${formatPrice(exitPx)}${isBooked ? " · locked" : ""}</div>
        </div>
        <div class="col-span-3 sm:col-span-3 text-right">
          <div class="tabular-nums text-sm font-bold ${retCls}">${displayRet == null ? "—" : (displayRet >= 0 ? "+" : "") + displayRet.toFixed(2) + "%"}</div>
          ${heldNote}
        </div>
        <div class="col-span-3 sm:col-span-4 text-right">${rosterStatusBadge(isBooked, r.booking?.reason)}</div>
      </button>`;
  }).join("");

  return `
    <div>
      <div class="flex items-center gap-1.5 mb-2 flex-wrap">
        <span class="inline-block w-2 h-2 rounded-full bg-amber-500"></span>
        <div class="text-[11px] font-bold uppercase tracking-wider text-slate-700">Manual picks</div>
        <span class="text-[10px] text-slate-400 truncate">· ${manualPicks.length} stocks · ${inCoverage} in coverage</span>
      </div>
      <div class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 p-1 space-y-0.5">${rows}</div>
    </div>
  `;
}

// Two-column per-pick accuracy view (AI + Manual). Restored after the
// segment-basket refactor — used below the basket roster to surface
// target / SL / status / Just Hit / proximity tint per pick (founder
// said this is the main affordance).
function renderActivePickRowsSplit(view) {
  return `
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-3">
      ${renderActivePickColumn("AI Picks · per-pick status", "indigo", view.picks, "ai")}
      ${renderActivePickColumn("Manual Picks · per-pick status", "amber", view.manualPicks, "manual")}
    </div>
  `;
}

function renderActivePickColumn(title, palette, picks, side) {
  const dot = palette === "indigo" ? "bg-indigo-500" : "bg-amber-500";
  if (!picks || !picks.length) {
    return `
      <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
        <div class="flex items-center gap-2 mb-3">
          <span class="inline-block w-2 h-2 rounded-full ${dot}"></span>
          <h3 class="font-display font-bold text-slate-900 text-sm">${escapeHtml(title)}</h3>
        </div>
        <div class="text-xs text-slate-500">No picks for this column.</div>
      </div>`;
  }
  const fmtPct = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
  // Collapse long columns (AI accumulates a row per week×stock) to a
  // preview that lines up with the shorter Manual column; a "+ N more"
  // button expands the rest in place. Founder ask: keep the two columns
  // visually level, no big scroll.
  const PICK_PREVIEW_LIMIT = 7;
  const rows = picks.map((r, i) => {
    // Overflow rows are hidden via inline display:none (always wins over
    // the grid/flex utility classes) and toggled by the "+ N more" button.
    const isExtra = i >= PICK_PREVIEW_LIMIT;
    const extraCls = isExtra ? " pick-extra-row" : "";
    const extraStyle = isExtra ? ' style="display:none"' : "";
    if (r.notCovered) {
      return `
        <div${extraStyle} class="grid grid-cols-12 items-center gap-2 py-1.5 px-2 rounded-lg bg-slate-50/40${extraCls}">
          <div class="col-span-8 min-w-0">
            <div class="font-semibold text-slate-500 text-xs truncate" title="${escapeHtml(r.outReason || "")}">${escapeHtml(r.name)}</div>
            <div class="text-[10px] text-slate-400 truncate">${escapeHtml(r.outReason || "")}</div>
          </div>
          <div class="col-span-4 text-right">
            <span class="inline-flex items-center px-1.5 py-0 rounded bg-slate-200 text-slate-600 ring-1 ring-slate-300 text-[9px] font-bold uppercase tracking-wider">Not Covered</span>
          </div>
        </div>`;
    }
    const statusCls = r.status === "TARGET_HIT" ? "bg-emerald-100 text-emerald-700 ring-emerald-200"
      : r.status === "SL_HIT" ? "bg-rose-100 text-rose-700 ring-rose-200"
      : "bg-slate-100 text-slate-600 ring-slate-200";
    const statusLabel = r.status === "TARGET_HIT" ? "🔒 Closed"
      : r.status === "SL_HIT" ? "🔒 SL Exit"
      : "Open";
    // Closed picks are locked at the target / SL LEVEL (the real exit), not
    // the overshooting close — so the "exit" line reads as the locked price.
    const exitLevel = r.status === "TARGET_HIT" ? r.target : r.status === "SL_HIT" ? r.sl : null;
    const exit = exitLevel != null ? `₹${formatPrice(exitLevel)} · ${fmtDateDMY(r.hitDate)} · ${r.daysToHit}d`
      : r.currentClose != null ? `now ₹${formatPrice(r.currentClose)}` : "—";
    let rowTint = "hover:bg-slate-50";
    if (r.status === "TARGET_HIT") rowTint = "bg-emerald-50 ring-1 ring-emerald-200";
    else if (r.status === "SL_HIT") rowTint = "bg-rose-50 ring-1 ring-rose-200";
    else if (r.proximity != null && r.proximity >= 0.75) rowTint = "bg-emerald-50/40 hover:bg-emerald-50/70";
    else if (r.proximity != null && r.proximity <= 0.25) rowTint = "bg-rose-50/40 hover:bg-rose-50/70";
    const currentReturnHtml = r.currentReturnPct != null
      ? `<span class="text-[10px] tabular-nums font-semibold ${r.currentReturnPct >= 0 ? "text-emerald-700" : "text-rose-700"} ml-1">${fmtPct(r.currentReturnPct)}</span>`
      : "";
    // Peak excursion nuggets — max upside / max downside since entry,
    // each with the day count to reach it. Founder ask: judge the
    // optimal rebalance horizon from where picks actually peak.
    const peak = r.peak;
    const peakHtml = peak ? `
        <div class="col-span-12 flex items-center gap-1.5 flex-wrap pt-1 mt-0.5 border-t border-slate-100">
          <span class="text-[8px] uppercase tracking-wider text-slate-400 font-bold">Excursion</span>
          <span title="Peak upside since entry and days taken to reach it" class="inline-flex items-center gap-1 px-1.5 py-0 rounded bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200 text-[9px] font-semibold tabular-nums">▲ ${fmtPct(peak.maxUpsidePct)} · ${peak.daysToMaxUpside}d</span>
          <span title="Worst drawdown since entry and days taken to reach it" class="inline-flex items-center gap-1 px-1.5 py-0 rounded bg-rose-50 text-rose-700 ring-1 ring-rose-200 text-[9px] font-semibold tabular-nums">▼ ${fmtPct(peak.maxDownsidePct)} · ${peak.daysToMaxDownside}d</span>
        </div>` : "";
    return `
      <button type="button"${extraStyle} data-cohort-row data-cohort-side="${side}" data-ticker="${escapeHtml(r.ticker)}" data-seg-anchor="${escapeHtml(r.entryDate || "")}" class="w-full text-left grid grid-cols-12 items-center gap-x-2 gap-y-0 py-1.5 px-2 rounded-lg cursor-pointer transition ${rowTint}${extraCls} hover:ring-1 hover:ring-indigo-200">
        <div class="col-span-5 min-w-0">
          <div class="font-semibold text-slate-900 text-xs truncate">${escapeHtml(r.name)}${r.cohortLabel ? ` <span class="text-[9px] text-slate-400 font-normal">· ${escapeHtml(r.cohortLabel)}</span>` : ""}</div>
          <div class="text-[10px] text-slate-500 tabular-nums">Entry ${fmtDateDMY(r.entryDate)} · ₹${formatPrice(r.entryPrice)}${currentReturnHtml}</div>
        </div>
        <div class="col-span-3 text-[10px] tabular-nums text-slate-600 text-right">
          <div>T: ₹${formatPrice(r.target)} <span class="text-emerald-600">${fmtPct(r.targetPct)}</span></div>
          <div>SL: ₹${formatPrice(r.sl)} <span class="text-rose-600">${fmtPct(r.slPct)}</span></div>
        </div>
        <div class="col-span-4 text-right">
          <div class="flex items-center justify-end flex-wrap gap-1">
            <span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 ${statusCls}">${statusLabel}</span>
          </div>
          <div class="text-[10px] text-slate-500 mt-0.5 truncate">${exit}</div>
        </div>
        ${peakHtml}
      </button>`;
  }).join("");
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-baseline justify-between mb-3 flex-wrap gap-2">
        <div class="flex items-center gap-2">
          <span class="inline-block w-2 h-2 rounded-full ${dot}"></span>
          <h3 class="font-display font-bold text-slate-900 text-sm">${escapeHtml(title)}</h3>
        </div>
        <span class="text-[11px] text-slate-500">${picks.filter((p) => !p.notCovered).length} trackable</span>
      </div>
      <div data-pick-list="${side}" class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 px-1 py-1 space-y-0.5">${rows}</div>
      ${picks.length > PICK_PREVIEW_LIMIT
        ? `<button type="button" data-pick-toggle="${side}" data-expanded="0" class="mt-2 w-full text-center text-[11px] font-semibold text-indigo-600 hover:text-indigo-700 py-1.5 rounded-lg ring-1 ring-indigo-200 hover:bg-indigo-50/60 transition">+ ${picks.length - PICK_PREVIEW_LIMIT} more</button>`
        : ""}
    </div>
  `;
}

// Renderers moved above renderActiveBetaCaveat (split AI / Manual layouts).

// Sector rebalance-timing roll-up. Groups this cadence's AI picks by
// sector and averages each sector's peak excursion + days-to-peak — so
// the desk reads the natural rebalance horizon per sector (founder ask:
// "defense peaks in ~10 days, real estate ~3 — rebalance each on its own
// clock"). Built on the same per-pick excursion data as the nuggets.
// Sector (broad) vs Industry (GICS sub-industry, finer) grouping for the
// rebalance-timing table — toggled in the UI via `sectorTimingBy`.
let sectorTimingBy = "sector";
// Strategy-tab "Balancing" sub-tab grouping — Sector / Industry / Stock. Stock
// is the default (client: "for us stocks are more important than the sector or
// industry"). Merges the old separate Sector + Industry sub-tabs into one.
let strategyBalanceBy = "stock";

// Flatten BOTH baskets (AI + Manual) into one list of tracked picks the
// balancing view groups/lists. AI picks already carry sector/industry; manual
// picks get theirs looked up from the latest snapshot. Only in-coverage picks
// with excursion (peak) data qualify. Client ask: "add your AI stock as well,
// not just the manual stocks."
function balancingPicks(view) {
  const snaps = state.cache.history?.snapshots || [];
  const lastSnap = snaps[snaps.length - 1];
  const metaByTicker = new Map();
  if (lastSnap) for (const s of lastSnap.stocks) if (s.ticker) metaByTicker.set(s.ticker, { sector: s.sector || null, industry: s.industry || null });
  const booked = state.manualReturnMode !== "held";
  const ai = (view.picks || []).filter((p) => !p.notCovered && p.peak && p.ticker).map((p) => ({
    ticker: p.ticker, name: p.name, side: "AI",
    sector: p.sector || metaByTicker.get(p.ticker)?.sector || null,
    industry: p.industry || metaByTicker.get(p.ticker)?.industry || null,
    // Booked mode: a closed AI pick returns its target/SL level (like the
    // frozen basket + chart), not the drifting mark-to-market close.
    ret: (booked && p.status === "TARGET_HIT") ? p.targetPct
       : (booked && p.status === "SL_HIT") ? p.slPct
       : p.currentReturnPct,
    peak: p.peak, entryDate: p.entryDate,
  }));
  const manual = (view.manualPicks || []).filter((p) => !p.notCovered && p.peak && p.ticker).map((p) => {
    const isBooked = booked && p.booking?.booked;
    const meta = metaByTicker.get(p.ticker) || {};
    return {
      ticker: p.ticker, name: p.name, side: "Manual",
      sector: p.sector || meta.sector || null,
      industry: p.industry || meta.industry || null,
      ret: isBooked ? p.bookedRet : p.heldRet, peak: p.peak, entryDate: p.entryDate,
    };
  });
  return [...ai, ...manual];
}

// The merged Sector / Industry / Stock balancing panel (replaces the two
// separate Sector and Industry sub-tabs). One toggle switches the grouping;
// "Stocks" gives the per-stock table the client asked for.
function renderStrategyBalancing(view) {
  const by = ["sector", "industry", "stock"].includes(strategyBalanceBy) ? strategyBalanceBy : "sector";
  const picks = balancingPicks(view);
  const nAI = picks.filter((p) => p.side === "AI").length;
  const nMan = picks.filter((p) => p.side === "Manual").length;
  const toggleBtn = (val, text) => `<button data-balance-by="${val}" class="px-3 py-1 rounded-lg text-[11px] font-bold uppercase tracking-wider transition ${by === val ? "bg-indigo-600 text-white shadow-sm" : "text-slate-500 hover:text-slate-700"}">${text}</button>`;
  const toggle = `<div class="inline-flex items-center gap-0.5 bg-slate-100 rounded-xl p-0.5">${toggleBtn("sector", "Sector")}${toggleBtn("industry", "Industry")}${toggleBtn("stock", "Stocks")}</div>`;
  const body = picks.length
    ? (by === "stock" ? renderBalancingStockTable(picks) : renderBalancingGroupTable(picks, by))
    : `<div class="text-center text-sm text-slate-500 py-6">No picks with tracking data yet — needs a few AI / manual picks carrying excursion data across this window.</div>`;
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-center justify-between gap-2 flex-wrap mb-1">
        <div class="flex items-center gap-2">
          <span class="text-base">🧭</span>
          <h3 class="font-display font-bold text-slate-900 text-sm">Sector · Industry · Stock balancing</h3>
        </div>
        ${toggle}
      </div>
      <div class="text-[11px] text-slate-500 mb-3">Covers <strong>both baskets</strong> — ${nAI} AI + ${nMan} Manual pick${nMan === 1 ? "" : "s"}. ${by === "stock" ? "Every pick, stock by stock — click a row to drill in." : `Grouped by ${by}, fastest-peaking first — the natural rebalance horizon.`}</div>
      ${body}
    </div>`;
}

// Grouped roll-up (Sector or Industry) built from the COMBINED AI + Manual
// picks — same shape as the old sector-timing table, now covering both baskets.
function renderBalancingGroupTable(picks, by) {
  const rows = buildSectorTiming(picks, by);
  const label = by === "industry" ? "Industry" : "Sector";
  if (!rows.length) return `<div class="text-center text-sm text-slate-500 py-6">No ${label.toLowerCase()} data on these picks yet.</div>`;
  const pct = (v) => `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const body = rows.map((r) => `
    <tr class="border-t border-slate-100">
      <td class="py-2 pr-2 min-w-0"><div class="font-semibold text-slate-800 text-xs truncate" title="${escapeHtml(r.sector)}">${escapeHtml(r.sector)}</div><div class="text-[10px] text-slate-400">${r.n} pick${r.n === 1 ? "" : "s"}</div></td>
      <td class="py-2 px-2 text-right tabular-nums text-emerald-700 font-semibold">${pct(r.avgUpside)}</td>
      <td class="py-2 px-2 text-right tabular-nums text-slate-700 font-semibold">${r.avgDaysToPeak.toFixed(1)}d</td>
      <td class="py-2 px-2 text-right tabular-nums text-rose-700 font-semibold">${pct(r.avgDownside)}</td>
      <td class="py-2 pl-2 text-right"><span class="inline-flex items-center px-2 py-0.5 rounded-lg bg-indigo-50 text-indigo-700 ring-1 ring-indigo-200 text-[11px] font-bold tabular-nums">~${r.suggestedRebalance}d</span></td>
    </tr>`).join("");
  return `
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="text-[10px] font-bold uppercase tracking-wider text-slate-500">
        <th class="text-left pb-1 pr-2">${label}</th><th class="text-right pb-1 px-2">Avg peak</th><th class="text-right pb-1 px-2">Days to peak</th><th class="text-right pb-1 px-2">Avg drawdown</th><th class="text-right pb-1 pl-2">Rebalance</th>
      </tr></thead>
      <tbody>${body}</tbody>
    </table></div>`;
}

// The per-stock table (client ask): Stock · Sector · Manual/AI · Return ·
// Days-to-peak · Max drawdown · Rebalance. Fastest-peaking first; rows drill
// into the same modal the rest of the strategy tab uses.
function renderBalancingStockTable(picks) {
  const sorted = [...picks].sort((a, b) => (a.peak?.daysToMaxUpside ?? 999) - (b.peak?.daysToMaxUpside ?? 999));
  const pct = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const body = sorted.map((p) => {
    const sideBadge = p.side === "AI"
      ? `<span class="inline-flex items-center px-1.5 py-0 rounded bg-indigo-100 text-indigo-700 ring-1 ring-indigo-200 text-[9px] font-bold uppercase tracking-wider">AI</span>`
      : `<span class="inline-flex items-center px-1.5 py-0 rounded bg-amber-100 text-amber-700 ring-1 ring-amber-200 text-[9px] font-bold uppercase tracking-wider">Manual</span>`;
    const retCls = p.ret == null ? "text-slate-500" : p.ret >= 0 ? "text-emerald-700" : "text-rose-700";
    const reb = Math.max(1, Math.round(p.peak?.daysToMaxUpside || 0));
    return `
      <tr data-cohort-row data-cohort-side="${p.side === "AI" ? "ai" : "manual"}" data-ticker="${escapeHtml(p.ticker)}" data-seg-anchor="${escapeHtml(p.entryDate || "")}" class="border-t border-slate-100 cursor-pointer hover:bg-indigo-50/40 transition">
        <td class="py-2 pr-2 min-w-0"><div class="font-semibold text-slate-800 text-xs truncate" title="${escapeHtml(p.name || p.ticker)}">${escapeHtml(p.name || p.ticker)}</div><div class="text-[10px] text-slate-400 truncate">${escapeHtml(p.ticker)}</div></td>
        <td class="py-2 px-2 text-[11px] text-slate-600"><div class="truncate max-w-[130px]" title="${escapeHtml(p.sector || "")}">${escapeHtml(p.sector || "—")}</div></td>
        <td class="py-2 px-2 text-center">${sideBadge}</td>
        <td class="py-2 px-2 text-right tabular-nums font-bold ${retCls}">${pct(p.ret)}</td>
        <td class="py-2 px-2 text-right tabular-nums text-slate-700">${p.peak?.daysToMaxUpside ?? "—"}d</td>
        <td class="py-2 px-2 text-right tabular-nums text-rose-700 font-semibold">${pct(p.peak?.maxDownsidePct)}</td>
        <td class="py-2 pl-2 text-right"><span class="inline-flex items-center px-2 py-0.5 rounded-lg bg-indigo-50 text-indigo-700 ring-1 ring-indigo-200 text-[11px] font-bold tabular-nums">~${reb}d</span></td>
      </tr>`;
  }).join("");
  return `
    <div class="overflow-x-auto"><table class="w-full text-sm">
      <thead><tr class="text-[10px] font-bold uppercase tracking-wider text-slate-500">
        <th class="text-left pb-1 pr-2">Stock</th><th class="text-left pb-1 px-2">Sector</th><th class="text-center pb-1 px-2">Basket</th><th class="text-right pb-1 px-2">Return</th><th class="text-right pb-1 px-2">Days to peak</th><th class="text-right pb-1 px-2">Max drawdown</th><th class="text-right pb-1 pl-2">Rebalance</th>
      </tr></thead>
      <tbody>${body}</tbody>
    </table></div>`;
}

// Toggle handler for the Balancing sub-tab's Sector / Industry / Stock switch.
function wireStrategyBalanceToggle() {
  $$("#active-content [data-balance-by]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.balanceBy;
    if (!["sector", "industry", "stock"].includes(v) || v === strategyBalanceBy) return;
    strategyBalanceBy = v;
    rerenderKeepingScroll(renderActive);
  }));
}
function buildSectorTiming(picks, groupBy) {
  const by = groupBy === "industry" ? "industry" : "sector";
  const groups = new Map();
  for (const p of (picks || [])) {
    const key = p[by] || (by === "industry" ? p.sector : null);   // fall back to sector if industry missing
    if (!p.peak || !key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(p);
  }
  const mean = (arr) => arr.reduce((a, v) => a + v, 0) / arr.length;
  const rows = [];
  for (const [sector, ps] of groups.entries()) {
    rows.push({
      sector, n: ps.length,
      avgUpside: mean(ps.map((p) => p.peak.maxUpsidePct)),
      avgDaysToPeak: mean(ps.map((p) => p.peak.daysToMaxUpside)),
      avgDownside: mean(ps.map((p) => p.peak.maxDownsidePct)),
      avgDaysToTrough: mean(ps.map((p) => p.peak.daysToMaxDownside)),
      suggestedRebalance: Math.max(1, Math.round(mean(ps.map((p) => p.peak.daysToMaxUpside)))),
    });
  }
  // Fastest-peaking groups first; ties broken by larger (less noisy) sample.
  rows.sort((a, b) => a.avgDaysToPeak - b.avgDaysToPeak || b.n - a.n);
  return rows;
}

// byOverride ("sector" | "industry") forces the grouping and hides the inline
// toggle — used by the Strategy tab's dedicated Sector / Industry sub-tabs.
// Without it (Custom tab), it honours the shared sectorTimingBy toggle.
function renderSectorTiming(view, byOverride) {
  const forced = byOverride === "sector" || byOverride === "industry";
  const by = forced ? byOverride : (sectorTimingBy === "industry" ? "industry" : "sector");
  const rows = buildSectorTiming(view.picks, by);
  if (!rows.length) return "";
  // Only offer the Industry view if picks actually carry industry data.
  const hasIndustry = (view.picks || []).some((p) => p.industry);
  const label = by === "industry" ? "Industry" : "Sector";
  const toggleBtn = (val, text) => `<button data-sector-timing="${val}" class="px-2 py-0.5 rounded-md text-[10px] font-bold uppercase tracking-wider transition ${by === val ? "bg-indigo-600 text-white" : "text-slate-500 hover:text-slate-700"}">${text}</button>`;
  const toggle = (!forced && hasIndustry)
    ? `<div class="inline-flex items-center gap-0.5 bg-slate-100 rounded-lg p-0.5">${toggleBtn("sector", "Sector")}${toggleBtn("industry", "Industry")}</div>`
    : "";
  const pct = (v) => `${v >= 0 ? "+" : ""}${v.toFixed(2)}%`;
  const body = rows.map((r) => `
    <tr class="border-t border-slate-100">
      <td class="py-2 pr-2 min-w-0">
        <div class="font-semibold text-slate-800 text-xs truncate" title="${escapeHtml(r.sector)}">${escapeHtml(r.sector)}</div>
        <div class="text-[10px] text-slate-400">${r.n} pick${r.n === 1 ? "" : "s"}</div>
      </td>
      <td class="py-2 px-2 text-right tabular-nums text-emerald-700 font-semibold">${pct(r.avgUpside)}</td>
      <td class="py-2 px-2 text-right tabular-nums text-slate-700 font-semibold">${r.avgDaysToPeak.toFixed(1)}d</td>
      <td class="py-2 px-2 text-right tabular-nums text-rose-700 font-semibold">${pct(r.avgDownside)}</td>
      <td class="py-2 pl-2 text-right">
        <span class="inline-flex items-center px-2 py-0.5 rounded-lg bg-indigo-50 text-indigo-700 ring-1 ring-indigo-200 text-[11px] font-bold tabular-nums">~${r.suggestedRebalance}d</span>
      </td>
    </tr>`).join("");
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-center justify-between gap-2 flex-wrap mb-1">
        <div class="flex items-center gap-2">
          <span class="text-base">🧭</span>
          <h3 class="font-display font-bold text-slate-900 text-sm">${label} rebalance timing</h3>
        </div>
        <div class="flex items-center gap-2">${toggle}</div>
      </div>
      <div class="text-[11px] text-slate-500 mb-2">How long each ${label.toLowerCase()}'s picks take to peak on average — the natural rebalance horizon for that ${label.toLowerCase()}. Fastest-peaking on top.</div>
      <div class="overflow-x-auto">
        <table class="w-full text-sm">
          <thead>
            <tr class="text-[10px] font-bold uppercase tracking-wider text-slate-500">
              <th class="text-left pb-1 pr-2">${label}</th>
              <th class="text-right pb-1 px-2">Avg peak</th>
              <th class="text-right pb-1 px-2">Days to peak</th>
              <th class="text-right pb-1 px-2">Avg drawdown</th>
              <th class="text-right pb-1 pl-2">Rebalance</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </div>`;
}

function renderActiveBetaCaveat(view, anchorDate) {
  return `
    <div class="text-xs text-slate-500 bg-slate-50 ring-1 ring-slate-200 rounded-xl p-3 leading-relaxed">
      <strong>Beta caveat</strong> — strategy backtested from client upload date (${fmtDateDMY(anchorDate)}). Buys + sells transact at the same EOD close that marks the portfolio to market, so entry-day P&amp;L is zero by construction. Brokerage, STT, exchange &amp; GST are modelled via the Capital &amp; charges panel (round-trip per holding; rates editable) — slippage is not modelled yet.
    </div>
  `;
}

function wireActiveCadenceToggle() {
  $$("#active-content [data-cadence]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.cadence;
    if (!ACTIVE_CADENCES.includes(v) || v === state.activeCadence) return;
    state.activeCadence = v;
    state.strategySegmentIdx = null;   // reset to latest segment on cadence change
    saveActiveCadence(v);
    renderActive();
  }));
}

function wireStrategyModeToggle() {
  $$("#active-content [data-strategy-mode]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.strategyMode;
    if (!STRATEGY_MODES.includes(v) || v === state.strategyMode) return;
    state.strategyMode = v;
    state.strategySegmentIdx = null;   // reset to latest segment on mode change
    saveStrategyMode(v);
    renderActive();
  }));
}

function wireManualReturnToggle() {
  $$("#active-content [data-manual-return]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.manualReturn === "held" ? "held" : "booked";
    if (v === state.manualReturnMode) return;
    state.manualReturnMode = v;
    saveManualReturnMode(v);
    renderActive();
  }));
}

function wireStrategySubNav() {
  // Trade Plan capital. Re-renders the whole Strategy panel so both tables
  // recompute; debounced so typing "250000" does not re-render five times.
  const capEl = $("#plan-capital");
  if (capEl) {
    let t = null;
    capEl.addEventListener("input", () => {
      clearTimeout(t);
      t = setTimeout(() => {
        const v = Math.max(1000, Number(capEl.value) || 0);
        if (v === state.planCapital) return;
        state.planCapital = v; savePlanCapital(v);
        renderActive();
      }, 450);
    });
  }
  $$("#active-content [data-paper-month]").forEach((b) => b.addEventListener("click", () => {
    const m = b.dataset.paperMonth;
    if (!m || m === state.paperMonth) return;
    state.paperMonth = m;
    renderActive();
  }));

  $$("#active-content [data-plan-preset]").forEach((b) => b.addEventListener("click", () => {
    const v = Number(b.dataset.planPreset);
    if (!Number.isFinite(v) || v === state.planCapital) return;
    state.planCapital = v; savePlanCapital(v);
    renderActive();
  }));

  $$("#active-content [data-strategy-subtab]").forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.strategySubtab;
    if (!STRATEGY_SUBTABS.includes(v) || v === state.strategySubTab) return;
    state.strategySubTab = v;
    saveStrategySubTab(v);
    rerenderKeepingScroll(renderActive);
  }));
}

function wireManualMonthPills() {
  $$("#active-content [data-manual-month]").forEach((btn) => btn.addEventListener("click", () => {
    const m = btn.dataset.manualMonth;
    if (!m || m === state.manualMonth) return;
    state.manualMonth = m;
    state.strategySegmentIdx = null;   // reset to latest segment on month change
    renderActive();
  }));
}

function wireStrategySegmentPills() {
  $$("#active-content [data-strategy-seg]").forEach((btn) => btn.addEventListener("click", () => {
    const idx = Number(btn.dataset.strategySeg);
    if (!Number.isFinite(idx) || idx === state.strategySegmentIdx) return;
    state.strategySegmentIdx = idx;
    renderActive();
  }));
}

// Re-render a tab without yanking the viewport to the top. The async tab
// renderers briefly swap in a "Loading…" placeholder that collapses the
// page height and resets window scroll; capture the offset and restore it
// once the fresh content is painted (a microtask after innerHTML is set,
// i.e. before the next paint — so with cached data there's no visible jump).
function rerenderKeepingScroll(rerender) {
  const y = window.scrollY;
  Promise.resolve(rerender()).then(() => window.scrollTo(0, y)).catch(() => {});
}

// Sector ↔ Industry toggle on the rebalance-timing table (shared by the
// Active and Custom tabs — pass the tab's root + its re-render function).
function wireSectorTimingToggle(root, rerender) {
  $$(`${root} [data-sector-timing]`).forEach((btn) => btn.addEventListener("click", () => {
    const v = btn.dataset.sectorTiming;
    if ((v !== "sector" && v !== "industry") || v === sectorTimingBy) return;
    sectorTimingBy = v;
    rerenderKeepingScroll(rerender);
  }));
}

function wireStrategyAlertsDropdown() {
  const btn = $("#strategy-alerts-btn");
  const dropdown = $("#strategy-alerts-dropdown");
  if (!btn || !dropdown) return;
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    dropdown.classList.toggle("hidden");
  });
  document.addEventListener("click", (e) => {
    if (!dropdown.classList.contains("hidden") && !dropdown.contains(e.target) && !btn.contains(e.target)) {
      dropdown.classList.add("hidden");
    }
  }, { once: true });
}

function renderActiveBody(sim, stats, niftyOn) {
  const retColor = stats.totalRet >= 0 ? "text-emerald-600" : "text-rose-600";
  const retSign = stats.totalRet >= 0 ? "+" : "";
  const alphaTile = stats.alpha != null
    ? `<span class="${stats.alpha >= 0 ? "text-emerald-600" : "text-rose-600"} font-bold">${stats.alpha >= 0 ? "+" : ""}${stats.alpha.toFixed(2)}%</span> vs Nifty (${stats.niftyRet >= 0 ? "+" : ""}${stats.niftyRet.toFixed(2)}%)`
    : `<span class="text-slate-400">benchmark missing</span>`;

  const statTile = (label, value, sub) => `
    <div class="bg-white rounded-xl ring-1 ring-slate-200 p-3">
      <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500">${label}</div>
      <div class="text-lg font-bold text-slate-900 mt-0.5">${value}</div>
      ${sub ? `<div class="text-[11px] text-slate-500 mt-0.5">${sub}</div>` : ""}
    </div>`;

  const hitRateStr = stats.hitRate == null ? "—" : `${stats.hitRate.toFixed(0)}%`;
  const avgHoldStr = stats.avgHoldDays == null ? "—" : `${stats.avgHoldDays.toFixed(1)}d`;
  const avgWinStr = stats.avgWin == null ? "—" : `+${stats.avgWin.toFixed(2)}%`;
  const avgLossStr = stats.avgLoss == null ? "—" : `${stats.avgLoss.toFixed(2)}%`;

  return `
    <div class="space-y-4">
      <!-- Hero card -->
      <div class="bg-gradient-to-br from-indigo-50 via-white to-purple-50 rounded-2xl ring-1 ring-indigo-100 p-5">
        <div class="flex items-start justify-between gap-4 flex-wrap">
          <div>
            <div class="flex items-center gap-2">
              <h2 class="font-display font-bold text-xl text-slate-900">Active Basket — Daily Rebalance</h2>
              <span class="text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 ring-1 ring-amber-200">Beta</span>
            </div>
            <div class="text-sm text-slate-600 mt-1">
              Simulates buying every STRONG BUY at today's close and selling when it drops out.
              Starting capital ₹${ACTIVE_INITIAL_CAPITAL.toLocaleString("en-IN")}, equal-weight rebalanced daily, no transaction costs.
            </div>
            <div class="text-xs text-slate-500 mt-2">
              Period: ${fmtDateDMY(stats.startDate)} → ${fmtDateDMY(stats.endDate)} · ${stats.days} snapshot days
            </div>
          </div>
          <div class="text-right">
            <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500">Total return</div>
            <div class="${retColor} text-4xl font-bold leading-tight">${retSign}${stats.totalRet.toFixed(2)}%</div>
            <div class="text-sm mt-1">${alphaTile}</div>
            <div class="text-[11px] text-slate-500 mt-1">Portfolio ₹${Math.round(stats.finalValue).toLocaleString("en-IN")}</div>
          </div>
        </div>
      </div>

      <!-- Equity curve chart -->
      <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
        <div class="flex items-center justify-between mb-2">
          <h3 class="font-semibold text-slate-900 text-sm">Cumulative return</h3>
          <div class="flex items-center gap-3 text-[11px] text-slate-500">
            <span class="inline-flex items-center gap-1.5"><span class="w-2.5 h-0.5 bg-indigo-600"></span>Active basket</span>
            <span class="inline-flex items-center gap-1.5"><span class="w-2.5 h-0.5 border-t border-dashed border-slate-400"></span>Smallcap 250</span>
          </div>
        </div>
        ${renderActiveChart(sim, niftyOn)}
      </div>

      <!-- Stats grid -->
      <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
        ${statTile("Trades", String(stats.tradeCount), `${stats.buyCount} BUY · ${stats.sellCount} SELL`)}
        ${statTile("Hit rate", hitRateStr, `${stats.sellCount} closed`)}
        ${statTile("Avg hold", avgHoldStr, "per closed trade")}
        ${statTile("Avg winner", avgWinStr, "")}
        ${statTile("Avg loser", avgLossStr, "")}
        ${statTile("Max drawdown", `${stats.maxDD.toFixed(2)}%`, `Live ${stats.liveHoldings} stocks`)}
      </div>

      <!-- Current holdings + recent trades, side by side on lg -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
        ${renderActiveHoldings(sim, niftyOn)}
        ${renderActiveTrades(sim)}
      </div>

      <!-- Beta caveat -->
      <div class="text-xs text-slate-500 bg-slate-50 ring-1 ring-slate-200 rounded-xl p-3 leading-relaxed">
        <strong>Beta caveat</strong> — short backtest window (${stats.days} snapshot days starting ${fmtDateDMY(stats.startDate)}). Numbers will firm up as the snapshot trail grows.
        Buys + sells transact at the same EOD close that marks the portfolio to market, so entry-day P&L is zero by construction. No brokerage, slippage, or STT modelled yet.
      </div>
    </div>
  `;
}

function renderActiveChart(sim, niftyOn) {
  const W = 800, H = 220;
  const M = { left: 44, right: 16, top: 14, bottom: 30 };
  const innerW = W - M.left - M.right;
  const innerH = H - M.top - M.bottom;

  const pts = sim.equity.map((e, i) => {
    const ret = (e.value / sim.startCapital - 1) * 100;
    const nClose = niftyOn ? niftyOn(e.date) : null;
    return { date: e.date, ret, nClose };
  });
  // Nifty cumulative anchored at first available close
  const nFirst = pts.find((p) => p.nClose != null)?.nClose || null;
  for (const p of pts) p.niftyRet = (nFirst != null && p.nClose != null) ? (p.nClose / nFirst - 1) * 100 : null;

  const allVals = [];
  for (const p of pts) { allVals.push(p.ret); if (p.niftyRet != null) allVals.push(p.niftyRet); }
  const yMin = Math.min(0, ...allVals);
  const yMax = Math.max(0, ...allVals);
  const ySpan = Math.max(yMax - yMin, 0.5);
  const yLo = yMin - ySpan * 0.12, yHi = yMax + ySpan * 0.12;
  const xAt = (i) => M.left + (pts.length <= 1 ? 0 : (i / (pts.length - 1)) * innerW);
  const yAt = (v) => M.top + (1 - (v - yLo) / (yHi - yLo)) * innerH;

  const activePath = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${xAt(i).toFixed(2)} ${yAt(p.ret).toFixed(2)}`).join(" ");
  const niftyPts = pts.map((p, i) => p.niftyRet == null ? null : [xAt(i), yAt(p.niftyRet)]).filter(Boolean);
  const niftyPath = niftyPts.length >= 2 ? niftyPts.map(([x, y], i) => `${i === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`).join(" ") : "";

  const first = [xAt(0), yAt(pts[0].ret)];
  const last = [xAt(pts.length - 1), yAt(pts[pts.length - 1].ret)];
  const baseY = yAt(0).toFixed(2);
  const areaPath = `${activePath} L ${last[0].toFixed(2)} ${baseY} L ${first[0].toFixed(2)} ${baseY} Z`;

  const yTicks = [0, 1, 2, 3, 4].map((step) => {
    const v = yLo + (yHi - yLo) * (step / 4);
    const yy = (M.top + innerH - (step / 4) * innerH).toFixed(2);
    const isZero = Math.abs(v) < 0.05;
    return `
      <line x1="${M.left}" x2="${(W - M.right).toFixed(2)}" y1="${yy}" y2="${yy}" stroke="${isZero ? "#94a3b8" : "#e2e8f0"}" stroke-width="${isZero ? 0.9 : 0.6}" stroke-dasharray="${isZero ? "0" : "3 4"}" />
      <text x="${(M.left - 8).toFixed(2)}" y="${yy}" text-anchor="end" dominant-baseline="middle" font-size="10" font-weight="500" fill="#94a3b8">${v >= 0 ? "+" : ""}${v.toFixed(1)}%</text>`;
  }).join("");

  const tickEvery = Math.max(1, Math.ceil(pts.length / 6));
  const xTicks = pts.map((p, i) => (i % tickEvery !== 0 && i !== pts.length - 1)
    ? ""
    : `<text x="${xAt(i).toFixed(2)}" y="${(M.top + innerH + 16).toFixed(2)}" text-anchor="middle" font-size="10" fill="#64748b">${fmtDateDM(p.date)}</text>`
  ).join("");

  return `
    <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" class="w-full select-none" style="max-height:260px">
      <defs>
        <linearGradient id="activeArea" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#6366f1" stop-opacity="0.22"/>
          <stop offset="100%" stop-color="#6366f1" stop-opacity="0"/>
        </linearGradient>
      </defs>
      ${yTicks}
      <path d="${areaPath}" fill="url(#activeArea)" />
      ${niftyPath ? `<path d="${niftyPath}" fill="none" stroke="#94a3b8" stroke-width="1.6" stroke-dasharray="4 4" stroke-linecap="round" stroke-linejoin="round" />` : ""}
      <path d="${activePath}" fill="none" stroke="#6366f1" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round" />
      ${xTicks}
    </svg>
  `;
}

function renderActiveHoldings(sim, niftyOn) {
  if (!sim.equity.length) return "";
  const today = sim.equity[sim.equity.length - 1].date;
  const todaySnap = state.cache.history.snapshots.find((s) => s.date === today);
  const closeByTicker = {};
  if (todaySnap) for (const s of todaySnap.stocks) if (typeof s.close === "number") closeByTicker[s.ticker] = s.close;

  if (sim.holdings.size === 0) {
    return `
      <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
        <h3 class="font-semibold text-slate-900 text-sm mb-2">Live holdings</h3>
        <div class="text-xs text-slate-500">No active positions today — basket sat in cash.</div>
      </div>`;
  }

  const rows = [...sim.holdings.entries()]
    .map(([ticker, pos]) => {
      const px = closeByTicker[ticker] ?? pos.entryPrice;
      const ret = (px / pos.entryPrice - 1) * 100;
      const days = daysBetween(pos.entryDate, today);
      return { ticker, name: pos.name, sector: pos.sector, entryDate: pos.entryDate, entryPrice: pos.entryPrice, px, ret, days };
    })
    .sort((a, b) => b.ret - a.ret);

  const body = rows.map((r) => `
    <tr class="border-b border-slate-100 last:border-0">
      <td class="py-2 pr-3">
        <div class="font-semibold text-slate-900 text-sm">${escapeHtml(r.ticker)}</div>
        <div class="text-[11px] text-slate-500 truncate max-w-[180px]">${escapeHtml(r.name)}</div>
      </td>
      <td class="py-2 px-2 text-right text-xs text-slate-600 tabular-nums">${fmtDateDM(r.entryDate)}</td>
      <td class="py-2 px-2 text-right text-xs text-slate-600 tabular-nums">₹${formatPrice(r.entryPrice)}</td>
      <td class="py-2 px-2 text-right text-xs text-slate-700 tabular-nums">₹${formatPrice(r.px)}</td>
      <td class="py-2 pl-2 text-right text-sm font-bold tabular-nums ${r.ret >= 0 ? "text-emerald-600" : "text-rose-600"}">${r.ret >= 0 ? "+" : ""}${r.ret.toFixed(2)}%<div class="text-[10px] text-slate-400 font-normal">${r.days}d</div></td>
    </tr>`).join("");

  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
      <div class="flex items-center justify-between mb-3">
        <h3 class="font-semibold text-slate-900 text-sm">Live holdings</h3>
        <span class="text-[11px] text-slate-500">${rows.length} positions</span>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-xs">
          <thead class="text-[10px] font-bold uppercase tracking-wider text-slate-500">
            <tr class="border-b border-slate-200">
              <th class="py-2 pr-3 text-left">Stock</th>
              <th class="py-2 px-2 text-right">Entry</th>
              <th class="py-2 px-2 text-right">Buy ₹</th>
              <th class="py-2 px-2 text-right">Now ₹</th>
              <th class="py-2 pl-2 text-right">Return</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </div>
  `;
}

function renderActiveTrades(sim) {
  // Show last 20 trades, newest first. Closed trades (SELL) show realized
  // return; opens (BUY) just show entry.
  const recent = [...sim.trades].slice(-20).reverse();
  if (!recent.length) {
    return `
      <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
        <h3 class="font-semibold text-slate-900 text-sm mb-2">Recent trades</h3>
        <div class="text-xs text-slate-500">No trades yet.</div>
      </div>`;
  }
  const body = recent.map((t) => {
    const isSell = t.action === "SELL";
    const actionPill = isSell
      ? `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold ${t.ret >= 0 ? "bg-emerald-100 text-emerald-700" : "bg-rose-100 text-rose-700"}">SELL</span>`
      : `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-indigo-100 text-indigo-700">BUY</span>`;
    const retCell = isSell
      ? `<span class="font-bold tabular-nums ${t.ret >= 0 ? "text-emerald-600" : "text-rose-600"}">${t.ret >= 0 ? "+" : ""}${t.ret.toFixed(2)}%</span><div class="text-[10px] text-slate-400">${t.days}d held</div>`
      : `<span class="text-slate-400 text-xs">opened</span>`;
    return `
      <tr class="border-b border-slate-100 last:border-0">
        <td class="py-2 pr-2 text-xs text-slate-600 tabular-nums">${fmtDateDM(t.date)}</td>
        <td class="py-2 px-2">${actionPill}</td>
        <td class="py-2 px-2">
          <div class="font-semibold text-slate-900 text-sm">${escapeHtml(t.ticker)}</div>
          <div class="text-[11px] text-slate-500 truncate max-w-[150px]">${escapeHtml(t.name)}</div>
        </td>
        <td class="py-2 px-2 text-right text-xs text-slate-700 tabular-nums">₹${formatPrice(t.price)}</td>
        <td class="py-2 pl-2 text-right">${retCell}</td>
      </tr>`;
  }).join("");

  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
      <div class="flex items-center justify-between mb-3">
        <h3 class="font-semibold text-slate-900 text-sm">Recent trades</h3>
        <span class="text-[11px] text-slate-500">latest ${recent.length} of ${sim.trades.length}</span>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-xs">
          <thead class="text-[10px] font-bold uppercase tracking-wider text-slate-500">
            <tr class="border-b border-slate-200">
              <th class="py-2 pr-2 text-left">Date</th>
              <th class="py-2 px-2 text-left">Action</th>
              <th class="py-2 px-2 text-left">Stock</th>
              <th class="py-2 px-2 text-right">Price</th>
              <th class="py-2 pl-2 text-right">Return</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </div>
  `;
}

function daysBetween(a, b) {
  const ma = new Date(a + "T00:00:00Z").getTime();
  const mb = new Date(b + "T00:00:00Z").getTime();
  return Math.max(0, Math.round((mb - ma) / 86400000));
}

function formatPrice(n) {
  if (n == null || !isFinite(n)) return "—";
  if (n >= 1000) return n.toLocaleString("en-IN", { maximumFractionDigits: 0 });
  return n.toFixed(2);
}

// History drill — premium per-company chart modal.
// Layout (top to bottom):
//   - Company header (avatar / name / sector / ticker)
//   - "We said vs Now" 3-card hero strip
//   - Chart card:
//       Title row     · "Price · last N days"  | inline rating legend
//       Price line    · with rating-colored markers (halo on rating change)
//       Rating tape   · one colored cell per snapshot day (clearer than
//                       the old floating sparkbars)
//       Date labels
//   - Hover crosshair + floating tooltip card that follows the cursor
//     and snaps to the nearest data point
const RATING_FILL = {
  "STRONG BUY": "#10b981", "BUY": "#3b82f6", "WATCH": "#f59e0b", "AVOID": "#f43f5e",
  "FILTERED":   "#64748b", "UNRATED": "#cbd5e1",
};

function openHistoryDrill(pick) {
  if (!pick) return;
  const { color, initials } = avatarFor(pick.name || "—");
  const todayDate = state.cache.history.idx.dates[state.cache.history.idx.dates.length - 1];

  const points = pick.points.filter((p) => typeof p.close === "number");
  // Was < 2, which silently did nothing on the FIRST day of a cohort --
  // exactly when the basket is locked in and you most want to inspect a
  // pick. The real constraint was xAt() dividing by (points.length - 1);
  // that is now guarded, so a single point renders a legitimate one-day
  // view: entry price, composite, pillars, target and stop-loss bands. The
  // price line needs two points and simply has nothing to draw until then.
  if (points.length < 1) return;
  const singlePoint = points.length === 1;

  // --- SVG geometry ---
  // Single viewBox; everything inside is sized in viewBox units, the
  // <svg> itself scales responsively via w-full.
  const W = 820, H = 260;
  const M = { top: 14, right: 24, bottom: 56, left: 64 };     // outer margins
  const TAPE_H = 9;                                           // rating-tape strip height
  const TAPE_GAP = 8;                                         // gap between line chart and tape
  const innerW = W - M.left - M.right;
  const innerH = H - M.top - M.bottom - TAPE_H - TAPE_GAP;
  const TAPE_Y = M.top + innerH + TAPE_GAP;
  const X_LABEL_Y = TAPE_Y + TAPE_H + 18;

  // Target / Stop-loss levels for the overlay. LKP picks carry explicit
  // tgt1 + sl from the client upload; AI / cohort picks use the framework's
  // uniform +5% / −20% bands around the entry close. Either side can be
  // missing (e.g. cohort pick built from snapshot trail with no entry
  // close) — we render whichever line we can.
  let targetPrice = null, slPrice = null, levelAnchor = null;
  if (pick.isLkp) {
    targetPrice = typeof pick.tgt1 === "number" ? pick.tgt1 : null;
    slPrice = typeof pick.sl === "number" ? pick.sl : null;
    // Match the Accuracy row's entry reference: snapshot close at the
    // cohort anchor date (the same close buildAccuracyData uses for
    // computing targetPct / slPct). Falls back to the LKP entry midpoint
    // when the cohort anchor isn't in this ticker's snapshot trail.
    const cohortAnchor = state.cache.history?.cohortAnchor;
    if (cohortAnchor && Array.isArray(pick.points)) {
      const anchorPt = pick.points.find((p) => p.date === cohortAnchor && typeof p.close === "number")
        || pick.points.find((p) => p.date >= cohortAnchor && typeof p.close === "number");
      if (anchorPt) levelAnchor = anchorPt.close;
    }
    if (levelAnchor == null && typeof pick.entry === "number") levelAnchor = pick.entry;
  } else if (typeof pick.firstSBClose === "number") {
    targetPrice = pick.firstSBClose * (1 + AI_TARGET_PCT);
    slPrice = pick.firstSBClose * (1 - AI_SL_PCT);
    levelAnchor = pick.firstSBClose;
  }

  const closes = points.map((p) => p.close);
  let yMin = Math.min(...closes), yMax = Math.max(...closes);
  // Expand the range so the levels stay visible even when they fall
  // outside the realized price band.
  if (targetPrice != null) yMax = Math.max(yMax, targetPrice);
  if (slPrice != null) yMin = Math.min(yMin, slPrice);
  const ySpan = Math.max(yMax - yMin, 1);
  const yPad = ySpan * 0.1;
  const yLo = yMin - yPad, yHi = yMax + yPad;

  // Centre a lone point instead of dividing by zero.
  const xAt = (i) => M.left + (points.length > 1 ? i / (points.length - 1) : 0.5) * innerW;
  const yAt = (c) => M.top + innerH - ((c - yLo) / (yHi - yLo)) * innerH;

  // Path through closes
  const pathD = points.map((p, i) => `${i === 0 ? "M" : "L"} ${xAt(i).toFixed(2)} ${yAt(p.close).toFixed(2)}`).join(" ");
  const areaD = `${pathD} L ${xAt(points.length - 1).toFixed(2)} ${(M.top + innerH).toFixed(2)} L ${M.left.toFixed(2)} ${(M.top + innerH).toFixed(2)} Z`;

  // Y-axis price gridlines + labels (4 steps)
  const yTicks = [0, 1, 2, 3, 4].map((step) => {
    const price = yLo + (yHi - yLo) * (step / 4);
    const yy = (M.top + innerH - (step / 4) * innerH).toFixed(2);
    return `
      <line x1="${M.left}" x2="${(W - M.right).toFixed(2)}" y1="${yy}" y2="${yy}" stroke="#e2e8f0" stroke-width="0.7" stroke-dasharray="3 4" />
      <text x="${(M.left - 10).toFixed(2)}" y="${yy}" text-anchor="end" dominant-baseline="middle" font-size="10.5" font-weight="500" fill="#94a3b8">₹${formatPrice(price)}</text>
    `;
  }).join("");

  // Horizontal level lines for target / stop-loss. Sit between the
  // chart line and the rating markers — visible without covering the
  // rating-change halos. Labels right-aligned at the chart edge with
  // the % return relative to entry (computed from levelAnchor when
  // available; otherwise the price alone).
  function levelOverlay(price, color, sign, label, labelOffsetY) {
    if (price == null) return "";
    const y = yAt(price);
    const retTxt = levelAnchor ? ` · ${sign}${Math.abs((price / levelAnchor - 1) * 100).toFixed(1)}%` : "";
    const fullLabel = `${label} ₹${formatPrice(price)}${retTxt}`;
    return `
      <line x1="${M.left}" x2="${(W - M.right).toFixed(2)}" y1="${y.toFixed(2)}" y2="${y.toFixed(2)}" stroke="${color}" stroke-width="1.2" stroke-dasharray="6 4" opacity="0.85" />
      <rect x="${(W - M.right - 110).toFixed(2)}" y="${(y + labelOffsetY - 11).toFixed(2)}" width="106" height="14" rx="3" fill="#fff" fill-opacity="0.92" />
      <text x="${(W - M.right - 6).toFixed(2)}" y="${(y + labelOffsetY).toFixed(2)}" text-anchor="end" font-size="10" font-weight="700" fill="${color}">${fullLabel}</text>`;
  }
  const targetLine = levelOverlay(targetPrice, "#059669", "+", "Target", -4);
  const slLine = levelOverlay(slPrice, "#e11d48", "−", "SL", 12);

  // Markers — rating-change points get a halo + ring, same-rating days
  // get a tiny dot. Visual hierarchy makes the explainer footer redundant.
  let prevRating = null;
  const markers = points.map((p, i) => {
    const fill = RATING_FILL[p.rating] || "#cbd5e1";
    const changed = p.rating && p.rating !== prevRating;
    prevRating = p.rating;
    if (changed) {
      return `
        <circle cx="${xAt(i).toFixed(2)}" cy="${yAt(p.close).toFixed(2)}" r="10" fill="${fill}" fill-opacity="0.18" />
        <circle cx="${xAt(i).toFixed(2)}" cy="${yAt(p.close).toFixed(2)}" r="5.5" fill="${fill}" stroke="#fff" stroke-width="2" />
      `;
    }
    return `<circle cx="${xAt(i).toFixed(2)}" cy="${yAt(p.close).toFixed(2)}" r="2.6" fill="#fff" stroke="${fill}" stroke-width="1.5" />`;
  }).join("");

  // Rating tape — one colored cell per snapshot day. Shows rating
  // evolution at a glance under the chart, where the floating sparkbars
  // used to be.
  const cellW = innerW / points.length;
  const tapeCells = points.map((p, i) => {
    const fill = RATING_FILL[p.rating] || "#cbd5e1";
    const x = M.left + i * cellW;
    return `<rect x="${x.toFixed(2)}" y="${TAPE_Y}" width="${Math.max(cellW - 1, 1).toFixed(2)}" height="${TAPE_H}" fill="${fill}" fill-opacity="0.92" rx="1.5" />`;
  }).join("");

  // X-axis date labels — at most 7 to keep the row clean
  const tickEvery = Math.max(1, Math.ceil(points.length / 7));
  const xTicks = points.map((p, i) => {
    if (i % tickEvery !== 0 && i !== points.length - 1) return "";
    return `<text x="${xAt(i).toFixed(2)}" y="${X_LABEL_Y}" text-anchor="middle" font-size="10.5" font-weight="500" fill="#64748b">${p.date.slice(5)}</text>`;
  }).join("");

  // Hover guides — initially invisible, JS toggles opacity on mousemove
  const hoverLayer = `
    <line id="hist-guide" x1="0" y1="${M.top}" x2="0" y2="${(TAPE_Y + TAPE_H).toFixed(2)}" stroke="#94a3b8" stroke-width="0.8" stroke-dasharray="2 3" opacity="0" />
    <circle id="hist-hover-halo" cx="0" cy="0" r="14" fill="#6366f1" fill-opacity="0.16" opacity="0" />
    <circle id="hist-hover-point" cx="0" cy="0" r="6" fill="#fff" stroke="#6366f1" stroke-width="2.5" opacity="0" />
    <rect id="hist-hover-capture" x="0" y="0" width="${W}" height="${H}" fill="transparent" />
  `;

  // Rating legend (replaces "big rings = ..." footer)
  const legendRatings = ["STRONG BUY", "BUY", "WATCH", "AVOID", "FILTERED"];
  const legend = legendRatings.map((r) => `
    <span class="inline-flex items-center gap-1.5 text-[11px] text-slate-600">
      <span class="w-2.5 h-2.5 rounded-full" style="background:${RATING_FILL[r]}"></span>${r}
    </span>
  `).join("");

  // Hero strip values
  const ret = pick.ret;
  const retCls = ret == null ? "text-slate-500" : ret >= 0 ? "text-emerald-700" : "text-rose-700";
  const retBg = ret == null ? "from-slate-50 to-white ring-slate-200" : ret >= 0 ? "from-emerald-50 to-teal-50 ring-emerald-200" : "from-rose-50 to-pink-50 ring-rose-200";

  // LKP picks reuse this modal but with a client-entry framing: no "we said
  // STRONG BUY" anchor, and a targets/SL block instead of the pillar forensics.
  // Cohort-row clicks (AI table) when there's no prior STRONG BUY use a
  // "Cohort entry" framing so the header isn't misleading.
  const isLkp = !!pick.isLkp;
  const isCohortOnly = !!pick.isCohortLookup;
  const entryCardHtml = isLkp
    ? `<div class="rounded-xl ring-1 ring-indigo-200 bg-gradient-to-br from-indigo-50 to-blue-50 px-3 py-2">
         <div class="text-[9px] font-bold uppercase tracking-wider text-indigo-700">Client entry</div>
         <div class="text-sm font-display font-bold text-slate-900 mt-0.5 tabular-nums">₹${pick.entry_low}–${pick.entry_high}</div>
         <div class="text-[11px] text-slate-600 mt-0.5">midpoint ₹${formatPrice(pick.entry)}</div>
       </div>`
    : isCohortOnly
      ? `<div class="rounded-xl ring-1 ring-indigo-200 bg-gradient-to-br from-indigo-50 to-blue-50 px-3 py-2">
           <div class="text-[9px] font-bold uppercase tracking-wider text-indigo-700">Cohort entry</div>
           <div class="text-sm font-display font-bold text-slate-900 mt-0.5">${fmtDateDMY(pick.firstSBDate)} · <span class="tabular-nums">₹${formatPrice(pick.firstSBClose)}</span></div>
           ${pick.firstSBComposite != null ? `<div class="text-[11px] text-slate-600 mt-0.5">composite <span class="font-bold tabular-nums">${pick.firstSBComposite.toFixed(1)}</span></div>` : ""}
         </div>`
      : `<div class="rounded-xl ring-1 ring-emerald-200 bg-gradient-to-br from-emerald-50 to-teal-50 px-3 py-2">
           <div class="text-[9px] font-bold uppercase tracking-wider text-emerald-700">We said STRONG BUY</div>
           <div class="text-sm font-display font-bold text-slate-900 mt-0.5">${fmtDateDMY(pick.firstSBDate)} · <span class="tabular-nums">₹${formatPrice(pick.firstSBClose)}</span></div>
           ${pick.firstSBComposite != null ? `<div class="text-[11px] text-slate-600 mt-0.5">composite <span class="font-bold tabular-nums">${pick.firstSBComposite.toFixed(1)}</span></div>` : ""}
         </div>`;
  const retLabelHtml = isLkp ? "Return vs entry" : isCohortOnly ? "Return since cohort entry" : "Realized return";
  const retSubHtml = isLkp ? "vs entry midpoint" : `over ${pick.days} day${pick.days === 1 ? "" : "s"}`;
  const midBlockHtml = isLkp ? renderLkpTargets(pick) : renderScoreForensics(pick);

  // Peak excursion since entry — same numbers as the per-pick row
  // nuggets, expanded here with the exact peak / trough dates. Walks
  // this ticker's price trail forward from the entry anchor.
  const exEntryDate = pick.firstSBDate || state.cache.history?.cohortAnchor || points[0].date;
  const exEntryPrice = (typeof pick.firstSBClose === "number") ? pick.firstSBClose
    : (typeof levelAnchor === "number" && levelAnchor) ? levelAnchor
    : points[0].close;
  let exUp = null, exUpD = null, exDn = null, exDnD = null;
  for (const p of points) {
    if (p.date < exEntryDate || p.date > todayDate) continue;
    const pct = (p.close / exEntryPrice - 1) * 100;
    if (exUp == null || pct > exUp) { exUp = pct; exUpD = p.date; }
    if (exDn == null || pct < exDn) { exDn = pct; exDnD = p.date; }
  }
  const excursionHtml = exUp != null ? `
      <div class="grid grid-cols-2 gap-2 mb-2">
        <div class="rounded-xl ring-1 ring-emerald-200 bg-gradient-to-br from-emerald-50 to-teal-50 px-3 py-2">
          <div class="text-[9px] font-bold uppercase tracking-wider text-emerald-700">Max upside since entry</div>
          <div class="text-sm font-display font-bold text-slate-900 mt-0.5 tabular-nums">${exUp >= 0 ? "+" : ""}${exUp.toFixed(2)}% <span class="text-[11px] font-semibold text-slate-500">· ${daysBetween(exEntryDate, exUpD)}d</span></div>
          <div class="text-[11px] text-slate-600 mt-0.5">peak on ${fmtDateDMY(exUpD)}</div>
        </div>
        <div class="rounded-xl ring-1 ring-rose-200 bg-gradient-to-br from-rose-50 to-pink-50 px-3 py-2">
          <div class="text-[9px] font-bold uppercase tracking-wider text-rose-700">Max drawdown since entry</div>
          <div class="text-sm font-display font-bold text-slate-900 mt-0.5 tabular-nums">${exDn >= 0 ? "+" : ""}${exDn.toFixed(2)}% <span class="text-[11px] font-semibold text-slate-500">· ${daysBetween(exEntryDate, exDnD)}d</span></div>
          <div class="text-[11px] text-slate-600 mt-0.5">low on ${fmtDateDMY(exDnD)}</div>
        </div>
      </div>` : "";

  // Reason panel for currently-FILTERED stocks. Snapshots store only the
  // hardFailed boolean (not which rules fired), so we render a placeholder
  // immediately and asynchronously fetch the live composite for this ticker
  // to fill in the failing rule names below. The panel carries the
  // ticker as data-ticker so a stale async response from a previous
  // drill can't overwrite the current one.
  const isFiltered = pick.currentRating === "FILTERED";
  const hardFailPanelHtml = isFiltered ? `
    <div id="hist-hardfail-panel" data-ticker="${escapeHtml(pick.ticker)}" class="rounded-xl ring-1 ring-rose-200 bg-gradient-to-br from-rose-50 to-pink-50 px-3 py-2.5 mb-2">
      <div class="flex items-start gap-2">
        <div class="text-rose-600 text-base leading-none">⚠</div>
        <div class="flex-1 min-w-0">
          <div class="font-bold text-rose-900 text-sm">Currently auto-excluded from basket</div>
          <div class="text-[11px] text-rose-700/90 mt-0.5">Hard-failed per client framework. Held in the cohort per the held-basket rule, but the strategy itself would have exited.</div>
          <div id="hist-hardfail-rules" class="text-[11px] text-rose-700/80 mt-1.5">Checking which rule fired…</div>
        </div>
      </div>
    </div>` : "";

  openModal(`
    <div class="px-5 py-3">
      <div class="flex items-start justify-between gap-4 mb-3">
        <div class="flex items-center gap-3 min-w-0">
          <div class="w-11 h-11 rounded-xl bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-sm shadow-md flex-shrink-0">${initials}</div>
          <div class="min-w-0">
            <div class="font-display font-bold text-slate-900 text-lg leading-tight truncate">${escapeHtml(pick.name || "—")}</div>
            <div class="text-xs text-slate-500 mt-0.5 truncate">${escapeHtml(pick.sector || "")} · ${escapeHtml(pick.ticker)}</div>
          </div>
        </div>
        <button id="modal-close-btn" class="text-slate-400 hover:text-slate-700 text-2xl leading-none">×</button>
      </div>

      <div class="grid grid-cols-1 sm:grid-cols-3 gap-2 mb-2">
        ${entryCardHtml}
        <div class="rounded-xl ring-1 ring-slate-200 bg-gradient-to-br from-slate-50 to-white px-3 py-2">
          <div class="text-[9px] font-bold uppercase tracking-wider text-slate-500">Today</div>
          <div class="text-sm font-display font-bold text-slate-900 mt-0.5">${fmtDateDMY(todayDate)} · <span class="tabular-nums">₹${formatPrice(pick.todayClose)}</span></div>
          ${pick.currentRating ? `<div class="text-[11px] text-slate-600 mt-0.5"><span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 ${composite.ratingClass(pick.currentRating)}">${escapeHtml(pick.currentRating)}</span></div>` : ""}
        </div>
        <div class="rounded-xl ring-1 bg-gradient-to-br ${retBg} px-3 py-2 flex items-center justify-between gap-3">
          <div>
            <div class="text-[9px] font-bold uppercase tracking-wider ${retCls}">${retLabelHtml}</div>
            <div class="text-[11px] text-slate-600 mt-0.5">${retSubHtml}</div>
          </div>
          <div class="text-2xl font-display font-extrabold tabular-nums ${retCls} leading-none">${ret == null ? "—" : (ret >= 0 ? "+" : "") + ret.toFixed(2) + "%"}</div>
        </div>
      </div>

      ${excursionHtml}
      ${hardFailPanelHtml}
      ${midBlockHtml}

      <div class="rounded-2xl ring-1 ring-slate-200 bg-white px-3 py-2.5 sm:px-4 sm:py-3">
        <div class="flex flex-wrap items-center justify-between gap-2 mb-1.5">
          <div class="flex items-baseline gap-2">
            <div class="font-display font-bold text-slate-900 text-sm">Price &amp; rating timeline</div>
            <div class="text-[11px] text-slate-500">${points.length} days · ${points[0].date.slice(5)} → ${points[points.length - 1].date.slice(5)}</div>
          </div>
          <div class="flex flex-wrap items-center gap-x-3 gap-y-1">${legend}</div>
        </div>

        <div id="hist-chart-container" class="relative">
          <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" class="w-full select-none" style="max-height:280px">
            <defs>
              <linearGradient id="histArea" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stop-color="#6366f1" stop-opacity="0.22"/>
                <stop offset="100%" stop-color="#6366f1" stop-opacity="0"/>
              </linearGradient>
            </defs>
            ${yTicks}
            <path d="${areaD}" fill="url(#histArea)"/>
            ${targetLine}
            ${slLine}
            <path d="${pathD}" fill="none" stroke="#6366f1" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
            ${markers}
            ${tapeCells}
            ${xTicks}
            <text x="${(M.left - 10).toFixed(2)}" y="${(TAPE_Y + TAPE_H / 2).toFixed(2)}" text-anchor="end" dominant-baseline="middle" font-size="9" font-weight="700" fill="#94a3b8" letter-spacing="0.5">RATING</text>
            ${hoverLayer}
          </svg>
          <div id="hist-tooltip" class="hidden absolute z-10 pointer-events-none -translate-x-1/2 -translate-y-[calc(100%+14px)] bg-slate-900/95 backdrop-blur text-white text-xs rounded-xl shadow-2xl ring-1 ring-slate-700/60 px-3 py-2 whitespace-nowrap"></div>
        </div>
      </div>
    </div>
  `, { size: "magazine" });

  $("#modal-close-btn")?.addEventListener("click", closeModal);
  $("#modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "modal-overlay") closeModal();
  }, { once: true });

  // Async: when this stock is currently FILTERED, fetch the live composite
  // for its ticker so we can show which specific rule(s) tripped the
  // hard-fail. Snapshots only store the boolean — the rule list comes
  // from a fresh scoreCompositeOne() call.
  if (isFiltered) {
    populateHardFailRules(pick.ticker).catch(() => {
      const panel = $("#hist-hardfail-panel");
      const el = $("#hist-hardfail-rules");
      if (el && panel && panel.dataset.ticker === pick.ticker) {
        el.textContent = "Couldn't load live rule list — open AI Basket for the full breakdown.";
      }
    });
  }

  // Wire hover crosshair + tooltip. The capture <rect> spans the chart
  // + rating-tape band so hovering anywhere reads the same nearest day.
  const container = $("#hist-chart-container");
  const svg = container?.querySelector("svg");
  const capture = $("#hist-hover-capture");
  const guide = $("#hist-guide");
  const hoverPt = $("#hist-hover-point");
  const halo = $("#hist-hover-halo");
  const tip = $("#hist-tooltip");
  if (!container || !svg || !capture || !tip) return;

  function show(idx) {
    const p = points[idx];
    const px = xAt(idx);
    const py = yAt(p.close);
    guide.setAttribute("x1", px); guide.setAttribute("x2", px); guide.setAttribute("opacity", "1");
    hoverPt.setAttribute("cx", px); hoverPt.setAttribute("cy", py); hoverPt.setAttribute("opacity", "1");
    halo.setAttribute("cx", px); halo.setAttribute("cy", py); halo.setAttribute("opacity", "1");

    // viewBox → screen via the live CTM, so the tooltip tracks the point
    // even when preserveAspectRatio letterboxes the chart or the page is
    // zoomed (rect.width / W silently drifts in those cases).
    const contRect = container.getBoundingClientRect();
    const ctm = svg.getScreenCTM();
    let tipX, tipY;
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = px; sp.y = py;
      const scr = sp.matrixTransform(ctm);
      tipX = scr.x - contRect.left; tipY = scr.y - contRect.top;
    } else {
      const rect = svg.getBoundingClientRect();
      tipX = px * (rect.width / W); tipY = py * (rect.height / H);
    }

    const ratingColor = RATING_FILL[p.rating] || "#cbd5e1";
    const compTxt = p.composite != null ? p.composite.toFixed(1) : "—";
    const first = points[0];
    const chg = ((p.close - first.close) / first.close) * 100;
    const chgCls = chg >= 0 ? "text-emerald-300" : "text-rose-300";
    tip.innerHTML = `
      <div class="font-bold text-sm leading-tight">${p.date}</div>
      <div class="mt-1 flex items-center gap-2">
        <span class="text-base font-extrabold tabular-nums">₹${formatPrice(p.close)}</span>
        <span class="text-[10px] tabular-nums ${chgCls}">${chg >= 0 ? "+" : ""}${chg.toFixed(2)}%</span>
      </div>
      <div class="mt-1.5 flex items-center gap-2 text-[11px]">
        <span class="inline-flex items-center gap-1">
          <span class="w-2 h-2 rounded-full" style="background:${ratingColor}"></span>
          <span class="font-semibold">${p.rating || "—"}</span>
        </span>
        <span class="text-slate-300">·</span>
        <span>composite <span class="font-bold tabular-nums">${compTxt}</span></span>
      </div>
    `;
    tip.classList.remove("hidden");
    // Position then flip horizontally if too close to container edges.
    tip.style.left = `${tipX}px`;
    tip.style.top = `${tipY}px`;
    tip.style.transform = "translate(-50%, calc(-100% - 14px))";
    requestAnimationFrame(() => {
      const tr = tip.getBoundingClientRect();
      const cr = container.getBoundingClientRect();
      let dx = 0;
      if (tr.left < cr.left + 6) dx = (cr.left + 6) - tr.left;
      else if (tr.right > cr.right - 6) dx = (cr.right - 6) - tr.right;
      if (dx !== 0) tip.style.transform = `translate(calc(-50% + ${dx}px), calc(-100% - 14px))`;
    });
  }
  function hide() {
    guide.setAttribute("opacity", "0");
    hoverPt.setAttribute("opacity", "0");
    halo.setAttribute("opacity", "0");
    tip.classList.add("hidden");
  }

  function eventToIdx(e) {
    const t = e.touches ? e.touches[0] : e;
    let xView;
    const ctm = svg.getScreenCTM();
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = t.clientX; sp.y = t.clientY;
      xView = sp.matrixTransform(ctm.inverse()).x;   // letterbox / zoom safe
    } else {
      const rect = svg.getBoundingClientRect();
      xView = ((t.clientX - rect.left) / rect.width) * W;
    }
    let bestI = 0, bestD = Infinity;
    for (let i = 0; i < points.length; i++) {
      const d = Math.abs(xAt(i) - xView);
      if (d < bestD) { bestD = d; bestI = i; }
    }
    return bestI;
  }

  capture.addEventListener("mousemove", (e) => show(eventToIdx(e)));
  capture.addEventListener("mouseleave", hide);
  capture.addEventListener("touchstart", (e) => { show(eventToIdx(e)); e.preventDefault(); }, { passive: false });
  capture.addEventListener("touchmove",  (e) => { show(eventToIdx(e)); e.preventDefault(); }, { passive: false });
  capture.addEventListener("touchend", hide);

  // Surface today (last point) on open so the panel isn't empty.
  show(points.length - 1);
}

// ---------------- filtering / sorting ----------------
function applyFilters() {
  const c = cfg(); const st = tabState();
  const q = state.search.trim().toLowerCase();
  let rows = st.scored.filter((s) => {
    if (state.watchOnly && !state.watchlist.has(companyKey(s.company))) return false;
    if (q && !c.name(s.company).toLowerCase().includes(q)) return false;
    if (state.scoreFilter === "redflag") return !!s.isRedFlag;
    if (state.scoreFilter === "belowtrend") return s.hardFails.length > 0 && !s.isRedFlag;
    if (state.scoreFilter !== "all") {
      const tier = s.hardFails.length ? "hardfail" : scoreTier(s.scorePct);
      if (tier !== state.scoreFilter) return false;
    }
    return true;
  });
  rows.sort((a, b) => {
    const dir = state.sortDir === "asc" ? 1 : -1;
    let av, bv;
    if (state.sortBy === "score") { av = a.totalPoints; bv = b.totalPoints; }
    else if (state.sortBy === "name") { av = c.name(a.company).toLowerCase(); bv = c.name(b.company).toLowerCase(); }
    if (av < bv) return -1 * dir; if (av > bv) return 1 * dir; return 0;
  });
  st.filtered = rows;
  renderTable();
}

function renderTable() {
  const c = cfg(); const st = tabState();
  const rows = st.filtered;
  $("#row-count").textContent = `${rows.length} of ${st.scored.length}`;
  // dynamic header
  $("#table-head").innerHTML = `
    <tr class="text-left text-xs font-bold uppercase tracking-wider text-slate-600">
      <th class="px-4 py-3 w-12">#</th>
      <th class="px-4 py-3 tab-th" data-sort="name">Company</th>
      <th class="px-4 py-3 tab-th" data-sort="score">Score ▾</th>
      <th class="px-4 py-3">Signals</th>
      ${c.columns.map((col) => `<th class="px-4 py-3">${escapeHtml(col.label)}</th>`).join("")}
      <th class="px-4 py-3 text-right">Link</th>
    </tr>
  `;
  $$("#table-head th[data-sort]").forEach((th) => th.addEventListener("click", () => {
    const k = th.dataset.sort;
    if (state.sortBy === k) state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
    else { state.sortBy = k; state.sortDir = "desc"; }
    applyFilters();
  }));
  // body
  $("#table-body").innerHTML = rows.map((s, i) => {
    const name = c.name(s.company);
    const { color, initials } = avatarFor(name);
    const rank = st.scored.indexOf(s) + 1;
    const breakdownIcons = s.breakdown.slice(0, 8).map((b) => {
      const dot = ({ pass: "bg-emerald-500", partial: "bg-amber-400", fail: "bg-rose-400", hard_fail: "bg-rose-600", na: "bg-slate-300" })[b.status];
      return `<span class="w-1.5 h-1.5 rounded-full ${dot}" title="${escapeHtml(b.label)}: ${b.status}"></span>`;
    }).join("");
    // Split (fundamental red flag vs below-trend/illiquid filter) only applies
    // on the composite tab. On single-pillar tabs a hard-fail IS a genuine flag
    // for that pillar, so show it as a red flag with no "below trend" split.
    const redFlag = c.composite ? !!s.isRedFlag : s.hardFails.length > 0;
    const belowTrend = c.composite && s.hardFails.length > 0 && !s.isRedFlag;
    const slug = companyKey(s.company);
    const watched = state.watchlist.has(slug);
    return `
      <tr data-idx="${i}" class="row-clickable border-b border-slate-100 cursor-pointer transition-colors ${redFlag ? "bg-rose-50/40 hover:bg-rose-50" : "hover:bg-slate-50"}" ${redFlag ? `style="box-shadow: inset 3px 0 0 #f43f5e"` : belowTrend ? `style="box-shadow: inset 3px 0 0 #cbd5e1"` : ""}>
        <td class="px-4 py-3 text-sm text-slate-500 font-medium">
          <div class="flex items-center gap-1">
            <button data-watch="${escapeHtml(slug)}" class="watch-star text-base leading-none transition-colors ${watched ? "text-amber-400" : "text-slate-300 hover:text-amber-400"}" title="${watched ? "Remove from watchlist" : "Add to watchlist"}">${watched ? "★" : "☆"}</button>
            <span>${rank}</span>
          </div>
        </td>
        <td class="px-4 py-3">
          <div class="flex items-center gap-3">
            <div class="w-9 h-9 rounded-lg bg-gradient-to-br ${color} flex items-center justify-center text-white text-xs font-bold shadow-sm flex-shrink-0">${initials}</div>
            <div class="min-w-0">
              <div class="font-semibold text-slate-900 truncate">${escapeHtml(name)}</div>
              <div class="text-xs text-slate-500 truncate">${escapeHtml(c.marketCap(s.company))}</div>
            </div>
          </div>
        </td>
        <td class="px-4 py-3">
          <div class="flex items-center gap-2">
            <span class="inline-flex items-center justify-center min-w-[78px] px-2.5 py-1 rounded-lg text-sm font-bold tabular-nums ${scoreBadgeClass(s.scorePct)}">${c.composite ? Number(s.totalPoints).toFixed(1) : s.totalPoints}/${s.totalMax}</span>
            ${redFlag ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wide bg-rose-100 text-rose-700 ring-1 ring-rose-200" title="${escapeHtml((s.fundamentalFlags || s.hardFails || []).join(", "))}">⚠ Red Flag</span>` : belowTrend ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wide bg-slate-100 text-slate-500 ring-1 ring-slate-200" title="${escapeHtml((s.filterFlags || s.hardFails || []).join(", "))}">Below trend</span>` : ""}
            ${s.tickerError ? `<span class="text-[10px] text-slate-400 italic" title="${escapeHtml(s.tickerError)}">no data</span>` : ""}
          </div>
        </td>
        <td class="px-4 py-3"><div class="flex items-center gap-1">${breakdownIcons}</div></td>
        ${c.columns.map((col) => `<td class="px-4 py-3 text-sm text-slate-700">${col.html ? col.get(s.company) : escapeHtml(col.get(s.company))}</td>`).join("")}
        <td class="px-4 py-3 text-right">
          <a href="${escapeHtml(c.screenerUrl(s.company) || "")}" target="_blank" rel="noopener" class="text-indigo-600 hover:text-indigo-800 text-sm font-medium" onclick="event.stopPropagation()">↗</a>
        </td>
      </tr>
    `;
  }).join("") || `<tr><td colspan="${4 + c.columns.length + 2}" class="px-4 py-12 text-center text-slate-400">No companies match your filters.</td></tr>`;
  $$("#table-body .row-clickable").forEach((el) => el.addEventListener("click", () => openDrillDown(st.filtered[Number(el.dataset.idx)])));
  // Watchlist star handlers — stop propagation so the row click doesn't fire too
  $$("#table-body .watch-star").forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const k = btn.dataset.watch;
    if (!k) return;
    if (state.watchlist.has(k)) state.watchlist.delete(k);
    else state.watchlist.add(k);
    saveWatchlist(state.watchlist);
    updateWatchCount();
    // If "Watchlist only" is on, an unstar should drop the row from view
    if (state.watchOnly) applyFilters();
    else renderTable();
  }));
}

// Update the toolbar "Watchlist" pill count + active style.
function updateWatchCount() {
  const el = $("#watch-count");
  if (el) el.textContent = String(state.watchlist.size);
  const btn = $("#watch-toggle");
  if (btn) {
    const on = state.watchOnly;
    btn.classList.toggle("bg-amber-100", on);
    btn.classList.toggle("border-amber-300", on);
    btn.classList.toggle("text-amber-800", on);
    const icon = $("#watch-star-icon");
    if (icon) icon.textContent = on ? "★" : "☆";
  }
}

// ---------------- drill-down ----------------
function openDrillDown(s) {
  if (!s) return;
  const c = cfg();
  if (c.composite) return openCompositeDrill(s);
  const name = c.name(s.company);
  const { color, initials } = avatarFor(name);
  const co = s.company;
  const grouped = c.rules.reduce((acc, r) => { (acc[r.category] ||= []).push(r); return acc; }, {});
  const byKey = Object.fromEntries(s.breakdown.map((b) => [b.key, b]));

  const breakdownHtml = Object.entries(grouped).map(([cat, rules]) => `
    <div class="mb-5">
      <div class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">${escapeHtml(cat)}</div>
      <div class="space-y-2">
        ${rules.map((r) => {
          const b = byKey[r.key] || { status: "na", points: 0, max: r.fn ? 0 : 0, value: null, note: "—" };
          return `
            <div class="bg-white rounded-xl ring-1 ring-slate-100 p-3 hover:ring-slate-200 transition-shadow">
              <div class="flex items-start justify-between gap-2 mb-1">
                <div class="flex-1 min-w-0">
                  <div class="font-semibold text-slate-900 text-sm">${escapeHtml(r.label)}</div>
                  <div class="text-xs text-slate-500">Criteria: <span class="font-medium">${escapeHtml(r.criteria)}</span></div>
                </div>
                <div class="flex items-center gap-2 flex-shrink-0">
                  ${statusPill(b.status)}
                  <span class="text-sm font-bold text-slate-700">${b.points}/${b.max}</span>
                </div>
              </div>
              <div class="text-sm text-slate-700 mt-2">${b.value == null ? "—" : escapeHtml(b.value)}</div>
              <div class="text-xs text-slate-500 mt-1 italic">${escapeHtml(b.note)}</div>
              ${renderRuleMetaButtons(r.key, co)}
            </div>
          `;
        }).join("")}
      </div>
    </div>
  `).join("");

  // Only render the "Pending Data Source" block when there are actually
  // deferred rules — empty section was just visual noise on tabs whose
  // rules are all wired (Fundamentals, Technicals, etc. now have 0 deferred).
  const deferredHtml = (c.deferred && c.deferred.length) ? `
    <div class="mb-5">
      <div class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Pending Data Source</div>
      <div class="space-y-2">
        ${c.deferred.map((d) => `
          <div class="bg-amber-50 rounded-xl ring-1 ring-amber-100 p-3">
            <div class="flex items-start justify-between gap-2 mb-1">
              <div class="flex-1">
                <div class="font-semibold text-slate-900 text-sm">${escapeHtml(d.label)}</div>
                <div class="text-xs text-slate-500">Category: ${escapeHtml(d.category)} · Max ${d.max} pts</div>
              </div>
              ${statusPill("na")}
            </div>
            <div class="text-xs text-slate-600 mt-2">${escapeHtml(d.reason)}</div>
          </div>
        `).join("")}
      </div>
    </div>
  ` : "";

  const tier = s.hardFails.length ? "hardfail" : scoreTier(s.scorePct);
  const sector = c.sector(co), industry = c.industry(co);
  const headerStats = c.drillHeaderStats(co);

  $("#drill-content").innerHTML = `
    <div class="sticky top-0 bg-white/95 backdrop-blur-sm border-b border-slate-100 p-5 z-10">
      <button id="drill-close" class="absolute top-4 right-4 text-slate-400 hover:text-slate-700 text-2xl leading-none">×</button>
      <div class="flex items-center gap-4 pr-8">
        <div class="w-14 h-14 rounded-xl bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-lg shadow-md">${initials}</div>
        <div class="flex-1 min-w-0">
          <div class="font-bold text-xl text-slate-900 truncate">${escapeHtml(name)}</div>
          ${(sector || industry) ? `<div class="text-xs text-slate-500 truncate mt-0.5">${escapeHtml(sector || "")}${sector && industry ? " · " : ""}${escapeHtml(industry || "")}</div>` : ""}
          ${c.screenerUrl(co) ? `<a href="${escapeHtml(c.screenerUrl(co))}" target="_blank" rel="noopener" class="text-xs text-indigo-600 hover:text-indigo-800">View on Screener.in ↗</a>` : ""}
        </div>
      </div>
      <div class="grid grid-cols-3 gap-3 mt-4">
        <div class="bg-slate-50 rounded-lg p-3">
          <div class="text-xs text-slate-500 font-medium">${escapeHtml(c.label)} Score</div>
          <div class="text-2xl font-bold ${tierColor(tier)}">${s.totalPoints}<span class="text-sm text-slate-400">/${s.totalMax}</span></div>
          <div class="text-xs text-slate-500">${tierLabel(tier)}</div>
        </div>
        ${headerStats.map((hs) => `
          <div class="bg-slate-50 rounded-lg p-3">
            <div class="text-xs text-slate-500 font-medium">${escapeHtml(hs.label)}</div>
            ${hs.metrics ? `
              <div class="grid grid-cols-${hs.metrics.length} gap-1 mt-1">
                ${hs.metrics.map((m) => `
                  <div class="min-w-0">
                    <div class="text-[9px] text-slate-400 uppercase tracking-wider">${escapeHtml(m.name)}</div>
                    <div class="text-sm font-bold text-slate-900 leading-tight whitespace-nowrap overflow-hidden text-ellipsis" title="${escapeHtml(String(m.value))}">${escapeHtml(String(m.value))}</div>
                  </div>
                `).join("")}
              </div>
            ` : `<div class="text-base font-bold text-slate-900 truncate">${escapeHtml(hs.main || "")}</div>`}
            <div class="text-xs text-slate-500 truncate">${escapeHtml(hs.sub || "")}</div>
          </div>
        `).join("")}
      </div>
      ${s.tickerError ? `
        <div class="mt-3 p-3 bg-slate-100 rounded-lg ring-1 ring-slate-200">
          <div class="text-xs text-slate-600"><span class="font-semibold">Data note:</span> ${escapeHtml(s.tickerError)}. ${c.label} scoring not available for this company.</div>
        </div>` : ""}
      ${s.hardFails.length ? `
        <div class="mt-3 p-3 bg-rose-50 rounded-lg ring-1 ring-rose-100">
          <div class="flex items-start gap-2 mb-2">
            <div class="text-rose-500 text-lg leading-none">⚠</div>
            <div class="flex-1">
              <div class="font-semibold text-rose-800 text-sm">Red flag${s.hardFails.length>1?"s":""} (per client framework)</div>
              <div class="text-[11px] text-rose-700/80">All data is present — these signals are deliberately surfaced as cautionary.</div>
            </div>
          </div>
          <div class="space-y-1.5 mt-2">
            ${s.breakdown.filter(b=>b.status==="hard_fail").map(b=>`
              <div class="text-xs">
                <span class="font-bold text-rose-800">${escapeHtml(b.label)}:</span>
                <span class="text-rose-700">${escapeHtml(b.value || "—")}</span>
                <div class="text-rose-700/80 mt-0.5">${escapeHtml(b.note)}</div>
              </div>
            `).join("")}
          </div>
        </div>` : ""}
    </div>
    <div class="p-5">
      ${breakdownHtml}
      ${deferredHtml}
    </div>
  `;
  $("#drill-panel").classList.remove("translate-x-full");
  $("#drill-overlay").classList.remove("hidden");
  $("#drill-close").addEventListener("click", closeDrillDown);
}
// Render a circular progress gauge as inline SVG. Returns an HTML
// string with the score number centred inside the ring.
function renderScoreGauge(score, theme, max = 100, size = 144) {
  const r = (size - 16) / 2;
  const cx = size / 2; const cy = size / 2;
  const circ = 2 * Math.PI * r;
  const pct = score == null ? 0 : Math.max(0, Math.min(100, (score / max) * 100));
  const dash = (pct / 100) * circ;
  // Gradient stops keyed off the rating theme. Tailwind gradient
  // classes don't apply to SVG strokes, so we hardcode the hexes per
  // rating tier here.
  const gradMap = {
    "from-emerald-500 to-teal-500":   ["#10b981", "#14b8a6"],
    "from-blue-500 to-indigo-500":    ["#3b82f6", "#6366f1"],
    "from-amber-500 to-orange-500":   ["#f59e0b", "#f97316"],
    "from-rose-500 to-pink-500":      ["#f43f5e", "#ec4899"],
    "from-slate-500 to-slate-600":    ["#64748b", "#475569"],
    "from-slate-400 to-slate-500":    ["#94a3b8", "#64748b"],
  };
  const key = `${theme.from} ${theme.to}`;
  const [c1, c2] = gradMap[key] || ["#94a3b8", "#64748b"];
  const id = `g${Math.random().toString(36).slice(2,8)}`;
  return `
    <div class="relative" style="width:${size}px;height:${size}px">
      <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" class="-rotate-90">
        <defs>
          <linearGradient id="${id}" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="${c1}"/>
            <stop offset="100%" stop-color="${c2}"/>
          </linearGradient>
        </defs>
        <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#e2e8f0" stroke-width="10"/>
        <circle cx="${cx}" cy="${cy}" r="${r}" fill="none"
                stroke="url(#${id})" stroke-width="10" stroke-linecap="round"
                stroke-dasharray="${dash} ${circ}"
                data-gauge-arc="${dash} ${circ}" />
      </svg>
      <div class="absolute inset-0 flex flex-col items-center justify-center">
        <div class="text-4xl font-bold text-slate-900 leading-none">${score == null ? "—" : `<span class="count-up" data-target="${score}" data-decimals="1">0.0</span>`}</div>
        <div class="text-[10px] text-slate-500 uppercase tracking-wider mt-1">/ ${max}</div>
      </div>
    </div>
  `;
}

// Loads insider-trades.json (the daily NSE PIT scrape output) and merges
// in public/data/insider-recent-supplement.json (the routine's recent-
// disclosures output) into a single by-ticker map. The supplement covers
// the ~30-day window where NSE PIT's API has a blackout; the daily scrape
// owns everything older. Per-company aggregates are summed so the rule
// sees the FULL picture without us having to merge in scoring.js.
async function loadInsiderMerged() {
  const out = {};
  try {
    const base = await fetch("data/insider-trades.json").then((r) => r.json());
    Object.assign(out, base?.companies || {});
  } catch {}
  let supp = null;
  try { supp = await fetch("data/insider-recent-supplement.json").then((r) => r.json()); }
  catch { return out; }
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

// Lazy composite resolver — used by drill-down (any tab) to get the
// full 5-pillar shape for one company. If composite has already been
// computed (via SPIP tab or a prior drill-down), returns the cached
// result. Otherwise lazy-fetches technicals.json + macro.json and runs
// composite.scoreCompositeOne. Returns null if compute fails.
// Lazy-load the fundamentals universe and index by ticker so the cohort
// drill can resolve a ticker → company object without forcing the SPIP
// basket tab to be visited first.
//
// Note: prefers the enriched rows from state.cache.composite when SPIP
// basket has been loaded — those carry supplemental data (auditor opinion,
// governance flag, revenue-mix) that hard-fail rules depend on. Falls
// back to raw screener-companies.json otherwise.
async function fetchCompanyByTicker(ticker) {
  if (!ticker) return null;
  const tu = String(ticker).toUpperCase();
  if (state.cache.composite?.rows) {
    const enriched = state.cache.composite.rows.find((r) => {
      const m = String(r?.["Screener URL"] || "").match(/\/company\/([^/]+)/);
      return m && m[1].toUpperCase() === tu;
    });
    if (enriched) return enriched;
  }
  if (!state.lazyFundByTicker) {
    try {
      const data = await fetch("data/screener-companies.json").then((r) => r.json());
      const rows = Array.isArray(data) ? data : (data.companies || []);
      state.lazyFundByTicker = new Map();
      for (const r of rows) {
        const m = String(r["Screener URL"] || "").match(/\/company\/([^/]+)/);
        if (m) state.lazyFundByTicker.set(m[1].toUpperCase(), r);
      }
    } catch { state.lazyFundByTicker = new Map(); }
  }
  return state.lazyFundByTicker.get(tu) || null;
}

// Populate the cohort-drill hard-fail panel with the actual rule(s) that
// tripped. Snapshots store the hardFailed boolean but not the rule names —
// we fetch them via live composite scoring (warming the full composite
// cache first so supplemental rules like auditor / governance / revenue-
// mix fire correctly). Includes a stale-write guard via the panel's
// data-ticker attribute and a distinct load-error path when scoring fails.
async function populateHardFailRules(ticker) {
  const writeIfCurrent = (html, asText) => {
    const panel = $("#hist-hardfail-panel");
    const el = $("#hist-hardfail-rules");
    if (!el || !panel || panel.dataset.ticker !== ticker) return;
    if (asText) el.textContent = html;
    else el.innerHTML = html;
  };

  // Warm composite cache first so the company gets the same enrichments
  // (auditor opinion, governance flags, revenue-mix) the AI basket uses.
  if (!state.cache.composite) {
    try { await loadTab("composite"); } catch {}
  }

  const co = await fetchCompanyByTicker(ticker);
  if (!co) {
    writeIfCurrent("Couldn't resolve ticker — open AI Basket for the full breakdown.", true);
    return;
  }
  const result = await ensureCompositeFor(co);
  if (result == null) {
    writeIfCurrent("Couldn't load live rule list — open AI Basket for the full breakdown.", true);
    return;
  }
  const fails = Array.isArray(result.hardFails) ? result.hardFails : [];
  if (!fails.length) {
    writeIfCurrent("Rating is FILTERED in the snapshot but no live hard-fail rule fired — likely a snapshot-vs-live mismatch (rule passed since the last snapshot).", true);
    return;
  }
  writeIfCurrent(
    `<span class="font-semibold text-rose-800">Failing rule${fails.length === 1 ? "" : "s"}:</span> ` +
      fails.map((h) => `<span class="inline-flex px-1.5 py-0 rounded bg-white text-rose-700 ring-1 ring-rose-200 font-bold ml-1">${escapeHtml(h)}</span>`).join(""),
    false,
  );
}

async function ensureCompositeFor(co) {
  if (!co) return null;
  const slug = companyKey(co);
  if (!slug) return null;
  if (state.compositeBySlug.has(slug)) return state.compositeBySlug.get(slug);
  if (co._composite) {
    state.compositeBySlug.set(slug, co._composite);
    return co._composite;
  }
  // Need technicals + macro to compute pillars. Load each once globally.
  if (!state.lazyTechBySlug) {
    try {
      const tj = await fetch("data/technicals.json").then((r) => r.json());
      const trows = (tj.companies || tj) || [];
      state.lazyTechBySlug = new Map();
      for (const t of trows) {
        const k = String(t.ticker || "").toUpperCase();
        if (k) state.lazyTechBySlug.set(k, t);
      }
    } catch { state.lazyTechBySlug = new Map(); }
  }
  if (!state.lazyMacroCtx) {
    try { state.lazyMacroCtx = await fetch("data/macro.json").then((r) => r.json()); }
    catch { state.lazyMacroCtx = {}; }
  }
  try {
    const techCo = state.lazyTechBySlug.get(slug) || null;
    const result = composite.scoreCompositeOne(co, techCo, state.lazyMacroCtx);
    state.compositeBySlug.set(slug, result);
    return result;
  } catch { return null; }
}

// Pillar card — premium card per pillar showing raw score, %, weight, and
// weighted contribution. Each line gets its own row for clean alignment.
function renderPillarCard(label, p, weight) {
  if (!p || p.raw == null) {
    return `
      <div class="rounded-xl ring-1 ring-slate-200 bg-slate-50/40 p-4">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">${escapeHtml(label)}</div>
        <div class="text-slate-300 text-3xl font-bold leading-none">—</div>
        <div class="text-[10px] text-slate-400 mt-1.5">no data</div>
        <div class="mt-3 h-2 rounded-full bg-slate-100"></div>
        <div class="mt-3 pt-3 border-t border-slate-200/80 flex items-baseline justify-between">
          <span class="text-[10px] text-slate-400 uppercase tracking-wider">${weight}% weight</span>
          <span class="font-bold text-slate-300 text-sm">+0.0</span>
        </div>
      </div>
    `;
  }
  const pct = p.pct ?? 0;
  const tier = pct >= 75 ? "emerald" : pct >= 60 ? "blue" : pct >= 45 ? "amber" : "rose";
  const barFill = ({ emerald: "bg-emerald-500", blue: "bg-blue-500", amber: "bg-amber-500", rose: "bg-rose-500" })[tier];
  const accent  = ({ emerald: "text-emerald-700", blue: "text-blue-700", amber: "text-amber-700", rose: "text-rose-700" })[tier];
  const bg      = ({ emerald: "bg-emerald-50/60", blue: "bg-blue-50/60", amber: "bg-amber-50/60", rose: "bg-rose-50/60" })[tier];
  return `
    <div class="rounded-xl ring-1 ring-slate-200 ${bg} p-4 hover:shadow-md hover:-translate-y-0.5 transition-all">
      <!-- Pillar name -->
      <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 truncate">${escapeHtml(label)}</div>
      <!-- Raw score on its own row, large -->
      <div class="mt-2 flex items-baseline gap-1.5">
        <span class="text-3xl font-bold text-slate-900 leading-none">${p.raw}</span>
        <span class="text-sm text-slate-400">/ ${p.max}</span>
      </div>
      <!-- Percentage on its own row -->
      <div class="mt-1.5 text-xs font-semibold ${accent}">${pct}%</div>
      <!-- Progress bar -->
      <div class="mt-2.5 h-2 rounded-full bg-white/70 overflow-hidden ring-1 ring-slate-200/80">
        <div class="${barFill} h-2 rounded-full" data-bar-width="${Math.min(100, pct)}%" style="width: ${Math.min(100, pct)}%"></div>
      </div>
      <!-- Weight + contribution footer -->
      <div class="mt-3 pt-3 border-t border-slate-200/80 flex items-baseline justify-between">
        <span class="text-[10px] text-slate-500 uppercase tracking-wider">${weight}% weight</span>
        <span class="font-bold text-slate-900 text-sm">+${(p.weighted ?? 0).toFixed(1)}</span>
      </div>
    </div>
  `;
}

// ---------------- AI Basket — radar / thesis / animation / magazine ----------------

// 5-axis pillar radar (Fund / Tech / Macro / Sent / Liq) rendered as
// inline SVG. Concentric rings at 20/40/60/80/100, axes at 5 vertices,
// data polygon coloured to match the rating tier. Each axis is labelled
// at the perimeter with the pillar's raw/max + percentage.
//
// Layout: the `size` parameter is the chart RADIUS extent — labels
// extend into a `LABEL_PAD` zone outside that, so the actual SVG is
// (size + 2*LABEL_PAD) wide/tall. This prevents the right-edge labels
// (e.g. "17.5/24") from getting clipped when long.
// Two-pillar stand-in for the radar. Same data, same colours, readable at
// n=2 where a polygon is not.
function renderPillarBars(s, theme, keys) {
  const NAME = { fundamentals: "Fundamentals", technicals: "Technicals", macro: "Macro", sentiment: "Sentiment", liquidity: "Liquidity" };
  const rows = keys.map((k) => {
    const p = s.pillars?.[k] || {};
    const pct = p.pct ?? 0;
    const tier = pct >= 75 ? "bg-emerald-500" : pct >= 60 ? "bg-blue-500" : pct >= 45 ? "bg-amber-500" : "bg-rose-500";
    return `
      <div class="mb-4">
        <div class="flex items-baseline justify-between mb-1.5">
          <span class="text-xs font-bold uppercase tracking-wider text-slate-600">${NAME[k] || k}</span>
          <span class="text-sm font-bold tabular-nums text-slate-900">${pct}%
            <span class="text-[11px] font-normal text-slate-400">${p.raw ?? "—"}/${p.max ?? "—"}</span>
          </span>
        </div>
        <div class="h-3 rounded-full bg-slate-100 ring-1 ring-slate-200 overflow-hidden">
          <div class="h-full ${tier} rounded-full transition-all" style="width:${Math.max(0, Math.min(100, pct))}%"></div>
        </div>
      </div>`;
  }).join("");
  return `<div class="w-full px-6 py-4">${rows}</div>`;
}

function renderPillarRadar(s, theme, size = 280) {
  // A radar needs >= 3 axes to enclose any area. With two live pillars the
  // polygon degenerates to a line and reads as a rendering bug, so fall back
  // to horizontal bars, which compare two values better anyway.
  const liveKeys = livePillars();
  if (liveKeys.length < 3) return renderPillarBars(s, theme, liveKeys);
  const LABEL_PAD = 64;             // room around chart for labels
  const svgSize = size + LABEL_PAD * 2;
  const cx = svgSize / 2, cy = svgSize / 2;
  const r = size / 2 - 6;           // chart radius proper
  // 5 axes — start at top (-90°) and step every 72°
  // Axes follow the live weights. A radar needs at least 3 axes to be a
  // shape at all -- with two pillars it collapses to a line -- so the caller
  // falls back to bars below that.
  const AXIS_LABEL = { fundamentals: "FUND", technicals: "TECH", macro: "MACRO", sentiment: "SENT", liquidity: "LIQ" };
  const axes = livePillars().map((k) => ({ key: k, label: AXIS_LABEL[k] }));
  const ptAt = (axisIdx, pct) => {
    const angle = -Math.PI / 2 + (axisIdx * 2 * Math.PI / axes.length);
    return { x: cx + Math.cos(angle) * r * (pct / 100), y: cy + Math.sin(angle) * r * (pct / 100) };
  };
  // Concentric guide rings (20/40/60/80/100 %)
  const rings = [20, 40, 60, 80, 100].map((p) => {
    const pts = axes.map((_, i) => ptAt(i, p)).map((q) => `${q.x.toFixed(1)},${q.y.toFixed(1)}`).join(" ");
    return `<polygon points="${pts}" fill="none" stroke="#e2e8f0" stroke-width="1"/>`;
  }).join("");
  // Axis lines (centre → vertex at 100%)
  const axisLines = axes.map((_, i) => {
    const p = ptAt(i, 100);
    return `<line x1="${cx}" y1="${cy}" x2="${p.x.toFixed(1)}" y2="${p.y.toFixed(1)}" stroke="#e2e8f0" stroke-width="1"/>`;
  }).join("");
  // Data polygon
  const dataPts = axes.map((a, i) => {
    const pct = s.pillars?.[a.key]?.pct ?? 0;
    const p = ptAt(i, pct);
    return `${p.x.toFixed(1)},${p.y.toFixed(1)}`;
  }).join(" ");
  // Gradient stops keyed off the theme.
  const gradMap = {
    "from-emerald-500 to-teal-500":   ["#10b981", "#14b8a6"],
    "from-blue-500 to-indigo-500":    ["#3b82f6", "#6366f1"],
    "from-amber-500 to-orange-500":   ["#f59e0b", "#f97316"],
    "from-rose-500 to-pink-500":      ["#f43f5e", "#ec4899"],
    "from-slate-500 to-slate-600":    ["#64748b", "#475569"],
    "from-slate-400 to-slate-500":    ["#94a3b8", "#64748b"],
  };
  const key = `${theme.from} ${theme.to}`;
  const [c1, c2] = gradMap[key] || ["#94a3b8", "#64748b"];
  const gid = `rg${Math.random().toString(36).slice(2, 8)}`;
  // Axis labels — placed at a fixed pixel offset OUTSIDE the chart
  // radius so they never overlap the polygon. We use the LABEL_PAD
  // padding zone for them, so text never clips at the SVG edge.
  const labels = axes.map((a, i) => {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI / 5);
    const labelOffset = r + 26;     // distance from centre to the label baseline
    const lx = cx + Math.cos(angle) * labelOffset;
    const ly = cy + Math.sin(angle) * labelOffset;
    const p = s.pillars?.[a.key];
    const pct = p?.pct ?? 0;
    const raw = p?.raw ?? "—";
    const max = p?.max ?? "—";
    // Anchor: middle for top/bottom (i=0,3), start for right (i=1,2),
    // end for left (i=3,4). Determined by sign of cos(angle).
    const cos = Math.cos(angle);
    const anchor = Math.abs(cos) < 0.1 ? "middle" : (cos > 0 ? "start" : "end");
    // Vertical anchor: shift dy upward for top vertex, downward for
    // bottom vertices, so the 3-line label block doesn't overlap the
    // chart polygon visually.
    const sin = Math.sin(angle);
    const blockHeight = 36;          // 3 lines × ~12 px
    const dy = sin < -0.5 ? -blockHeight + 8     // top-ish vertex: pull whole block up
             : sin > 0.5  ? 6                    // bottom-ish vertex: leave it just below
             : -8;                                // sides: centre-ish
    return `
      <g text-anchor="${anchor}" transform="translate(0, ${dy})">
        <text x="${lx.toFixed(1)}" y="${ly.toFixed(1)}" class="font-bold" font-size="10" fill="#64748b" letter-spacing="0.08em">${a.label}</text>
        <text x="${lx.toFixed(1)}" y="${(ly + 13).toFixed(1)}" font-size="12" font-weight="700" fill="#0f172a">${raw}/${max}</text>
        <text x="${lx.toFixed(1)}" y="${(ly + 25).toFixed(1)}" font-size="10" fill="#94a3b8">${pct}%</text>
      </g>
    `;
  }).join("");
  // Vertex dots on the data polygon
  const vertices = axes.map((a, i) => {
    const pct = s.pillars?.[a.key]?.pct ?? 0;
    const p = ptAt(i, pct);
    return `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4" fill="#fff" stroke="${c1}" stroke-width="2"/>`;
  }).join("");

  return `
    <div class="relative" style="width:${svgSize}px;height:${svgSize}px" data-radar="1">
      <svg width="${svgSize}" height="${svgSize}" viewBox="0 0 ${svgSize} ${svgSize}" style="overflow: visible;">
        <defs>
          <linearGradient id="${gid}" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="${c1}" stop-opacity="0.65"/>
            <stop offset="100%" stop-color="${c2}" stop-opacity="0.35"/>
          </linearGradient>
        </defs>
        ${rings}
        ${axisLines}
        <g class="radar-data" style="transform-origin: ${cx}px ${cy}px; transform: scale(0); transition: transform 700ms cubic-bezier(0.34, 1.56, 0.64, 1);">
          <polygon points="${dataPts}" fill="url(#${gid})" stroke="${c1}" stroke-width="2.5"/>
          ${vertices}
        </g>
        ${labels}
      </svg>
    </div>
  `;
}

// Count-up animator. Animates `.count-up` elements from 0 to their
// data-target value over ~900ms with ease-out. Call after the modal
// content is mounted; the same call also fires the radar scale-in
// (CSS transition) and the gauge dash sweep.
function animateScoreEntrance(root = document) {
  // Count-up numbers
  root.querySelectorAll(".count-up").forEach((el) => {
    const target = parseFloat(el.dataset.target);
    const decimals = parseInt(el.dataset.decimals || "0", 10);
    if (!Number.isFinite(target)) return;
    const start = performance.now();
    const dur = 900;
    function tick(now) {
      const t = Math.min(1, (now - start) / dur);
      // ease-out cubic
      const eased = 1 - Math.pow(1 - t, 3);
      el.textContent = (target * eased).toFixed(decimals);
      if (t < 1) requestAnimationFrame(tick);
      else el.textContent = target.toFixed(decimals);
    }
    requestAnimationFrame(tick);
  });
  // Gauge ring sweep: start at 0 dasharray, transition to final
  root.querySelectorAll("[data-gauge-arc]").forEach((arc) => {
    const finalDash = arc.dataset.gaugeArc;
    arc.setAttribute("stroke-dasharray", "0 100000");
    requestAnimationFrame(() => {
      arc.style.transition = "stroke-dasharray 1100ms cubic-bezier(0.34, 1.56, 0.64, 1)";
      arc.setAttribute("stroke-dasharray", finalDash);
    });
  });
  // Radar polygon scale-in (transform set inline; just kick the property)
  root.querySelectorAll(".radar-data").forEach((el) => {
    requestAnimationFrame(() => { el.style.transform = "scale(1)"; });
  });
  // Pillar bar width sweep
  root.querySelectorAll("[data-bar-width]").forEach((bar) => {
    const finalWidth = bar.dataset.barWidth;
    bar.style.width = "0%";
    requestAnimationFrame(() => {
      bar.style.transition = "width 900ms cubic-bezier(0.34, 1.56, 0.64, 1)";
      bar.style.width = finalWidth;
    });
  });
}

// Synthesised "thesis" — short narrative built from the actual scoring
// data. Picks the strongest pillar contribution, the weakest pillar,
// and the recommended action; assembles them into 2–3 sentences. No
// LLM call, no fabricated numbers; everything reads from `s`.
function synthesizeThesis(s) {
  const p = s.pillars || {};
  const valid = Object.entries(p).filter(([_, v]) => v.raw != null && v.pct != null);
  if (!valid.length) return "Scoring data incomplete — thesis unavailable.";
  const labelMap = { fundamentals: "Fundamentals", technicals: "Technicals", macro: "Macro tailwind", sentiment: "Sentiment", liquidity: "Liquidity" };
  // Strongest = highest percentage; weakest = lowest percentage among those that scored.
  const sorted = [...valid].sort((a, b) => b[1].pct - a[1].pct);
  const top = sorted[0];
  const bot = sorted[sorted.length - 1];
  const composite = s.composite != null ? s.composite.toFixed(1) : "—";
  const rating = s.rating || "—";

  const sentences = [];
  sentences.push(`Composite scores <strong>${composite}/100</strong> — ${rating} territory per the scoring framework.`);
  if (top && top[1].pct >= 75) {
    sentences.push(`Anchored by <strong>${labelMap[top[0]] || top[0]}</strong> at <strong>${top[1].raw}/${top[1].max}</strong> (${top[1].pct}%), contributing <strong>+${(top[1].weighted ?? 0).toFixed(1)}</strong> of the composite.`);
  } else if (top) {
    sentences.push(`Leading contributor: ${labelMap[top[0]] || top[0]} at ${top[1].raw}/${top[1].max} (${top[1].pct}%).`);
  }
  if (bot && bot[1].pct < 50 && bot[0] !== top[0]) {
    const isMarketWide = (bot[0] === "sentiment" || bot[0] === "macro");
    const qualifier = isMarketWide ? " (market-wide factor — not company-specific)" : "";
    sentences.push(`Watch: <strong>${labelMap[bot[0]] || bot[0]}</strong> drags at ${bot[1].pct}%${qualifier}.`);
  }
  if (s.hardFails && s.hardFails.length) {
    sentences.push(`<strong>Hard-fail triggered:</strong> ${s.hardFails.join(", ")} — stock excluded from basket regardless of composite.`);
  }
  return sentences.join(" ");
}

// Magazine ("Full Story") drill — fullscreen takeover with a hero,
// pillar radar, by-the-numbers sidebar, thesis pull-quote, and
// recommended action. Triggered from the standard drill via the
// "View Full Story" button (only when not hard-failed).
function openMagazineDrill(s) {
  const co = s.company || {};
  const name = co.Company || "—";
  const sector = co.Sector || "";
  const theme = composite.ratingTheme(s.rating);
  const decision = composite.decisionFor(s.rating);
  const thesis = synthesizeThesis(s);

  // "By the numbers" — pull whatever ratios are available from the
  // fundamentals row. Each cell falls back to "—" when missing. The
  // first two (Composite + Rating) are rendered in their own hero
  // block above the secondary metrics list.
  const heroMetrics = [
    { label: "Composite",  value: s.composite != null ? s.composite.toFixed(1) : "—", suffix: "/ 100", accent: theme.accent },
    { label: "Rating",     value: s.rating, suffix: "",       accent: theme.accent },
  ];
  const ratios = [
    { label: "Market Cap",      value: co["Market Cap"] || "—" },
    { label: "CMP",             value: co["Current Price"] ? `₹${co["Current Price"]}` : "—" },
    { label: "P/E",             value: co["Stock P/E"] || "—" },
    { label: "ROE",             value: co["ROE"] || "—" },
    { label: "ROCE",            value: co["ROCE"] || "—" },
    { label: "Debt / Equity",   value: co["Debt to equity"] || "—" },
    { label: "Rev 3Y CAGR",     value: co["Sales growth 3Years"] || "—" },
    { label: "PAT 3Y CAGR",     value: co["Profit Var 3Yrs"] || "—" },
  ];

  openModal(`
    <div class="relative overflow-hidden">
      <div class="absolute inset-0 bg-gradient-to-br ${theme.from} ${theme.to}"></div>
      <div class="absolute inset-0 bg-[radial-gradient(circle_at_top_right,rgba(255,255,255,0.20),transparent_55%),radial-gradient(circle_at_bottom_left,rgba(255,255,255,0.10),transparent_60%)]"></div>
      <button id="modal-close" class="absolute top-4 right-4 z-10 text-white/85 hover:text-white text-2xl leading-none w-10 h-10 rounded-full bg-white/15 hover:bg-white/25 backdrop-blur flex items-center justify-center transition-colors">×</button>
      <div class="relative px-8 py-7 ${theme.textOn}">
        <div class="flex items-center gap-2 text-[11px] font-semibold uppercase tracking-[0.25em] opacity-85">
          <span>Basket Brief</span>
          <span class="opacity-50">·</span>
          <span>${escapeHtml(s.rating)}</span>
          <span class="opacity-50">·</span>
          <span>${escapeHtml(sector || "—")}</span>
        </div>
        <h1 class="font-display font-extrabold text-4xl sm:text-5xl mt-3 leading-[1.05] tracking-tight">${escapeHtml(name)}</h1>
        <p class="text-[15px] opacity-95 mt-3 max-w-3xl leading-[1.65]">${thesis}</p>
      </div>
    </div>

    <div class="grid grid-cols-1 lg:grid-cols-3 gap-0 max-h-[calc(95vh-240px)] overflow-y-auto">
      <!-- LEFT 2/3 — radar + decision -->
      <div class="lg:col-span-2 p-7 border-r border-slate-100">
        <div class="flex items-baseline justify-between mb-2">
          <div class="text-[10px] font-bold uppercase tracking-[0.15em] text-slate-500">Pillar Composition</div>
          <div class="text-[11px] text-slate-500">Weighted ${weightsLabel().nums}</div>
        </div>
        <div class="flex items-center justify-center -my-2">
          ${renderPillarRadar(s, theme, 280)}
        </div>

        <!-- Decision band -->
        <div class="bg-gradient-to-br ${theme.soft} rounded-2xl ring-1 ${theme.ring} p-5">
          <div class="flex items-center justify-between mb-2">
            <div class="text-[10px] font-bold uppercase tracking-[0.15em] ${theme.accent}">Recommended Action</div>
            <span class="text-[10px] text-slate-500 whitespace-nowrap">Decision framework</span>
          </div>
          <div class="text-lg font-bold text-slate-900 mb-4 leading-snug">${escapeHtml(decision.action)}</div>
          <div class="grid grid-cols-3 gap-4">
            <div class="border-l-2 ${theme.ring.replace("ring-", "border-")} pl-3">
              <div class="text-[10px] text-slate-500 uppercase tracking-wider mb-1">Size</div>
              <div class="font-semibold text-slate-900 leading-snug text-xs">${escapeHtml(decision.size)}</div>
            </div>
            <div class="border-l-2 ${theme.ring.replace("ring-", "border-")} pl-3">
              <div class="text-[10px] text-slate-500 uppercase tracking-wider mb-1">Review</div>
              <div class="font-semibold text-slate-900 leading-snug text-xs">${escapeHtml(decision.review)}</div>
            </div>
            <div class="border-l-2 ${theme.ring.replace("ring-", "border-")} pl-3">
              <div class="text-[10px] text-slate-500 uppercase tracking-wider mb-1">Exit Trigger</div>
              <div class="font-semibold text-slate-900 leading-snug text-[11px]">${escapeHtml(decision.exit)}</div>
            </div>
          </div>
        </div>
      </div>

      <!-- RIGHT 1/3 — by the numbers -->
      <div class="p-7 bg-slate-50/40">
        <div class="text-[10px] font-bold uppercase tracking-[0.15em] text-slate-500 mb-4">By the Numbers</div>

        <!-- Hero metrics: Composite + Rating, larger format -->
        <div class="space-y-2.5 mb-5">
          ${heroMetrics.map((m) => `
            <div class="bg-white rounded-xl p-3 ring-1 ring-slate-200/70">
              <div class="text-[10px] text-slate-500 uppercase tracking-wider">${escapeHtml(m.label)}</div>
              <div class="flex items-baseline gap-1.5 mt-0.5">
                <span class="text-2xl font-bold ${m.accent} leading-none">${escapeHtml(m.value || "—")}</span>
                ${m.suffix ? `<span class="text-xs text-slate-400">${escapeHtml(m.suffix)}</span>` : ""}
              </div>
            </div>
          `).join("")}
        </div>

        <!-- Secondary metrics: vertical-accent rows -->
        <div class="space-y-0">
          ${ratios.map((r, i) => `
            <div class="flex items-center justify-between gap-3 py-2.5 ${i < ratios.length - 1 ? "border-b border-slate-200/60" : ""}">
              <div class="flex items-center gap-2.5">
                <div class="w-0.5 h-4 rounded-full bg-gradient-to-b ${theme.from} ${theme.to} opacity-60"></div>
                <div class="text-xs text-slate-600">${escapeHtml(r.label)}</div>
              </div>
              <div class="font-bold text-slate-900 text-sm text-right">${escapeHtml(r.value || "—")}</div>
            </div>
          `).join("")}
        </div>

        <div class="mt-6 pt-4 border-t border-slate-200">
          <button id="mag-back" class="text-xs font-semibold text-indigo-600 hover:text-indigo-800 transition-colors inline-flex items-center gap-1">
            <span class="text-base leading-none">←</span> Back to standard view
          </button>
        </div>
      </div>
    </div>
  `, { size: "magazine" });

  // Wire close / back / dismiss + run the entrance animations
  $("#modal-close")?.addEventListener("click", closeModal);
  $("#mag-back")?.addEventListener("click", () => { closeModal(); openCompositeDrill(s); });
  $("#modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "modal-overlay") closeModal();
  }, { once: true });
  // Fire animations on the radar (gauge isn't on this view)
  requestAnimationFrame(() => animateScoreEntrance($("#modal-content")));
}

// AI Basket drill — opens as a CENTRED MODAL (not the side panel).
// Hero gauge + decision card on top, 5-card pillar composition below,
// then drill-deeper shortcuts to the per-pillar tabs.
function openCompositeDrill(s) {
  const co = s.company || {};
  const name = co.Company || "—";
  const sector = co.Sector || ""; const industry = co["Broad Industry"] || "";
  const url = co["Screener URL"] || null;
  const theme = composite.ratingTheme(s.rating);
  const decision = composite.decisionFor(s.rating);
  const initials = avatarFor(name).initials;

  // --- Hero header
  const heroHeader = `
    <div class="relative overflow-hidden">
      <div class="absolute inset-0 bg-gradient-to-br ${theme.from} ${theme.to}"></div>
      <div class="absolute inset-0 bg-[radial-gradient(circle_at_top_right,rgba(255,255,255,0.20),transparent_55%)]"></div>
      <button id="modal-close" class="absolute top-3 right-3 z-10 text-white/85 hover:text-white text-2xl leading-none w-9 h-9 rounded-full bg-white/15 hover:bg-white/25 backdrop-blur flex items-center justify-center transition-colors">×</button>
      <div class="relative p-6 ${theme.textOn}">
        <div class="flex items-center gap-4 pr-12">
          <div class="w-16 h-16 rounded-2xl bg-white/20 backdrop-blur ring-1 ring-white/30 flex items-center justify-center text-white font-bold text-xl shadow-lg flex-shrink-0">${initials}</div>
          <div class="flex-1 min-w-0">
            <div class="font-bold text-2xl sm:text-3xl truncate">${escapeHtml(name)}</div>
            ${(sector || industry) ? `<div class="text-sm opacity-90 truncate mt-1">${escapeHtml(sector)}${sector && industry ? " · " : ""}${escapeHtml(industry)}</div>` : ""}
            <div class="flex flex-wrap items-center gap-x-3 gap-y-1 mt-2 text-xs">
              ${url ? `<a href="${escapeHtml(url)}" target="_blank" rel="noopener" class="opacity-95 hover:opacity-100 underline decoration-white/40 hover:decoration-white whitespace-nowrap">View on Screener.in ↗</a>` : ""}
              <span class="opacity-95 whitespace-nowrap">${escapeHtml(co["Market Cap"] || "—")} mkt cap</span>
              <span class="opacity-95 whitespace-nowrap">CMP ${escapeHtml(co["Current Price"] || "—")}</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;

  // --- Hero panel: gauge | decision card. Modal is wide, so we can give
  // each panel real breathing room rather than cramming them.
  const heroPanel = `
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
      <div class="bg-white rounded-2xl ring-1 ring-slate-200 p-6 flex items-center gap-6">
        ${renderScoreGauge(s.composite, theme, 100, 168)}
        <div class="flex-1 min-w-0">
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Composite Score</div>
          <div class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-gradient-to-r ${theme.from} ${theme.to} ${theme.textOn} text-sm font-bold shadow-sm">
            <span class="w-1.5 h-1.5 rounded-full bg-white"></span>
            ${escapeHtml(s.rating)}
          </div>
          <div class="text-sm text-slate-600 mt-3 leading-snug">${escapeHtml(decision.profile)}</div>
        </div>
      </div>
      <div class="bg-gradient-to-br ${theme.soft} rounded-2xl ring-1 ${theme.ring} p-6">
        <div class="flex items-center justify-between mb-2">
          <div class="text-[10px] font-bold uppercase tracking-wider ${theme.accent}">Recommended Action</div>
          <span class="text-[10px] text-slate-500 whitespace-nowrap">Decision framework</span>
        </div>
        <div class="text-lg font-bold text-slate-900 mb-4 leading-snug">${escapeHtml(decision.action)}</div>
        <div class="grid grid-cols-3 gap-3 text-xs">
          <div>
            <div class="text-[10px] text-slate-500 uppercase tracking-wider mb-1">Size</div>
            <div class="font-semibold text-slate-900 leading-snug">${escapeHtml(decision.size)}</div>
          </div>
          <div>
            <div class="text-[10px] text-slate-500 uppercase tracking-wider mb-1">Review</div>
            <div class="font-semibold text-slate-900 leading-snug">${escapeHtml(decision.review)}</div>
          </div>
          <div>
            <div class="text-[10px] text-slate-500 uppercase tracking-wider mb-1">Exit Trigger</div>
            <div class="font-semibold text-slate-900 leading-snug text-[11px]">${escapeHtml(decision.exit)}</div>
          </div>
        </div>
      </div>
    </div>
  `;

  const hardFailPanel = s.hardFails.length ? `
    <div class="mb-6 p-5 bg-gradient-to-br from-rose-50 to-pink-50 rounded-2xl ring-1 ring-rose-200">
      <div class="flex items-start gap-3 mb-3">
        <div class="w-10 h-10 rounded-xl bg-rose-500 text-white flex items-center justify-center text-xl flex-shrink-0">⚠</div>
        <div class="flex-1">
          <div class="font-bold text-rose-900 text-base">Excluded from AI Basket</div>
          <div class="text-sm text-rose-700/90 mt-0.5">Hard fail per client framework · stock exits pipeline regardless of composite score</div>
        </div>
      </div>
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-${s.hardFails.length > 2 ? 3 : 2} gap-2 mt-3">
        ${s.hardFails.map((h) => `
          <div class="bg-white rounded-lg p-3 ring-1 ring-rose-100">
            <div class="text-[10px] font-bold uppercase tracking-wider text-rose-500">Red Flag</div>
            <div class="text-sm font-bold text-rose-900 mt-0.5">${escapeHtml(h)}</div>
          </div>
        `).join("")}
      </div>
    </div>
  ` : "";

  const unratedPanel = (!s.hardFails.length && s.unrated) ? `
    <div class="mb-6 p-5 bg-gradient-to-br from-amber-50 to-orange-50 rounded-2xl ring-1 ring-amber-200">
      <div class="flex items-start gap-3">
        <div class="w-10 h-10 rounded-xl bg-amber-500 text-white flex items-center justify-center text-xl flex-shrink-0">ℹ</div>
        <div class="flex-1">
          <div class="font-bold text-amber-900 text-base">Unrated — Technicals data missing</div>
          <div class="text-sm text-amber-800/90 mt-0.5">Yahoo doesn't carry OHLCV for this ticker (REITs / InvITs / very-recent listings). Composite will populate when a vendor with NSE coverage is wired.</div>
        </div>
      </div>
    </div>
  ` : "";

  // --- Pillar composition. Cards and their weights both come from the live
  // mix: this used to hardcode five cards at 40/35/15/5/5, so it rendered
  // three permanent "no data / 15% weight" tiles for pillars that had been
  // removed, and misreported the weights of the two that remained.
  const PILLAR_CARD_NAME = { fundamentals: "Fundamentals", technicals: "Technicals", macro: "Macro", sentiment: "Sentiment", liquidity: "Liquidity" };
  const cardKeys = livePillars();
  const wNow = state.pillarWeights || composite.PILLAR_WEIGHTS;
  const pillarCards = `
    <div class="mb-6">
      <div class="flex items-baseline justify-between mb-3">
        <div class="text-xs font-bold uppercase tracking-wider text-slate-500">Pillar Composition</div>      </div>
      <div class="grid gap-3" style="grid-template-columns: repeat(${Math.min(cardKeys.length, 5)}, minmax(0, 1fr));">
        ${cardKeys.map((k) => renderPillarCard(PILLAR_CARD_NAME[k] || k, s.pillars?.[k], wNow[k] ?? 0)).join("")}
      </div>
      <div class="mt-3 flex items-center justify-between bg-gradient-to-r from-slate-50 to-slate-100 rounded-xl px-4 py-3 ring-1 ring-slate-200">
        <div class="flex items-center gap-2">
          <span class="text-[10px] font-bold uppercase tracking-wider text-slate-500">Composite Score</span>
          <span class="text-[10px] text-slate-400">(weighted sum)</span>
        </div>
        <div class="flex items-baseline gap-1">
          <span class="text-2xl font-bold ${theme.accent}">${s.composite != null ? s.composite.toFixed(1) : "—"}</span>
          <span class="text-sm text-slate-400">/ 100</span>
        </div>
      </div>
    </div>
  `;

  // Only pillars that still exist as tabs, with counts read from the rule
  // modules rather than typed in -- the old hardcoded "19 rules" was already
  // wrong, and Macro/Sentiment linked to tabs that no longer exist.
  const PILLAR_INFO = {
    fundamentals: { icon: "📊", label: "Fundamentals", count: `${fund.ACTIVE_RULES.length} rules`, color: "from-violet-500 to-purple-500" },
    technicals:   { icon: "📈", label: "Technicals",   count: `${tech.ACTIVE_RULES.length} rules`, color: "from-sky-500 to-blue-500" },
  };
  const tabShortcuts = `
    <div>
      <div class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">Drill Deeper Into Each Pillar</div>
      <div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
        ${Object.entries(PILLAR_INFO).map(([tab, info]) => `
          <button class="composite-jump group relative overflow-hidden rounded-xl ring-1 ring-slate-200 hover:ring-slate-300 bg-white hover:shadow-md transition-all text-left p-3"
                  data-jump-tab="${tab}">
            <div class="absolute inset-0 bg-gradient-to-br ${info.color} opacity-0 group-hover:opacity-5 transition-opacity"></div>
            <div class="relative flex items-start gap-2.5">
              <div class="w-9 h-9 rounded-lg bg-gradient-to-br ${info.color} text-white flex items-center justify-center text-sm shadow-sm flex-shrink-0">${info.icon}</div>
              <div class="flex-1 min-w-0">
                <div class="font-bold text-slate-900 text-sm">${escapeHtml(info.label)}</div>
                <div class="text-[10px] text-slate-500">${info.count} → see breakdown</div>
              </div>
              <div class="text-slate-300 group-hover:text-slate-500 transition-colors text-sm">→</div>
            </div>
          </button>
        `).join("")}
      </div>
    </div>
  `;

  // Pillar radar — sits centred between the hero panel and pillar cards
  // when there's data to chart (skip on hard-fails / unrated to keep
  // the message focused on the exclusion reason).
  const hasPillarData = Object.values(s.pillars || {}).some((p) => p && p.raw != null);
  const radarSection = (!s.hardFails.length && !s.unrated && hasPillarData) ? `
    <div class="mb-6 bg-white rounded-2xl ring-1 ring-slate-200 p-5">
      <div class="flex items-baseline justify-between mb-2">
        <div class="text-xs font-bold uppercase tracking-wider text-slate-500">Pillar Shape</div>      </div>
      <div class="flex items-center justify-center pt-2 pb-3">
        ${renderPillarRadar(s, theme, 300)}
      </div>
    </div>
  ` : "";

  // "View Full Story" button — opens the magazine layout. Offered for
  // any rated stock that didn't hard-fail. Strong-buy gets a louder pill.
  const magazineCta = (!s.hardFails.length && !s.unrated) ? `
    <div class="mb-6 flex items-center justify-between gap-4 p-4 bg-gradient-to-br from-slate-50 to-indigo-50/60 rounded-2xl ring-1 ring-slate-200">
      <div class="flex-1 min-w-0">
        <div class="font-bold text-slate-900 text-sm">Magazine Brief</div>      </div>
      <button id="open-magazine" class="px-4 py-2 rounded-lg bg-gradient-to-r ${theme.from} ${theme.to} ${theme.textOn} text-sm font-bold shadow-sm hover:shadow-md transition-all flex items-center gap-1.5 flex-shrink-0">
        View Full Story <span class="text-base leading-none">→</span>
      </button>
    </div>
  ` : "";

  openModal(`
    ${heroHeader}
    <div class="p-6 max-h-[calc(90vh-200px)] overflow-y-auto">
      ${hardFailPanel}
      ${unratedPanel}
      ${heroPanel}
      ${radarSection}
      ${pillarCards}
      ${magazineCta}
      ${tabShortcuts}
    </div>
  `, { size: "wide" });

  // Entrance animation: gauge sweep, radar polygon scale-in, pillar
  // bars sweep, composite number counts up.
  requestAnimationFrame(() => animateScoreEntrance($("#modal-content")));

  $("#open-magazine")?.addEventListener("click", () => { closeModal(); openMagazineDrill(s); });
  $("#modal-close")?.addEventListener("click", closeModal);
  $("#modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "modal-overlay") closeModal();
  }, { once: true });
  // Drill-deeper button: close composite modal, switch to the target
  // pillar tab, then immediately open the per-company side panel for
  // THIS company on that tab (instead of dropping the user into an
  // unfocused table).
  $$(".composite-jump").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const targetTab = btn.dataset.jumpTab;
      const companyName = co.Company || co.name || "";
      closeModal();
      await switchTab(targetTab);
      // Find the score-row that corresponds to the same company on
      // the destination tab — different tabs key on Company vs name.
      const st = state.cache[targetTab];
      const target = CONFIGS[targetTab];
      const match = st?.scored.find((s) => {
        try { return target.name(s.company) === companyName; } catch { return false; }
      });
      if (match) openDrillDown(match);
    });
  });
}

function closeDrillDown() {
  $("#drill-panel").classList.add("translate-x-full");
  $("#drill-overlay").classList.add("hidden");
}

// ---------------- Centred modal (composite drill + help) ----------------
// Side-panel drill works for the per-pillar tabs (lots of rules,
// vertical scroll). The composite drill and the help screen need the
// full attention of the page, so they open in a centred modal.

function openModal(innerHtml, opts = {}) {
  const sizeClass = opts.size === "magazine" ? "max-w-6xl" : opts.size === "wide" ? "max-w-5xl" : "max-w-4xl";
  const container = $("#modal-container");
  container.className = `relative bg-white rounded-3xl shadow-2xl w-full ${sizeClass} my-8 scale-95 opacity-0 transition-all duration-200 overflow-hidden`;
  $("#modal-content").innerHTML = innerHtml;
  $("#modal-overlay").classList.add("is-open");
  $("#modal-overlay").classList.remove("hidden");
  // requestAnimationFrame so the entrance transition fires
  requestAnimationFrame(() => container.classList.replace("scale-95", "scale-100"));
  // ESC + click-outside close
  const onKey = (e) => { if (e.key === "Escape") closeModal(); };
  document.addEventListener("keydown", onKey);
  $("#modal-overlay").__onKey = onKey;
}
function closeModal() {
  const ov = $("#modal-overlay");
  ov.classList.remove("is-open");
  ov.classList.add("hidden");
  $("#modal-content").innerHTML = "";
  if (ov.__onKey) { document.removeEventListener("keydown", ov.__onKey); ov.__onKey = null; }
}

// Scoring-framework help content for each tab.
// Pillar tabs: list every rule the tab scores, grouped by category,
//              with the max points and pass criterion.
// Composite tab: pillar weighting + rating bands + decision framework
//                + every hard-fail rule across all pillars.
function buildHelpModal(tabId) {
  const cfgTab = CONFIGS[tabId];

  // AI Basket — composite-specific help: weights + bands + hard fails.
  if (cfgTab.composite) {
    const weightedRow = (label, weight, max, color) => `
      <div class="flex items-center gap-3 py-2.5 border-b border-slate-100 last:border-0">
        <div class="w-10 h-10 rounded-lg bg-gradient-to-br ${color} text-white flex items-center justify-center font-bold text-sm shadow-sm flex-shrink-0">${weight}%</div>
        <div class="flex-1 min-w-0">
          <div class="font-bold text-slate-900 text-sm">${label}</div>
          <div class="text-xs text-slate-500">Max ${max} raw pts → weighted contribution up to ${weight} pts</div>
        </div>
        <div class="text-xs text-slate-400 font-mono">×${weight}%</div>
      </div>
    `;
    const ratingRow = (rating, range, color, action) => `
      <div class="flex items-center gap-3 py-2 border-b border-slate-100 last:border-0">
        <div class="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${color} flex-shrink-0 min-w-[88px] text-center">${rating}</div>
        <div class="text-xs text-slate-500 font-mono flex-shrink-0 w-16">${range}</div>
        <div class="text-xs text-slate-700">${action}</div>
      </div>
    `;
    const hardFails = [
      { pillar: "Fundamentals", rule: "Operating Cash Flow", trigger: "Negative CFO in latest year" },
      { pillar: "Fundamentals", rule: "Interest Coverage",    trigger: "ICR < 1.5 (debt-serviceability risk)" },
      { pillar: "Fundamentals", rule: "Promoter Pledge",      trigger: "Pledge > 20%" },
      { pillar: "Fundamentals", rule: "Debt / Equity",        trigger: "D/E > 2 (extreme leverage)" },
      { pillar: "Fundamentals", rule: "Auditor Remarks",      trigger: "Adverse opinion / disclaimer" },
      { pillar: "Fundamentals", rule: "SEBI Governance",      trigger: "Active SEBI investigation" },
      { pillar: "Technicals",   rule: "Price Above 200 DMA",  trigger: "Stock below 200 DMA (primary trend filter)" },
      { pillar: "Liquidity",    rule: "Avg Daily Traded Value", trigger: "ADTV < ₹5 Cr (illiquid)" },
    ];

    return `
      <div class="relative bg-gradient-to-br from-indigo-500 via-purple-500 to-pink-500 p-6 text-white">
        <button id="modal-close" class="absolute top-3 right-3 text-white/80 hover:text-white text-2xl leading-none w-8 h-8 rounded-full bg-white/10 hover:bg-white/20 flex items-center justify-center transition-colors">×</button>
        <div class="text-xs font-semibold uppercase tracking-wider opacity-90">AI Basket — Composite Stock Selection</div>
        <h2 class="text-2xl font-bold mt-1">How the AI Basket Is Scored</h2>
        <p class="text-sm opacity-90 mt-1">Weighted composite of 5 pillars · rating bands per client framework · hard-fail rules exclude stocks from the basket entirely.</p>
      </div>
      <div class="p-6 max-h-[calc(90vh-180px)] overflow-y-auto">

        <div class="mb-6">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">Pillar Weights — composite formula</h3>
          <div class="bg-slate-50 rounded-xl p-4 ring-1 ring-slate-200">
            <div class="font-mono text-[11px] text-slate-700 text-center mb-3 leading-relaxed">
              Composite = (Fund/29)×40 + (Tech/24)×35 + (Macro/17)×15 + (Sent/6)×5 + (Liq/6)×5
            </div>
            ${weightedRow("Fundamentals — 19 rules", 40, 29, "from-violet-500 to-purple-500")}
            ${weightedRow("Technicals — 16 rules",    35, 24, "from-sky-500 to-blue-500")}
            ${weightedRow("Macro / Sector — 11 rules", 15, 17, "from-emerald-500 to-teal-500")}
            ${weightedRow("Sentiment — 4 rules",       5,  6,  "from-amber-500 to-orange-500")}
            ${weightedRow("Liquidity — 4 rules",       5,  6,  "from-rose-500 to-pink-500")}
          </div>
        </div>

        <div class="mb-6">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">Rating Bands — what each score band means</h3>
          <div class="bg-slate-50 rounded-xl p-4 ring-1 ring-slate-200">
            ${ratingRow("STRONG BUY", "≥ 75",   "bg-emerald-100 text-emerald-800 ring-1 ring-emerald-200", "Initiate FULL position · 8–10% of basket · monthly review")}
            ${ratingRow("BUY",        "60–74", "bg-blue-100 text-blue-800 ring-1 ring-blue-200",          "Initiate 50–75% position · 5–7% of basket · bi-weekly review")}
            ${ratingRow("WATCH",      "45–59", "bg-amber-100 text-amber-800 ring-1 ring-amber-200",       "Watchlist · do not initiate · weekly review")}
            ${ratingRow("AVOID",      "< 45",  "bg-rose-100 text-rose-800 ring-1 ring-rose-200",          "Exit position · zero allocation · immediate action")}
            ${ratingRow("HARD FAIL",  "—",     "bg-slate-200 text-slate-800 ring-1 ring-slate-300",       "Excluded from basket regardless of composite score")}
          </div>
        </div>

        <div>
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">Hard-Fail Rules — stock exits pipeline</h3>
          <div class="bg-rose-50/40 rounded-xl p-4 ring-1 ring-rose-200">
            <div class="text-xs text-rose-800 mb-3">If any of these trigger, the stock is excluded from the AI basket regardless of its composite score:</div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
              ${hardFails.map((hf) => `
                <div class="bg-white rounded-lg p-2.5 ring-1 ring-rose-100">
                  <div class="text-[10px] font-bold uppercase tracking-wider text-rose-500">${escapeHtml(hf.pillar)}</div>
                  <div class="text-sm font-bold text-rose-900 mt-0.5">${escapeHtml(hf.rule)}</div>
                  <div class="text-[11px] text-rose-700/90 mt-0.5">${escapeHtml(hf.trigger)}</div>
                </div>
              `).join("")}
            </div>
          </div>
        </div>

      </div>
    `;
  }

  // Pillar tabs (Fundamentals / Technicals / Macro / Sentiment & Liquidity):
  // list every rule with category, max points, and pass criterion.
  const rules = cfgTab.rules || [];
  const grouped = rules.reduce((acc, r) => { (acc[r.category] ||= []).push(r); return acc; }, {});
  const totalMax = rules.reduce((s, r) => s + ((r.fn ? r.fn({}).max : 0) || 0), 0) || cfgTab.stats?.maxScore?.match(/\d+/)?.[0] || 0;
  const headerColors = {
    fundamentals: "from-violet-500 via-purple-500 to-pink-500",
    technicals:   "from-sky-500 via-blue-500 to-indigo-500",
    macro:        "from-emerald-500 via-teal-500 to-cyan-500",
    sentiment:    "from-amber-500 via-orange-500 to-rose-500",
  };
  const grad = headerColors[tabId] || "from-slate-500 to-slate-700";

  // META holds the verbatim client scoring sheet text — surfaces the
  // full point-band breakdown (pass / partial / fail) per rule so the
  // user sees not just the headline threshold but the full structure.
  const metaForTab = RULE_META[tabId] || {};

  const ruleRow = (r) => {
    let max = 0;
    try { max = r.fn ? (r.fn({}).max || 0) : 0; } catch { max = 0; }
    // Max-point tier styling so 2 pts pops more than 1 pt.
    const pointStyle = max >= 2
      ? "bg-gradient-to-br from-indigo-500 to-purple-500 text-white shadow-sm"
      : max === 1
        ? "bg-indigo-100 text-indigo-700 ring-1 ring-indigo-200"
        : "bg-slate-100 text-slate-500";
    const clientLogic = metaForTab[r.key]?.clientLogic || "";
    return `
      <div class="flex items-start gap-3 py-3 border-b border-slate-100 last:border-0">
        <div class="flex-1 min-w-0">
          <div class="flex items-center justify-between gap-3 mb-1">
            <div class="font-semibold text-slate-900 text-sm">${escapeHtml(r.label)}</div>
            <div class="inline-flex items-baseline gap-0.5 px-2 py-0.5 rounded-md text-[11px] font-bold ${pointStyle} flex-shrink-0">
              <span>${max}</span>
              <span class="opacity-90">pt${max === 1 ? "" : "s"}</span>
              <span class="opacity-75 ml-0.5 text-[9px] font-semibold uppercase tracking-wider">max</span>
            </div>
          </div>
          <div class="text-[11px] text-slate-500"><span class="font-semibold text-slate-600">Threshold:</span> ${escapeHtml(r.criteria || "—")}</div>
          ${clientLogic ? `<div class="text-[11px] text-slate-600 mt-1 leading-snug bg-slate-50 rounded-md px-2 py-1.5 ring-1 ring-slate-100">${escapeHtml(clientLogic)}</div>` : ""}
        </div>
      </div>
    `;
  };

  return `
    <div class="relative bg-gradient-to-br ${grad} p-6 text-white">
      <button id="modal-close" class="absolute top-3 right-3 text-white/80 hover:text-white text-2xl leading-none w-8 h-8 rounded-full bg-white/10 hover:bg-white/20 flex items-center justify-center transition-colors">×</button>
      <div class="text-xs font-semibold uppercase tracking-wider opacity-90">Client framework</div>
      <h2 class="text-2xl font-bold mt-1">How the ${escapeHtml(cfgTab.label)} Tab Is Scored</h2>
      <p class="text-sm opacity-90 mt-1">${rules.length} rules across ${Object.keys(grouped).length} categories · max ${totalMax} pts · scored verbatim from the scoring framework.</p>
    </div>
    <div class="p-6 max-h-[calc(90vh-180px)] overflow-y-auto">
      ${Object.entries(grouped).map(([cat, rs]) => {
        const catMax = rs.reduce((s, r) => { try { return s + (r.fn ? (r.fn({}).max || 0) : 0); } catch { return s; } }, 0);
        return `
          <div class="mb-5">
            <div class="flex items-baseline justify-between mb-2">
              <h3 class="text-xs font-bold uppercase tracking-wider text-slate-500">${escapeHtml(cat)}</h3>
              <span class="text-[10px] text-slate-400 font-mono">${rs.length} rules · ${catMax} pts</span>
            </div>
            <div class="bg-slate-50 rounded-xl px-4 py-1 ring-1 ring-slate-200">
              ${rs.map(ruleRow).join("")}
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function openHelpModal() {
  openModal(buildHelpModal(state.activeTab));
  $("#modal-close")?.addEventListener("click", closeModal);
  $("#modal-overlay").addEventListener("click", (e) => {
    if (e.target.id === "modal-overlay") closeModal();
  }, { once: true });
}

// ---------------- Excel export (active tab) ----------------
async function exportToExcel() {
  const c = cfg(); const st = tabState();
  if (!st || !st.scored) return;

  let effCfg = c, effScored = st.scored, effRuleMeta = RULE_META;

  // Composite tab has cfg.rules=[] (it scores pillars, not rules), so the
  // Excel framework / distribution / ranked sheets came out empty. For
  // the export, synthesise an effective config that aggregates every rule
  // from all 4 pillars, plus merged per-company breakdowns drawn from
  // s.company._composite.pillarResults — so the workbook shows every rule
  // for every company with their actual values.
  if (state.activeTab === "composite") {
    const pillarSpecs = [
      { tag: "fundamentals", rules: fund.ACTIVE_RULES,   prKey: "fund"   },
      { tag: "technicals",   rules: tech.ACTIVE_RULES,   prKey: "tech"   },
    ];
    const allRules = [];
    for (const p of pillarSpecs) {
      for (const r of p.rules) allRules.push({ ...r, _pillar: p.tag });
    }
    effCfg = { ...c, rules: allRules };
    effScored = st.scored.map((s) => {
      const pr = s.company?._composite?.pillarResults;
      let merged = [];
      if (pr) {
        for (const p of pillarSpecs) {
          const arr = pr[p.prKey]?.breakdown;
          if (Array.isArray(arr)) merged = merged.concat(arr);
        }
      }
      return { ...s, breakdown: merged };
    });
    // Flatten per-pillar rule metas under a "composite" key so the export's
    // lookup `ruleMeta[tab][ruleKey]` finds clientLogic for every rule.
    const compositeMeta = {};
    for (const p of pillarSpecs) {
      const m = RULE_META[p.tag] || {};
      for (const k of Object.keys(m)) compositeMeta[k] = m[k];
    }
    effRuleMeta = { ...RULE_META, composite: compositeMeta };
  }

  await exportToExcelNew({
    tab: state.activeTab,
    tabLabel: c.label,
    cfg: effCfg,
    scored: effScored,
    ruleMeta: effRuleMeta,
  });
}

// ---------------- global search (header typeahead) ----------------
// Universal search across the whole universe. Selecting a company
// opens the composite drill so every pillar's data (fundamentals,
// technicals, macro, sentiment, liquidity) appears in one modal —
// no tab-switching required. Composite cache lazy-loads on first
// focus if the user hasn't visited the AI basket yet.
const globalSearch = { matches: [], selectedIdx: -1, lastQuery: "" };

async function ensureCompositeForSearch() {
  if (state.cache.composite) return true;
  const results = $("#global-search-results");
  if (results) {
    results.innerHTML = `<div class="px-4 py-6 text-center text-sm text-slate-400">Loading universe…</div>`;
    results.classList.remove("hidden");
  }
  try {
    await loadTab("composite");
    return true;
  } catch (e) {
    if (results) results.innerHTML = `<div class="px-4 py-6 text-center text-sm text-rose-500">Couldn't load — try again</div>`;
    return false;
  }
}

function setupGlobalSearch() {
  const input = $("#global-search");
  const results = $("#global-search-results");
  if (!input || !results) return;

  // Use Mac key symbol if user is on Mac, else "Ctrl K"
  const kbd = $("#global-search-kbd");
  if (kbd && !/Mac|iPhone|iPad/.test(navigator.platform || "")) kbd.textContent = "Ctrl K";

  function closeDropdown() {
    results.classList.add("hidden");
    globalSearch.matches = [];
    globalSearch.selectedIdx = -1;
  }

  function renderRow(s, i) {
    const name = s.company?.Company || "—";
    const sector = s.company?.Sector || s.company?.["Broad Industry"] || "";
    const marketCap = s.company?.["Market Cap"] || "";
    const { color, initials } = avatarFor(name);
    const cls = composite.ratingClass(s.rating);
    const comp = s.composite != null ? s.composite.toFixed(1) : "—";
    const active = i === globalSearch.selectedIdx;
    return `
      <button data-search-idx="${i}" class="result-row w-full text-left flex items-center gap-3 px-3.5 py-2.5 border-b border-slate-100 last:border-b-0 transition-colors ${active ? "bg-indigo-50" : "hover:bg-slate-50"}">
        <div class="w-9 h-9 rounded-lg bg-gradient-to-br ${color} flex items-center justify-center text-white font-bold text-xs flex-shrink-0">${initials}</div>
        <div class="min-w-0 flex-1">
          <div class="font-semibold text-slate-900 text-sm truncate">${escapeHtml(name)}</div>
          <div class="text-[11px] text-slate-500 truncate">${escapeHtml(sector)}${marketCap ? ` · ${escapeHtml(marketCap)}` : ""}</div>
        </div>
        <div class="text-right flex-shrink-0">
          <div class="text-base font-bold tabular-nums text-slate-900 leading-none">${comp}</div>
          <span class="inline-flex items-center mt-1 px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider ring-1 whitespace-nowrap ${cls}">${escapeHtml(s.rating || "—")}</span>
        </div>
      </button>
    `;
  }

  function repaint() {
    if (globalSearch.matches.length === 0) {
      results.innerHTML = `<div class="px-4 py-6 text-center text-sm text-slate-400">No companies found</div>`;
    } else {
      results.innerHTML =
        `<div class="px-3.5 py-2 text-[10px] font-bold uppercase tracking-wider text-slate-400 bg-slate-50/70 border-b border-slate-100">${globalSearch.matches.length} match${globalSearch.matches.length === 1 ? "" : "es"} · ↑↓ navigate · ↵ open</div>` +
        globalSearch.matches.map((s, i) => renderRow(s, i)).join("");
      results.querySelectorAll(".result-row").forEach((el) => {
        el.addEventListener("click", () => openSearchResult(Number(el.dataset.searchIdx)));
        el.addEventListener("mouseenter", () => {
          globalSearch.selectedIdx = Number(el.dataset.searchIdx);
          highlightSelected();
        });
      });
    }
    results.classList.remove("hidden");
  }

  function highlightSelected() {
    results.querySelectorAll(".result-row").forEach((el, i) => {
      const active = i === globalSearch.selectedIdx;
      el.classList.toggle("bg-indigo-50", active);
      el.classList.toggle("hover:bg-slate-50", !active);
      if (active) el.scrollIntoView({ block: "nearest" });
    });
  }

  function openSearchResult(idx) {
    const match = globalSearch.matches[idx];
    if (!match) return;
    closeDropdown();
    input.blur();
    input.value = "";
    openCompositeDrill(match);
  }

  async function update(q) {
    q = q.trim().toLowerCase();
    if (q.length < 1) { closeDropdown(); globalSearch.lastQuery = ""; return; }
    if (q === globalSearch.lastQuery) return;
    globalSearch.lastQuery = q;
    const ok = await ensureCompositeForSearch();
    if (!ok) return;
    const scored = state.cache.composite?.scored || [];
    // Rank: name starts-with first, then contains. Hide hard-failed at
    // the bottom but still surface them — analyst may want to inspect.
    const startsWith = [];
    const contains = [];
    for (const s of scored) {
      const n = (s.company?.Company || "").toLowerCase();
      if (!n) continue;
      if (n.startsWith(q)) startsWith.push(s);
      else if (n.includes(q)) contains.push(s);
    }
    globalSearch.matches = [...startsWith, ...contains].slice(0, 10);
    globalSearch.selectedIdx = globalSearch.matches.length > 0 ? 0 : -1;
    repaint();
  }

  let debounce;
  input.addEventListener("input", (e) => {
    clearTimeout(debounce);
    debounce = setTimeout(() => update(e.target.value), 80);
  });
  input.addEventListener("focus", () => {
    if (input.value.trim().length >= 1) update(input.value);
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { closeDropdown(); input.blur(); }
    else if (e.key === "ArrowDown") {
      e.preventDefault();
      if (globalSearch.matches.length > 0) {
        globalSearch.selectedIdx = (globalSearch.selectedIdx + 1) % globalSearch.matches.length;
        highlightSelected();
      }
    }
    else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (globalSearch.matches.length > 0) {
        globalSearch.selectedIdx = (globalSearch.selectedIdx - 1 + globalSearch.matches.length) % globalSearch.matches.length;
        highlightSelected();
      }
    }
    else if (e.key === "Enter") {
      e.preventDefault();
      if (globalSearch.selectedIdx >= 0) openSearchResult(globalSearch.selectedIdx);
    }
  });

  // Click outside closes the dropdown
  document.addEventListener("click", (e) => {
    if (!input.contains(e.target) && !results.contains(e.target)) closeDropdown();
  });

  // ⌘K / Ctrl+K focuses search from anywhere on the page
  document.addEventListener("keydown", (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") {
      e.preventDefault();
      input.focus();
      input.select();
    }
  });
}

// ---------------- wiring ----------------
function wire() {
  $$(".tab-btn").forEach((b) => b.addEventListener("click", () => switchTab(b.dataset.tab)));
  $("#search").addEventListener("input", (e) => { state.search = e.target.value; applyFilters(); });
  $("#score-filter").addEventListener("change", (e) => { state.scoreFilter = e.target.value; applyFilters(); });
  // Watchlist toggle — click toggles "show only watchlisted" mode
  const watchBtn = $("#watch-toggle");
  if (watchBtn) watchBtn.addEventListener("click", () => {
    state.watchOnly = !state.watchOnly;
    updateWatchCount();
    applyFilters();
  });
  updateWatchCount();
  $("#export-btn").addEventListener("click", exportToExcel);
  $("#top-picks-btn")?.addEventListener("click", () => switchTab("topPicks"));
  document.addEventListener("click", (e) => { if (e.target.closest("#tp-back-btn")) switchTab("composite"); });
  $("#drill-overlay").addEventListener("click", closeDrillDown);
  $("#help-btn")?.addEventListener("click", openHelpModal);
  $("#sources-btn")?.addEventListener("click", openSourcesModal);
  setupGlobalSearch();
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeDrillDown(); });
}

// ============================================================
// CUSTOM STRATEGY LAB (Phase 4)
// ============================================================
// User-defined strategies (a "playbook"): each carries its own
// composite threshold, basket size, and the three exit triggers
// (profit target / stop-loss / max holding) plus a rebalance cadence.
// Capital + charges are shared (simPrefs) so strategies compare on
// equal footing. Everything backtests over the snapshot history and
// reuses the Strategy tab's chart / KPI / per-pick / sector renderers.
const CUSTOM_STRATS_KEY = "klpdash-custom-strategies-v3";
const STRAT_FIELDS = [
  { key: "threshold",     label: "Composite ≥",    min: 50, max: 90,  step: 1,   suffix: "" },
  { key: "basketSize",    label: "Basket size",    min: 3,  max: 15,  step: 1,   suffix: " stocks" },
  { key: "target",        label: "Target",         min: 1,  max: 30,  step: 0.5, suffix: "%" },
  { key: "sl",            label: "Stop-loss",      min: 1,  max: 20,  step: 0.5, suffix: "%" },
  { key: "maxHoldDays",   label: "Max holding",    min: 5,  max: 120, step: 1,   suffix: " days" },
  { key: "rebalanceDays", label: "Rebalance every", min: 1, max: 30,  step: 1,   suffix: " days" },
];
function newStratId() { return "strat-" + Date.now().toString(36) + "-" + Math.floor(Math.random() * 1e4).toString(36); }
function defaultStrategy(name) {
  return { id: newStratId(), name: name || "New strategy", origin: "user", threshold: 75, basketSize: 7, target: 5, sl: 3, maxHoldDays: 30, rebalanceDays: 7, params: {} };
}

// Tweakable strategy parameters for the Custom Lab, across every pillar.
// `pillar` = the top-level category (picked from a dropdown); `cat` = the
// sub-group within it. `source` says where each value is read from:
//   tech   → stock.techVals[field]  (granular technicals, accrue forward)
//   fund   → stock.fundVals[field]  (fundamental ratios, accrue forward)
//   pillar → stock.pillars[field].pct (pillar score, backtests fully)
// A stock qualifies only if it passes every switched-on parameter (AND).
const STRAT_PARAMS = [
  // ── Technicals (granular) ──
  { key: "rsi",  pillar: "Technicals", cat: "Momentum", label: "RSI (14) between",      kind: "range", source: "tech", field: "rsi",  def: { min: 55, max: 75 }, step: 1 },
  { key: "adx",  pillar: "Technicals", cat: "Momentum", label: "ADX (14) ≥",            kind: "min",   source: "tech", field: "adx",  def: { val: 25 },  step: 1 },
  { key: "macd", pillar: "Technicals", cat: "Momentum", label: "MACD positive",         kind: "bool",  source: "tech", field: "macd" },
  { key: "rs",   pillar: "Technicals", cat: "Momentum", label: "Outperforming Nifty",   kind: "bool",  source: "tech", field: "rs" },
  { key: "vol",  pillar: "Technicals", cat: "Volume",   label: "Volume ≥ (× avg)",      kind: "min",   source: "tech", field: "vol",  def: { val: 1.5 }, step: 0.1 },
  { key: "dlv",  pillar: "Technicals", cat: "Volume",   label: "Rising delivery %",     kind: "bool",  source: "tech", field: "dlv" },
  { key: "inst", pillar: "Technicals", cat: "Volume",   label: "Net FII+DII buying",    kind: "bool",  source: "tech", field: "inst" },
  { key: "d52",  pillar: "Technicals", cat: "Breakout", label: "Within % of 52w high",  kind: "max",   source: "tech", field: "d52",  def: { val: 10 },  step: 0.5 },
  { key: "cons", pillar: "Technicals", cat: "Breakout", label: "Consolidation breakout", kind: "bool", source: "tech", field: "cons" },
  { key: "base", pillar: "Technicals", cat: "Breakout", label: "Base formation",        kind: "bool",  source: "tech", field: "base" },
  { key: "e50",  pillar: "Technicals", cat: "Trend",    label: "Price above 50 EMA",    kind: "bool",  source: "tech", field: "e50" },
  { key: "d200", pillar: "Technicals", cat: "Trend",    label: "Price above 200 DMA",   kind: "bool",  source: "tech", field: "d200" },
  { key: "gc",   pillar: "Technicals", cat: "Trend",    label: "Golden Cross",          kind: "bool",  source: "tech", field: "gc" },
  { key: "hh",   pillar: "Technicals", cat: "Trend",    label: "Higher Highs–Lows",     kind: "bool",  source: "tech", field: "hh" },
  { key: "beta", pillar: "Technicals", cat: "Risk",     label: "Beta between",          kind: "range", source: "tech", field: "beta", def: { min: 0.7, max: 1.3 }, step: 0.1 },
  { key: "atr",  pillar: "Technicals", cat: "Risk",     label: "ATR < (% of price)",    kind: "max",   source: "tech", field: "atr",  def: { val: 2.5 }, step: 0.1 },
  // ── Fundamentals ──
  { key: "fundScore", pillar: "Fundamentals", cat: "Overall",       label: "Fundamentals score ≥ (%)", kind: "min", source: "pillar", field: "fundamentals", def: { val: 60 }, step: 1 },
  { key: "roe",       pillar: "Fundamentals", cat: "Profitability", label: "ROE ≥ (%)",   kind: "min", source: "fund", field: "roe",    def: { val: 15 }, step: 1 },
  { key: "roce",      pillar: "Fundamentals", cat: "Profitability", label: "ROCE ≥ (%)",  kind: "min", source: "fund", field: "roce",   def: { val: 15 }, step: 1 },
  { key: "de",        pillar: "Fundamentals", cat: "Balance sheet", label: "Debt / Equity ≤", kind: "max", source: "fund", field: "de", def: { val: 1 },  step: 0.1 },
  { key: "pe",        pillar: "Fundamentals", cat: "Valuation",     label: "P/E ≤",       kind: "max", source: "fund", field: "pe",     def: { val: 40 }, step: 1 },
  { key: "salesG",    pillar: "Fundamentals", cat: "Growth",        label: "Sales growth 3Y ≥ (%)", kind: "min", source: "fund", field: "salesG", def: { val: 10 }, step: 1 },
  { key: "patG",      pillar: "Fundamentals", cat: "Growth",        label: "Profit growth 3Y ≥ (%)", kind: "min", source: "fund", field: "patG", def: { val: 10 }, step: 1 },
  // ── Sentiment ──
  { key: "sentScore", pillar: "Sentiment", cat: "Overall", label: "Sentiment score ≥ (%)", kind: "min", source: "pillar", field: "sentiment", def: { val: 55 }, step: 1 },
  // ── Macro ──
  { key: "macroScore", pillar: "Macro", cat: "Overall", label: "Macro score ≥ (%)", kind: "min", source: "pillar", field: "macro", def: { val: 55 }, step: 1 },
  // ── Liquidity ──
  { key: "liqScore", pillar: "Liquidity", cat: "Overall", label: "Liquidity score ≥ (%)", kind: "min", source: "pillar", field: "liquidity", def: { val: 55 }, step: 1 },
];
const PARAM_PILLARS = ["Technicals", "Fundamentals", "Sentiment", "Macro", "Liquidity"];

// Read a parameter's value off a snapshot stock, by source.
function paramValue(stock, def) {
  if (!stock) return null;
  if (def.source === "pillar") return stock.pillars?.[def.field]?.pct ?? null;
  if (def.source === "tech")   return stock.techVals?.[def.field] ?? null;
  if (def.source === "fund")   return stock.fundVals?.[def.field] ?? null;
  return null;
}

// AND-gate: a stock passes only if every switched-on parameter is met.
// Params whose source record is absent on a given day are SKIPPED (not
// failed), so pillar-score params backtest fully while granular tech/fund
// params only bind once the day carries their record (accrue forward).
function passesParams(stock, params) {
  const keys = Object.keys(params || {}).filter((k) => params[k]?.on);
  if (!keys.length) return true;
  for (const key of keys) {
    const def = STRAT_PARAMS.find((p) => p.key === key); if (!def) continue;
    if (def.source === "tech" && !stock.techVals) continue;   // day lacks the record → skip
    if (def.source === "fund" && !stock.fundVals) continue;
    const v = paramValue(stock, def); const c = params[key];
    if (def.kind === "bool") { if (v !== true) return false; }
    else if (v == null) return false;
    else if (def.kind === "range") { if (v < c.min || v > c.max) return false; }
    else if (def.kind === "min") { if (v < c.val) return false; }
    else if (def.kind === "max") { if (v > c.val) return false; }
  }
  return true;
}
function activeParamCount(params) { return Object.keys(params || {}).filter((k) => params[k]?.on).length; }
function activeParamCountForPillar(params, pillar) {
  return Object.keys(params || {}).filter((k) => params[k]?.on && STRAT_PARAMS.find((p) => p.key === k)?.pillar === pillar).length;
}
// Seeded from an offline grid search (screener-test/grid-search-strategies.mjs)
// over ~1,000 parameter combinations, backtested on the snapshot history:
// the best performer in each of 5 distinct styles. The user can add / edit
// / delete freely — these are just a strong starting playbook.
function seedStrategies() {
  const mk = (name, o) => ({ ...defaultStrategy(name), ...o, origin: "ai" });
  return [
    mk("★ Concentrated — top 5 names",       { threshold: 75, basketSize: 5,  target: 10, sl: 3, maxHoldDays: 15, rebalanceDays: 10 }),
    mk("Weekly rotation — top 5",            { threshold: 75, basketSize: 5,  target: 8,  sl: 5, maxHoldDays: 15, rebalanceDays: 7 }),
    mk("Quick rotation — top 7, 3-day",      { threshold: 75, basketSize: 7,  target: 8,  sl: 3, maxHoldDays: 20, rebalanceDays: 3 }),
    mk("Broad & steady — top 10",            { threshold: 70, basketSize: 10, target: 20, sl: 8, maxHoldDays: 20, rebalanceDays: 10 }),
    mk("Balanced — top 7, 5-day",            { threshold: 75, basketSize: 7,  target: 8,  sl: 5, maxHoldDays: 15, rebalanceDays: 5 }),
  ];
}
function loadStrategies() {
  try { const raw = JSON.parse(localStorage.getItem(CUSTOM_STRATS_KEY)); if (Array.isArray(raw) && raw.length) return raw; } catch {}
  return seedStrategies();
}
function saveStrategies(list) { try { localStorage.setItem(CUSTOM_STRATS_KEY, JSON.stringify(list)); } catch {} }
let customStrategies = loadStrategies();
let customSelectedId = null;
let customPerfOpen = false;   // showing the Performance comparison page?

function lastRet(curve) { for (let i = (curve?.length || 0) - 1; i >= 0; i--) if (curve[i].retPct != null) return curve[i].retPct; return 0; }
function fmtSignedPct(v) { return (v >= 0 ? "+" : "") + v.toFixed(2) + "%"; }

// Trailing-window returns off a cumulative-return curve ([{date, retPct}]).
// Returns are compounded on the equity FACTOR (1 + retPct/100), so the
// window return is the growth between the two points, not a subtraction of
// cumulative percentages. null when there isn't enough history to fill it.
function shiftDateStr(d, deltaDays) {
  const dt = new Date(d + "T00:00:00Z");
  dt.setUTCDate(dt.getUTCDate() + deltaDays);
  return dt.toISOString().slice(0, 10);
}
function windowReturnBetween(curve, fromRetPct, toRetPct) {
  return ((1 + toRetPct / 100) / (1 + fromRetPct / 100) - 1) * 100;
}
function oneDayReturnPct(curve) {
  if (!curve || curve.length < 2) return null;
  return windowReturnBetween(curve, curve[curve.length - 2].retPct, curve[curve.length - 1].retPct);
}
function trailingReturnPct(curve, calendarDays) {
  if (!curve || curve.length < 2) return null;
  const last = curve[curve.length - 1];
  const cutoff = shiftDateStr(last.date, -calendarDays);
  if (curve[0].date > cutoff) return null;   // window predates inception
  let base = null;
  for (const p of curve) { if (p.date <= cutoff) base = p; else break; }
  if (!base) return null;
  return windowReturnBetween(curve, base.retPct, last.retPct);
}

// Custom-strategy backtest. Walks the snapshot trail from the upload
// anchor holding up to N stocks (composite ≥ threshold). Each day every
// holding is checked against the three exit triggers — whichever hits
// first books a SELL; on a rebalance day, holdings that fall out of the
// top-N qualifying set are also dropped. Freed slots refill with the
// best qualifiers not held. Capital, buffer and per-side charges come
// from simPrefs (shared across strategies).
function simulateCustomStrategy(snapshots, anchorDate, strat, simP = simPrefs) {
  if (!snapshots?.length) return null;
  let sorted = [...snapshots].sort((a, b) => a.date.localeCompare(b.date));
  if (anchorDate) sorted = sorted.filter((s) => s.date >= anchorDate);
  if (!sorted.length) return null;

  const N = Math.max(1, Math.round(strat.basketSize || 7));
  const thr = strat.threshold ?? 75;
  const targetPct = (strat.target ?? 5) / 100;
  const slPct = (strat.sl ?? 3) / 100;
  const maxHold = Math.max(1, Math.round(strat.maxHoldDays || 30));
  const rebalDays = Math.max(1, Math.round(strat.rebalanceDays || 7));
  const params = strat.params || {};
  const hasParams = activeParamCount(params) > 0;
  const capital = simP?.capital ?? ACTIVE_INITIAL_CAPITAL;
  const bufferAmt = capital * Math.max(0, simP?.bufferPct ?? 0) / 100;
  const chg = simP === ZERO_CHARGER ? ZERO_CHARGER : makeCharger(simP);

  let cash = capital, totalCharges = 0;
  let lastRebal = sorted[0].date;
  const holdings = new Map();   // ticker → { units, entryDate, entryPrice, name, sector }
  const trades = [];
  const equity = [];

  for (const snap of sorted) {
    const date = snap.date;
    const closeBy = new Map();
    for (const s of snap.stocks) if (typeof s.close === "number") closeBy.set(s.ticker, s.close);

    const isRebalanceDay = date === sorted[0].date || daysBetween(lastRebal, date) >= rebalDays;
    if (isRebalanceDay) lastRebal = date;

    const qualifying = snap.stocks
      .filter((s) => s.composite != null && s.composite >= thr && s.dataComplete && !s.hardFailed && typeof s.close === "number" && s.ticker)
      // Parameter AND-gate. passesParams reads pillar scores (always
      // present → full backtest) and granular tech/fund values (present
      // forward → accrue), skipping any param whose record a given day lacks.
      .filter((s) => !hasParams || passesParams(s, params))
      .sort((a, b) => b.composite - a.composite);
    const topN = qualifying.slice(0, N);
    const topNset = new Set(topN.map((s) => s.ticker));
    const exitedToday = new Set();

    // 1. Exit checks — three triggers, plus rebalance de-qualification.
    for (const [ticker, pos] of [...holdings.entries()]) {
      const px = closeBy.get(ticker);
      if (typeof px !== "number") continue;
      const gain = px / pos.entryPrice - 1;
      let reason = null;
      if (gain >= targetPct) reason = "TARGET";
      else if (gain <= -slPct) reason = "SL";
      else if (daysBetween(pos.entryDate, date) >= maxHold) reason = "TIME";
      else if (isRebalanceDay && !topNset.has(ticker)) reason = "REBAL";
      if (!reason) continue;
      const gross = pos.units * px;
      const fee = chg.sell(gross);
      cash += gross - fee; totalCharges += fee;
      trades.push({ action: "SELL", ticker, name: pos.name, sector: pos.sector, date, price: px, entryDate: pos.entryDate, entryPrice: pos.entryPrice, days: daysBetween(pos.entryDate, date), ret: (px / pos.entryPrice - 1) * 100, reason });
      holdings.delete(ticker);
      exitedToday.add(ticker);
    }

    // 2. Refill empty slots with best qualifiers not held / not just exited.
    if (holdings.size < N) {
      for (const s of topN) {
        if (holdings.size >= N) break;
        if (holdings.has(s.ticker) || exitedToday.has(s.ticker)) continue;
        let nav = cash;
        for (const [t, p] of holdings) nav += p.units * (closeBy.get(t) ?? p.entryPrice);
        const slot = Math.max(0, nav - bufferAmt) / N;
        const buyValue = Math.min(slot, Math.max(0, cash) / (1 + buyRate(chg.prefs || {})));
        if (buyValue < 0.01) break;
        const fee = chg.buy(buyValue);
        holdings.set(s.ticker, { units: buyValue / s.close, entryDate: date, entryPrice: s.close, name: s.name || s.ticker, sector: s.sector || null });
        trades.push({ action: "BUY", ticker: s.ticker, name: s.name || s.ticker, sector: s.sector || null, date, price: s.close });
        cash -= buyValue + fee; totalCharges += fee;
      }
    }

    // 3. Mark to market.
    let value = cash;
    for (const [t, p] of holdings) value += p.units * (closeBy.get(t) ?? p.entryPrice);
    equity.push({ date, value, holdingCount: holdings.size, cash });
  }
  return { equity, trades, holdings, startCapital: capital, totalCharges };
}

function buildCustomPicks(sim, snapshots, todayDate, strat) {
  const picks = [];
  const tPct = (strat.target ?? 5) / 100, sPct = (strat.sl ?? 3) / 100;
  for (const t of sim.trades) {
    if (t.action !== "BUY") continue;
    const entryPrice = t.price;
    const target = entryPrice * (1 + tPct), sl = entryPrice * (1 - sPct);
    const status = computeHitStatus(t.ticker, t.date, entryPrice, target, sl, snapshots, todayDate);
    const peak = computePeakStats(t.ticker, t.date, entryPrice, snapshots, todayDate);
    picks.push({ ticker: t.ticker, name: t.name || t.ticker, sector: t.sector || null, entryDate: t.date, entryPrice, target, sl, targetPct: tPct * 100, slPct: -sPct * 100, ...status, peak });
  }
  return enrichAndSortPicks(attachIndustry(picks, snapshots));
}

// Produces a view object compatible with the Strategy tab renderers
// (renderStrategyKpis / renderActivePickRowsSplit / renderSectorTiming /
// renderSimPanel), so a custom strategy's deep-dive reuses all of them.
function buildCustomView(snapshots, anchorDate, todayDate, strat, niftyOn, manualPicks) {
  const sim = simulateCustomStrategy(snapshots, anchorDate, strat, simPrefs);
  if (!sim || !sim.equity.length) return null;
  const simGross = simulateCustomStrategy(snapshots, anchorDate, strat, ZERO_CHARGER);
  const picks = buildCustomPicks(sim, snapshots, todayDate, strat);
  const hitSummary = computeOverallHitSummary(picks);
  const equityCurve = sim.equity.map((e) => ({ date: e.date, retPct: (e.value / sim.startCapital - 1) * 100 }));
  const finalReturn = equityCurve[equityCurve.length - 1].retPct;
  const grossEnd = simGross.equity[simGross.equity.length - 1];
  const grossFinalReturn = (grossEnd.value / simGross.startCapital - 1) * 100;
  const dates = equityCurve.map((e) => e.date);
  const niftyCurve = buildNiftyCurve(dates, niftyOn);
  const { manualRows, manualSummary, manualCurve, manualBooked } = buildManualBundle(manualPicks, snapshots, anchorDate, todayDate, dates);
  const niftyRet = niftyCurve.length ? niftyCurve[niftyCurve.length - 1].retPct : null;
  return {
    kind: "custom", strat, sim, picks, hitSummary,
    equityCurve, niftyCurve, manualCurve, manualPicks: manualRows, manualSummary, manualBooked,
    periodLabel: `${strat.name} · from ${fmtDateDMY(anchorDate)}`,
    finalReturn, niftyRet, alpha: niftyRet != null ? finalReturn - niftyRet : null,
    manualFinalReturn: manualCurve.length ? (manualCurve[manualCurve.length - 1].retPct ?? null) : null,
    startDate: equityCurve[0].date, endDate: equityCurve[equityCurve.length - 1].date,
    finalValue: sim.equity[sim.equity.length - 1].value, startCapital: sim.startCapital,
    grossFinalReturn, totalCharges: sim.totalCharges,
    tradeCount: sim.trades.length, liveHoldings: sim.holdings.size,
  };
}

// Lightweight multi-curve line chart (own markup, no shared IDs) — used
// for the overview comparison and each strategy's deep-dive.
let _mccCtx = null;   // geometry + series for the multi-curve hover handler
function renderMultiCurveChart(series, title, subtitle) {
  const valid = (series || []).filter((s) => s.curve && s.curve.some((p) => p.retPct != null));
  if (!valid.length) { _mccCtx = null; return `<div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5 text-xs text-slate-500">No curve data yet.</div>`; }
  const W = 820, H = 240, M = { left: 48, right: 16, top: 14, bottom: 26 };
  const innerW = W - M.left - M.right, innerH = H - M.top - M.bottom;
  const dates = valid.reduce((best, s) => s.curve.length > best.length ? s.curve.map((p) => p.date) : best, []);
  const allV = valid.flatMap((s) => s.curve.map((p) => p.retPct).filter((v) => v != null)).concat([0]);
  let lo = Math.min(...allV), hi = Math.max(...allV);
  const pad = (hi - lo) * 0.12 || 1; lo -= pad; hi += pad;
  const n = dates.length;
  const xAt = (i) => M.left + (i / Math.max(1, n - 1)) * innerW;
  const yAt = (v) => M.top + innerH - ((v - lo) / (hi - lo)) * innerH;
  const grid = [0, 0.25, 0.5, 0.75, 1].map((t) => {
    const v = lo + (hi - lo) * t, y = (M.top + innerH - t * innerH).toFixed(1);
    return `<line x1="${M.left}" x2="${W - M.right}" y1="${y}" y2="${y}" stroke="#e2e8f0" stroke-width="0.7" stroke-dasharray="3 4"/><text x="${M.left - 8}" y="${y}" text-anchor="end" dominant-baseline="middle" font-size="9.5" fill="#94a3b8">${v >= 0 ? "+" : ""}${v.toFixed(1)}%</text>`;
  }).join("");
  const paths = valid.map((s) => {
    const byDate = new Map(s.curve.map((p) => [p.date, p.retPct]));
    let d = "", started = false;
    dates.forEach((dt, i) => { const v = byDate.get(dt); if (v == null) return; d += `${started ? "L" : "M"} ${xAt(i).toFixed(1)} ${yAt(v).toFixed(1)} `; started = true; });
    return `<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.width || 2}" stroke-linejoin="round" stroke-linecap="round" ${s.dash ? `stroke-dasharray="${s.dash}"` : ""}/>`;
  }).join("");
  const tickEvery = Math.max(1, Math.ceil(n / 7));
  const xticks = dates.map((dt, i) => (i % tickEvery !== 0 && i !== n - 1) ? "" : `<text x="${xAt(i).toFixed(1)}" y="${H - 8}" text-anchor="middle" font-size="9.5" fill="#64748b">${dt.slice(5)}</text>`).join("");
  const legend = valid.map((s) => { const r = lastRet(s.curve); return `<span class="inline-flex items-center gap-1.5 text-[11px] text-slate-600"><span class="w-3 h-0.5 rounded" style="background:${s.color}"></span>${escapeHtml(s.label)} <span class="tabular-nums font-semibold ${r >= 0 ? "text-emerald-700" : "text-rose-700"}">${fmtSignedPct(r)}</span></span>`; }).join("");
  _mccCtx = { W, H, M, innerW, innerH, lo, hi, dates, series: valid.map((s) => ({ label: s.label, color: s.color, byDate: new Map(s.curve.map((p) => [p.date, p.retPct])) })) };
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-baseline justify-between gap-2 flex-wrap mb-2">
        <div><h3 class="font-display font-bold text-slate-900 text-sm">${escapeHtml(title)}</h3>${subtitle ? `<div class="text-[11px] text-slate-500">${escapeHtml(subtitle)}</div>` : ""}</div>
        <div class="flex flex-wrap items-center gap-x-3 gap-y-1">${legend}</div>
      </div>
      <div id="mcc-wrap" class="relative">
        <svg id="mcc-svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" class="w-full select-none" style="max-height:260px">
          ${grid}
          <line x1="${M.left}" x2="${W - M.right}" y1="${yAt(0).toFixed(1)}" y2="${yAt(0).toFixed(1)}" stroke="#cbd5e1" stroke-width="1"/>
          ${paths}${xticks}
          <line id="mcc-guide" x1="0" y1="${M.top}" x2="0" y2="${(M.top + innerH).toFixed(1)}" stroke="#94a3b8" stroke-width="0.8" stroke-dasharray="2 3" opacity="0"/>
          <g id="mcc-dots"></g>
          <rect id="mcc-capture" x="${M.left}" y="${M.top}" width="${innerW}" height="${innerH}" fill="transparent" style="cursor:crosshair"/>
        </svg>
        <div id="mcc-tip" class="hidden absolute z-10 pointer-events-none -translate-x-1/2 -translate-y-[calc(100%+10px)] bg-slate-900/95 backdrop-blur text-white text-[11px] rounded-xl shadow-2xl ring-1 ring-slate-700/60 px-3 py-2 whitespace-nowrap space-y-0.5"></div>
      </div>
    </div>`;
}

// Hover crosshair + tooltip for the multi-curve chart. Reads geometry
// from _mccCtx (set on the last render); only one chart is live at a time.
function wireMultiCurveHover(root) {
  const ctx = _mccCtx; if (!ctx) return;
  const wrap = $(`${root} #mcc-wrap`); if (!wrap) return;
  const svg = wrap.querySelector("#mcc-svg"), capture = wrap.querySelector("#mcc-capture");
  const guide = wrap.querySelector("#mcc-guide"), dots = wrap.querySelector("#mcc-dots"), tip = wrap.querySelector("#mcc-tip");
  if (!svg || !capture || !tip) return;
  const { W, M, innerW, innerH, lo, hi, dates, series } = ctx;
  const n = dates.length;
  const xAt = (i) => M.left + (i / Math.max(1, n - 1)) * innerW;
  const yAt = (v) => M.top + innerH - ((v - lo) / (hi - lo)) * innerH;
  function move(clientX) {
    // Map through the SVG's live CTM so preserveAspectRatio letterboxing,
    // page zoom and device-pixel-ratio can't shift the pointer→date mapping
    // (the plain clientX/rect.width math drifts once the chart is padded).
    const ctm = svg.getScreenCTM();
    let vx;
    if (ctm) {
      const sp = svg.createSVGPoint(); sp.x = clientX; sp.y = 0;
      vx = sp.matrixTransform(ctm.inverse()).x;
    } else {
      const rect = svg.getBoundingClientRect();
      vx = (clientX - rect.left) / rect.width * W;
    }
    let i = Math.round((vx - M.left) / innerW * (n - 1));
    i = Math.max(0, Math.min(n - 1, i));
    const dt = dates[i], gx = xAt(i);
    guide.setAttribute("x1", gx); guide.setAttribute("x2", gx); guide.setAttribute("opacity", "1");
    let dotsHtml = "", rows = "", topY = ctx.H;
    for (const s of series) {
      const v = s.byDate.get(dt); if (v == null) continue;
      const dy = yAt(v); topY = Math.min(topY, dy);
      dotsHtml += `<circle cx="${gx.toFixed(1)}" cy="${dy.toFixed(1)}" r="3.5" fill="#fff" stroke="${s.color}" stroke-width="2"/>`;
      rows += `<div class="flex items-center gap-1.5"><span class="w-2 h-2 rounded-full" style="background:${s.color}"></span><span>${escapeHtml(s.label)}</span><span class="ml-3 tabular-nums font-semibold ${v >= 0 ? "text-emerald-300" : "text-rose-300"}">${v >= 0 ? "+" : ""}${v.toFixed(2)}%</span></div>`;
    }
    dots.innerHTML = dotsHtml;
    tip.innerHTML = `<div class="font-bold">${dt}</div>${rows}`;
    tip.classList.remove("hidden");
    const wrapRect = wrap.getBoundingClientRect();
    let tipX, tipY;
    if (ctm) {
      const sp2 = svg.createSVGPoint(); sp2.x = gx; sp2.y = topY;
      const scr = sp2.matrixTransform(ctm);
      tipX = scr.x - wrapRect.left; tipY = scr.y - wrapRect.top;
    } else {
      const rect = svg.getBoundingClientRect(); const scale = rect.width / W;
      tipX = gx * scale; tipY = topY * scale;
    }
    tip.style.left = tipX + "px";
    tip.style.top = tipY + "px";
    // Flip / clamp so the last date's tooltip doesn't clip off the right edge.
    tip.style.transform = "translate(-50%, calc(-100% - 10px))";
    requestAnimationFrame(() => {
      const tr = tip.getBoundingClientRect();
      let dx = 0;
      if (tr.left < wrapRect.left + 6) dx = (wrapRect.left + 6) - tr.left;
      else if (tr.right > wrapRect.right - 6) dx = (wrapRect.right - 6) - tr.right;
      if (dx !== 0) tip.style.transform = `translate(calc(-50% + ${dx}px), calc(-100% - 10px))`;
    });
  }
  capture.addEventListener("mousemove", (e) => move(e.clientX));
  capture.addEventListener("mouseleave", () => { guide.setAttribute("opacity", "0"); dots.innerHTML = ""; tip.classList.add("hidden"); });
  capture.addEventListener("touchmove", (e) => { if (e.touches[0]) move(e.touches[0].clientX); }, { passive: true });
}

function stratChip(t) { return `<span class="inline-flex items-center px-1.5 py-0 rounded bg-slate-100 text-slate-600 ring-1 ring-slate-200 text-[9px] font-semibold tabular-nums">${t}</span>`; }

function renderStrategyCard(strat, view, color) {
  const net = view?.finalReturn;
  const dd = view ? curveMaxDrawdown(view.equityCurve) : null;
  const up = view ? curveMaxUpside(view.equityCurve) : null;
  const charges = view?.totalCharges ?? 0;
  const netCls = net == null ? "text-slate-500" : net >= 0 ? "text-emerald-700" : "text-rose-700";
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 flex flex-col gap-3">
      <div class="flex items-start justify-between gap-2">
        <div class="flex items-center gap-2 min-w-0">
          <span class="w-2.5 h-2.5 rounded-full flex-shrink-0" style="background:${color}"></span>
          <div class="font-display font-bold text-slate-900 text-sm truncate" title="${escapeHtml(strat.name)}">${escapeHtml(strat.name)}</div>
        </div>
        <div class="${netCls} text-xl font-bold tabular-nums leading-none">${net == null ? "—" : fmtSignedPct(net)}</div>
      </div>
      <div class="flex flex-wrap gap-1">
        ${stratChip(`T ${strat.target}%`)} ${stratChip(`SL ${strat.sl}%`)} ${stratChip(`${strat.maxHoldDays}d hold`)} ${stratChip(`rebal ${strat.rebalanceDays}d`)} ${stratChip(`top ${strat.basketSize}`)} ${stratChip(`≥${strat.threshold}`)}
      </div>
      <div class="grid grid-cols-3 gap-2 text-center">
        <div><div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">Max up</div><div class="text-xs font-bold tabular-nums text-emerald-700">${up == null ? "—" : fmtSignedPct(up)}</div></div>
        <div><div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">Max DD</div><div class="text-xs font-bold tabular-nums text-rose-700">${dd == null ? "—" : fmtSignedPct(dd)}</div></div>
        <div><div class="text-[9px] uppercase tracking-wider text-slate-400 font-bold">Charges</div><div class="text-xs font-bold tabular-nums text-slate-700">₹${Math.round(charges).toLocaleString("en-IN")}</div></div>
      </div>
      <div class="flex items-center gap-2 mt-auto pt-1">
        <button type="button" data-strat-open="${strat.id}" class="flex-1 px-3 py-1.5 rounded-lg bg-indigo-50 text-indigo-700 ring-1 ring-indigo-200 text-xs font-semibold hover:bg-indigo-100">Open</button>
        <button type="button" data-strat-dup="${strat.id}" class="px-2.5 py-1.5 rounded-lg ring-1 ring-slate-200 text-xs font-semibold text-slate-600 hover:bg-slate-50">Duplicate</button>
        <button type="button" data-strat-del="${strat.id}" class="px-2.5 py-1.5 rounded-lg ring-1 ring-slate-200 text-xs font-semibold text-slate-500 hover:text-rose-600 hover:bg-rose-50">Delete</button>
      </div>
    </div>`;
}

// Today's granular technicals (current per-indicator values), fetched
// once and joined by ticker. Lets the indicator picker evaluate live
// pass/fail per stock using the real tech-scoring rule functions.
let customTechByTicker = null;

// The names this strategy would pick RIGHT NOW: composite ≥ threshold,
// passing every switched-on technical parameter (evaluated on today's
// values), ranked by composite, capped at basket size.
function buildTodaysNames(strat, latestSnap) {
  const N = Math.max(1, Math.round(strat.basketSize || 7));
  const thr = strat.threshold ?? 75;
  const params = strat.params || {};
  const rows = [];
  for (const s of (latestSnap?.stocks || [])) {
    if (s.composite == null || s.composite < thr || !s.dataComplete || s.hardFailed) continue;
    if (!passesParams(s, params)) continue;
    rows.push({ ticker: s.ticker, name: s.name || s.ticker, sector: s.sector, composite: s.composite });
  }
  rows.sort((a, b) => b.composite - a.composite);
  return { picks: rows.slice(0, N), qualifyingCount: rows.length };
}

function renderTodaysNames(result, strat) {
  const n = activeParamCount(strat.params);
  const head = n
    ? `${result.qualifyingCount} stock${result.qualifyingCount === 1 ? "" : "s"} pass composite ≥ ${strat.threshold} AND all ${n} filter${n === 1 ? "" : "s"} you set — today`
    : `Top ${strat.basketSize} by composite ≥ ${strat.threshold} today (no indicator filters set)`;
  const rows = result.picks.map((p) => `
    <div class="flex items-center justify-between gap-2 py-1.5 px-2 rounded-lg hover:bg-slate-50">
      <div class="min-w-0"><div class="font-semibold text-slate-800 text-xs truncate" title="${escapeHtml(p.name)}">${escapeHtml(p.name)}</div><div class="text-[10px] text-slate-400 truncate">${escapeHtml(p.sector || "")}</div></div>
      <div class="text-right"><div class="text-xs font-bold tabular-nums text-indigo-700">${p.composite.toFixed(1)}</div><div class="text-[9px] text-slate-400 uppercase tracking-wider">score</div></div>
    </div>`).join("") || `<div class="text-xs text-slate-500 px-2 py-3">No stocks pass all your filters today — loosen a threshold or turn one off.</div>`;
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-center gap-2 mb-1"><span class="text-base">🎯</span><h3 class="font-display font-bold text-slate-900 text-sm">Today's qualifying names</h3></div>
      <div class="text-[11px] text-slate-500 mb-2">${head}</div>
      <div class="rounded-lg bg-slate-50/60 ring-1 ring-slate-100 p-1 space-y-0.5">${rows}</div>
      <div class="text-[10px] text-slate-400 mt-2 leading-snug">Your parameter filters drive the <strong>live</strong> names here (today's values, AND-gate) and the backtest on every day that carries indicator history — logged from today forward. Earlier snapshots fall back to composite-only selection, so the indicator backtest deepens each day.</div>
    </div>`;
}

// Colour-coded, grouped strategy configurator. Three labelled steps —
// which stocks, when to sell, how often to refresh — each slider with a
// value chip + one-line helper. Reads like a sentence at the top.
const STRAT_ACCENT = {
  indigo:  { chip: "text-indigo-700 bg-indigo-50 ring-indigo-200",   range: "accent-indigo-600" },
  emerald: { chip: "text-emerald-700 bg-emerald-50 ring-emerald-200", range: "accent-emerald-600" },
  rose:    { chip: "text-rose-700 bg-rose-50 ring-rose-200",         range: "accent-rose-600" },
  amber:   { chip: "text-amber-700 bg-amber-50 ring-amber-200",       range: "accent-amber-600" },
  sky:     { chip: "text-sky-700 bg-sky-50 ring-sky-200",             range: "accent-sky-600" },
};
const STRAT_FIELD_META = {
  threshold:     { group: 0, help: "min score to consider", accent: "indigo" },
  basketSize:    { group: 0, help: "how many to hold",       accent: "indigo" },
  target:        { group: 1, help: "take profit at",         accent: "emerald" },
  sl:            { group: 1, help: "cut the loss at",        accent: "rose" },
  maxHoldDays:   { group: 1, help: "sell if held this long",  accent: "amber" },
  rebalanceDays: { group: 2, help: "rebuild the list every",  accent: "sky" },
};
const STRAT_GROUPS = [
  { title: "① Which stocks to buy", cols: "grid-cols-2" },
  { title: "② When to sell each one", cols: "grid-cols-1 sm:grid-cols-3" },
  { title: "③ How often to refresh", cols: "grid-cols-1" },
];
function renderStrategyConfig(strat) {
  const sliderHtml = (f) => {
    const m = STRAT_FIELD_META[f.key]; const a = STRAT_ACCENT[m.accent];
    return `
      <div>
        <div class="flex items-center justify-between gap-2 mb-1">
          <span class="text-xs font-semibold text-slate-700 truncate">${escapeHtml(f.label)}</span>
          <span class="text-sm font-bold tabular-nums ${a.chip} ring-1 rounded-md px-2 py-0.5 flex-shrink-0" data-strat-out="${f.key}">${strat[f.key]}${f.suffix}</span>
        </div>
        <input type="range" data-strat-field="${f.key}" min="${f.min}" max="${f.max}" step="${f.step}" value="${strat[f.key]}" class="w-full ${a.range} cursor-pointer" />
        <div class="text-[10px] text-slate-400 mt-0.5">${escapeHtml(m.help)}</div>
      </div>`;
  };
  const groups = STRAT_GROUPS.map((g, gi) => {
    const fs = STRAT_FIELDS.filter((f) => STRAT_FIELD_META[f.key].group === gi);
    return `
      <div class="rounded-xl ring-1 ring-slate-100 bg-slate-50/50 p-3">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-2">${escapeHtml(g.title)}</div>
        <div class="grid ${g.cols} gap-4">${fs.map(sliderHtml).join("")}</div>
      </div>`;
  }).join("");
  return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-center justify-between gap-2 flex-wrap mb-3">
        <div class="flex items-center gap-2 min-w-0 flex-1">
          <span class="text-lg">🎛️</span>
          <input type="text" data-strat-name value="${escapeHtml(strat.name)}" class="font-display font-bold text-slate-900 text-base bg-transparent ring-1 ring-transparent hover:ring-slate-200 focus:ring-2 focus:ring-indigo-300 rounded-lg px-2 py-1 outline-none min-w-0 flex-1" title="Click to rename" />
        </div>
        <div class="flex items-center gap-3 flex-shrink-0">
          <button type="button" data-strat-dup="${strat.id}" class="text-[11px] font-semibold text-slate-500 hover:text-indigo-600">Duplicate</button>
          <button type="button" data-strat-del="${strat.id}" class="text-[11px] font-semibold text-slate-500 hover:text-rose-600">Delete</button>
        </div>
      </div>
      <div class="text-xs text-slate-500 mb-3 leading-relaxed">Reads as: <em>"buy the top <strong>${strat.basketSize}</strong> stocks scoring ≥ <strong>${strat.threshold}</strong>; sell each when it's up <strong>${strat.target}%</strong>, down <strong>${strat.sl}%</strong>, or held <strong>${strat.maxHoldDays}</strong> days; refresh the list every <strong>${strat.rebalanceDays}</strong> days."</em> Drag any slider — the result updates instantly.</div>
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-3">${groups}</div>
      ${renderParamPicker(strat)}
    </div>`;
}

// Advanced settings — pick parameters from ANY pillar (dropdown to switch
// between Technicals / Fundamentals / Sentiment / Macro / Liquidity) AND
// tweak their values. A stock must pass every switched-on parameter.
let paramPickerPillar = "Technicals";   // which pillar the dropdown shows
let paramAdvancedOpen = false;          // persists the <details> state across re-renders
function renderParamPicker(strat) {
  const params = strat.params || {};
  const pillar = PARAM_PILLARS.includes(paramPickerPillar) ? paramPickerPillar : "Technicals";
  const num = (v) => (v == null ? "" : v);
  const fieldInputs = (p) => {
    const c = params[p.key] || {}; const d = p.def || {};
    const inp = (which, val) => `<input type="number" data-param-input="${p.key}" data-param-which="${which}" value="${num(val)}" step="${p.step || 1}" class="w-14 rounded ring-1 ring-slate-200 px-1.5 py-0.5 text-[11px] tabular-nums outline-none focus:ring-2 focus:ring-indigo-300" />`;
    if (p.kind === "range") return `${inp("min", c.min ?? d.min)}<span class="text-[10px] text-slate-400">–</span>${inp("max", c.max ?? d.max)}`;
    if (p.kind === "min" || p.kind === "max") return inp("val", c.val ?? d.val);
    return "";
  };
  const cats = [...new Set(STRAT_PARAMS.filter((p) => p.pillar === pillar).map((p) => p.cat))];
  const groups = cats.map((cat) => {
    const ps = STRAT_PARAMS.filter((p) => p.pillar === pillar && p.cat === cat);
    return `
      <div>
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1.5">${escapeHtml(cat)}</div>
        <div class="space-y-1.5">
          ${ps.map((p) => {
            const on = !!params[p.key]?.on;
            return `<div class="flex items-center gap-2 text-xs">
              <input type="checkbox" data-param-on="${p.key}" ${on ? "checked" : ""} class="accent-indigo-600 w-3.5 h-3.5 flex-shrink-0" />
              <span class="text-slate-700 flex-1 min-w-0 truncate ${on ? "" : "text-slate-400"}">${escapeHtml(p.label)}</span>
              <span class="flex items-center gap-1 flex-shrink-0 ${on ? "" : "opacity-40"}">${fieldInputs(p)}</span>
            </div>`;
          }).join("")}
        </div>
      </div>`;
  }).join("");
  const opts = PARAM_PILLARS.map((pl) => {
    const c = activeParamCountForPillar(params, pl);
    return `<option value="${pl}" ${pl === pillar ? "selected" : ""}>${pl}${c ? ` (${c} on)` : ""}</option>`;
  }).join("");
  const total = activeParamCount(params);
  return `
    <details data-param-advanced class="mt-3" ${paramAdvancedOpen || total ? "open" : ""}>
      <summary class="cursor-pointer text-[11px] font-semibold text-indigo-600 hover:text-indigo-700 select-none">Advanced — extra parameters (${total} on) ▾</summary>
      <div class="flex items-center gap-2 mt-3 flex-wrap">
        <span class="text-[11px] font-semibold text-slate-600">Parameter type:</span>
        <select data-param-pillar class="rounded-lg ring-1 ring-slate-200 px-2 py-1.5 text-sm font-semibold text-slate-800 bg-white outline-none focus:ring-2 focus:ring-indigo-300">${opts}</select>      </div>
      <div class="text-[11px] text-slate-500 mt-2">Tick a parameter to require it, and <strong>tweak its value</strong>. A stock must pass <strong>all</strong> the ones you switch on (across every pillar).</div>
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-5 gap-y-3 mt-3">${groups}</div>
      <div class="text-[10px] text-slate-400 mt-3">Nothing on = pick purely by composite score. <strong>Pillar-score</strong> params (Fundamentals/Sentiment/Macro/Liquidity ≥ %) backtest over full history; <strong>granular</strong> tech/fundamental params bind on days that carry the data (accruing forward).</div>
    </details>`;
}

const STRAT_PALETTE = ["#6366f1", "#0ea5e9", "#10b981", "#f59e0b", "#ec4899", "#8b5cf6", "#14b8a6", "#ef4444"];
function stratColors(views) { return views.map((v, i) => ({ ...v, color: STRAT_PALETTE[i % STRAT_PALETTE.length] })); }

// ------------------- Weight Lab: pillar-weight tuning + AI search -------------------
// Replicates the SPIP composite (a weighted blend of the 5 pillars) but lets
// the desk change the blend and see which mix would have picked the best
// held basket — or let AI grid-search ~1,000 mixes for the best one.
const LAB_PILLARS = [
  { key: "fundamentals", label: "Fundamental", dot: "bg-indigo-500" },
  { key: "technicals",   label: "Technical",   dot: "bg-sky-500" },
];

// Normalise a raw weight mix to NON-NEGATIVE integers summing to 100, via the
// largest-remainder (Hamilton) method — floor each share, then hand the leftover
// units to the largest fractional parts. Avoids the old "last pillar absorbs the
// remainder" trick, which could push a pillar negative (e.g. 0/0/45/75/0).
// Falls back to the framework default if the mix is all-zero.
function normalizeWeights(w) {
  const keys = LAB_PILLARS.map((p) => p.key);
  const raw = keys.map((k) => Math.max(0, Number(w[k]) || 0));
  const sum = raw.reduce((a, v) => a + v, 0);
  if (sum <= 0) return { ...composite.PILLAR_WEIGHTS };
  const exact = raw.map((v) => (v / sum) * 100);
  const out = {};
  keys.forEach((k, i) => { out[k] = Math.floor(exact[i]); });
  let rem = 100 - keys.reduce((a, k) => a + out[k], 0);
  const order = exact.map((v, i) => ({ i, frac: v - Math.floor(v) })).sort((a, b) => b.frac - a.frac);
  for (let j = 0; j < order.length && rem > 0; j++) { out[keys[order[j].i]]++; rem--; }
  return out;
}

// Precompute what the lab needs once: the anchor candidates (with their pillar
// pct vector) and a per-day ticker→close map for the tracking window. Reused
// across every weight mix so the AI search stays fast.
function buildLabContext(snapshots, anchorDate) {
  const anchorSnap = snapshots.find((s) => s.date >= anchorDate) || snapshots[0];
  const tracking = snapshots.filter((s) => s.date >= anchorSnap.date);
  const trackMaps = tracking.map((s) => {
    const m = new Map();
    for (const st of s.stocks) if (typeof st.close === "number") m.set(st.ticker, st.close);
    return m;
  });
  const cands = [];
  for (const s of anchorSnap.stocks) {
    if (!s.dataComplete || s.hardFailed || typeof s.close !== "number" || !s.ticker) continue;
    const p = s.pillars;
    const v = LAB_PILLARS.map((x) => p?.[x.key]?.pct);
    if (v.some((x) => x == null)) continue;
    cands.push({ ticker: s.ticker, name: s.name || s.ticker, close: s.close, v });
  }
  return { anchorSnap, tracking, trackMaps, cands };
}

// Held (passive) top-N basket under a weight mix: rank candidates by the
// weighted-average pillar score, hold, track equal-weight to today. Returns
// { finalReturn, maxDD, picks, curve }.
function heldPerfCore(ctx, weights, N = 7) {
  const wsum = LAB_PILLARS.reduce((a, x) => a + (weights[x.key] || 0), 0) || 1;
  const wv = LAB_PILLARS.map((x) => (weights[x.key] || 0) / wsum);
  const ranked = ctx.cands
    .map((c) => { let sc = 0; for (let i = 0; i < 5; i++) sc += c.v[i] * wv[i]; return { c, sc }; })
    .sort((a, b) => b.sc - a.sc)
    .slice(0, N);
  if (!ranked.length) return null;
  const entry = {}; const picks = [];
  for (const { c, sc } of ranked) { entry[c.ticker] = c.close; picks.push({ ticker: c.ticker, name: c.name, composite: sc }); }
  const last = { ...entry };
  // Per-stock return series over the hold window — so a basket can be
  // expanded to show each name's own max upside / drawdown (client ask:
  // "click a basket and see the seven stocks with max drawdown / upside").
  const stockCurves = {}; for (const tk of Object.keys(entry)) stockCurves[tk] = [];
  const curve = ctx.trackMaps.map((m, i) => {
    let sum = 0, n = 0;
    for (const tk of Object.keys(entry)) {
      const cl = m.has(tk) ? m.get(tk) : last[tk];
      if (typeof cl === "number") { last[tk] = cl; sum += cl / entry[tk]; n++; stockCurves[tk].push({ date: ctx.tracking[i].date, retPct: (cl / entry[tk] - 1) * 100 }); }
    }
    return { date: ctx.tracking[i].date, retPct: n ? (sum / n - 1) * 100 : null };
  });
  for (const p of picks) { const sc = stockCurves[p.ticker] || []; p.up = curveMaxUpside(sc); p.dd = curveMaxDrawdown(sc); p.ret = lastRet(sc); }
  return { finalReturn: lastRet(curve), maxDD: curveMaxDrawdown(curve), picks, curve };
}

function weightedBasketPerf(snapshots, anchorDate, weights, N = 7) {
  if (!snapshots?.length) return null;
  return heldPerfCore(buildLabContext(snapshots, anchorDate), weights, N);
}

// Grid-search pillar weights (step 10, summing to 100 → ~1,000 mixes) for the
// best risk-adjusted score (return + 0.4·drawdown, drawdown ≤ 0).
function findBestWeights(ctx, N = 7) {
  const step = 10; let best = null, tried = 0;
  for (let f = 0; f <= 100; f += step)
    for (let t = 0; f + t <= 100; t += step)
      for (let m = 0; f + t + m <= 100; m += step)
        for (let se = 0; f + t + m + se <= 100; se += step) {
          const l = 100 - f - t - m - se;
          const weights = { fundamentals: f, technicals: t, macro: m, sentiment: se, liquidity: l };
          const perf = heldPerfCore(ctx, weights, N);
          tried++;
          if (!perf) continue;
          const score = perf.finalReturn + 0.4 * perf.maxDD;
          if (!best || score > best.score) best = { weights, perf, score };
        }
  return best ? { ...best, tried } : null;
}

function labWeightChips(w) {
  return LAB_PILLARS.map((p) => `<span class="inline-flex items-center gap-1 text-[10px] font-semibold text-slate-500 tabular-nums"><span class="inline-block w-1.5 h-1.5 rounded-full ${p.dot}"></span>${p.label.slice(0, 4)} ${w[p.key] ?? 0}</span>`).join(" ");
}

// One result card: a weight mix + its held-basket return, drawdown and picks.
function renderLabResultCard(title, badge, weights, perf, highlight) {
  const keys = LAB_PILLARS.map((p) => p.key);
  const wsum = keys.reduce((a, k) => a + (Number(weights[k]) || 0), 0);
  // Show the weights EXACTLY as set — no silent rescaling — so the chips on
  // the card match the sliders (fixes "I set 50 but it shows 30"). Only a
  // genuinely empty mix falls back to the framework default.
  const nW = wsum > 0 ? weights : composite.PILLAR_WEIGHTS;
  const ret = perf?.finalReturn;
  const dd = perf?.maxDD;
  const retCls = ret == null ? "text-slate-400" : ret >= 0 ? "text-emerald-600" : "text-rose-600";
  const picks = (perf?.picks || []).slice(0, 7);
  const sign = (v) => v == null ? "—" : `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
  // Per-stock rows revealed when the analyst clicks the basket — each name's
  // own return, max upside and max drawdown over the hold window.
  const stockRows = picks.map((p) => `
    <div class="flex items-center gap-2 py-0.5 text-[10px] tabular-nums">
      <span class="font-semibold text-slate-700 truncate flex-1 min-w-0" title="${escapeHtml(p.name || p.ticker)}">${escapeHtml(p.ticker)}</span>
      <span class="${(p.ret ?? 0) >= 0 ? "text-emerald-600" : "text-rose-600"} w-12 text-right font-semibold">${sign(p.ret)}</span>
      <span class="text-emerald-600 w-14 text-right" title="Max upside since entry">▲ ${sign(p.up)}</span>
      <span class="text-rose-600 w-14 text-right" title="Max drawdown since entry">▼ ${sign(p.dd)}</span>
    </div>`).join("");
  return `
    <div class="rounded-xl p-3 ring-1 ${highlight ? "ring-emerald-300 bg-emerald-50/40" : "ring-slate-200 bg-white"}">
      <div class="flex items-center justify-between gap-2 mb-1">
        <div class="text-[11px] font-bold uppercase tracking-wider text-slate-600">${escapeHtml(title)}</div>
        ${badge || ""}
      </div>
      <div class="flex items-baseline gap-2">
        <div class="${retCls} text-2xl font-bold tabular-nums leading-none">${ret == null ? "—" : (ret >= 0 ? "+" : "") + ret.toFixed(2) + "%"}</div>
        <div class="text-[10px] tabular-nums ${dd < -0.005 ? "text-rose-600" : "text-slate-400"}"><span class="uppercase tracking-wider font-semibold text-slate-400">DD</span> ${dd == null ? "—" : dd.toFixed(2) + "%"}</div>
      </div>
      <div class="mt-1.5 flex flex-wrap gap-x-2 gap-y-0.5">${labWeightChips(nW)}</div>
      ${picks.length ? `
      <details class="mt-2 pt-2 border-t border-slate-100 group">
        <summary class="cursor-pointer list-none flex items-center gap-1 text-[10px] font-semibold text-indigo-600 hover:text-indigo-700 select-none">
          <span class="transition-transform group-open:rotate-90">▸</span> Basket · ${picks.length} stock${picks.length === 1 ? "" : "s"} <span class="text-slate-400 font-normal">(ret · max ▲ / ▼)</span>
        </summary>
        <div class="mt-1.5 space-y-0">${stockRows}</div>
      </details>` : `<div class="mt-2 pt-2 border-t border-slate-100 text-[10px] text-slate-400">Basket: —</div>`}
    </div>`;
}

function renderWeightLab() {
  const snaps = state.cache.history?.snapshots;
  if (!snaps?.length) return `<section id="weight-lab"></section>`;
  const anchor = snaps[0].date, today = snaps[snaps.length - 1].date;
  const ctx = buildLabContext(snaps, anchor);
  // Anchor on the CURRENT live weights (what the AI basket actually uses),
  // not the framework baseline — so the Lab opens on "current" (founder ask).
  const currentW = { ...(state.pillarWeights || composite.PILLAR_WEIGHTS) };
  const yourW = state.labWeights || currentW;
  const defW = composite.PILLAR_WEIGHTS;   // framework baseline — the "Reset to default" target
  const yourTotal = LAB_PILLARS.reduce((a, p) => a + (Number(yourW[p.key]) || 0), 0);
  const perfCurrent = heldPerfCore(ctx, currentW);
  const perfYour = heldPerfCore(ctx, yourW);
  const aiBest = state.labAiBest;
  const nDays = ctx.tracking.length;

  const sliders = LAB_PILLARS.map((p) => {
    const val = yourW[p.key] ?? 0;
    return `
      <div class="flex items-center gap-2">
        <span class="inline-block w-1.5 h-1.5 rounded-full ${p.dot} flex-shrink-0"></span>
        <span class="text-[11px] font-semibold text-slate-600 w-20 flex-shrink-0">${p.label}</span>
        <input type="range" min="0" max="100" step="1" data-lab-w="${p.key}" value="${val}" class="flex-1 accent-indigo-600">
        <span class="text-xs font-bold tabular-nums text-slate-800 w-8 text-right" data-lab-wval="${p.key}">${val}</span>
      </div>`;
  }).join("");

  const aiBadge = `<span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider bg-emerald-100 text-emerald-700 ring-1 ring-emerald-200">AI best</span>`;
  const defBadge = `<span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider bg-slate-100 text-slate-600 ring-1 ring-slate-200">current</span>`;
  const yourBadge = `<span class="inline-flex items-center px-1.5 py-0 rounded text-[9px] font-bold uppercase tracking-wider bg-indigo-100 text-indigo-700 ring-1 ring-indigo-200">yours</span>`;

  const aiCard = aiBest
    ? renderLabResultCard("AI best mix", aiBadge, aiBest.weights, aiBest.perf, true)
    : `<div class="rounded-xl p-3 ring-1 ring-dashed ring-slate-300 bg-slate-50/40 flex flex-col items-center justify-center text-center gap-2">         <button type="button" data-lab-ai class="px-3 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-semibold hover:bg-slate-800">🤖 Find best weights</button>
       </div>`;

  return `
    <section id="weight-lab" class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <div class="flex items-start justify-between gap-3 flex-wrap mb-1">
        <div class="flex items-center gap-2">
          <span class="text-base">⚖️</span>
          <h3 class="font-display font-bold text-slate-900 text-base">Weight Lab</h3>
        </div>
        <div class="text-[11px] text-slate-500">Held top-7 · backtested over ${nDays} days · ${fmtDateDMY(anchor)} → ${fmtDateDMY(today)}</div>
      </div>
      <div class="text-sm text-slate-600 mb-4 max-w-2xl">Change how much each pillar counts and watch which basket the mix would have picked — then let AI hunt for the best mix. This is the same blend the AI Basket uses; nothing changes on the live basket until you hit <strong>Apply to Basket</strong>.</div>
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <div>
          <div class="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-2">Your weights</div>
          <div class="space-y-2.5">${sliders}</div>
          <div class="flex items-center justify-between gap-2 mt-3">
            <span class="text-[11px] text-slate-400">Weights can't exceed 100 — the total is capped as you drag.</span>
            <span data-lab-total class="text-[11px] font-bold tabular-nums px-2 py-0.5 rounded-lg ring-1 ${yourTotal === 100 ? "text-emerald-700 bg-emerald-50 ring-emerald-200" : "text-amber-700 bg-amber-50 ring-amber-200"}">Total ${yourTotal} / 100</span>
          </div>
          <div class="flex items-center gap-2 mt-3 flex-wrap">
            <button type="button" data-lab-ai class="px-3 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-semibold hover:bg-slate-800">🤖 AI find best weights</button>
            <button type="button" data-lab-apply class="px-3 py-1.5 rounded-lg bg-indigo-600 text-white text-xs font-semibold hover:bg-indigo-700">Apply to AI Basket</button>
            <button type="button" data-lab-reset title="Reset to the framework default — F${defW.fundamentals} · T${defW.technicals} · M${defW.macro} · S${defW.sentiment} · L${defW.liquidity}" class="px-3 py-1.5 rounded-lg text-slate-600 text-xs font-semibold hover:bg-slate-100">↺ Reset to default</button>
          </div>
        </div>
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
          ${renderLabResultCard("Current Basket", defBadge, currentW, perfCurrent, false)}
          ${renderLabResultCard("Your weights", yourBadge, yourW, perfYour, false)}
          ${aiCard}
        </div>
      </div>
    </section>`;
}

function refreshWeightLab() {
  const el = document.querySelector("#custom-content #weight-lab");
  if (!el) return;
  el.outerHTML = renderWeightLab();
  wireWeightLab();
}

function wireWeightLab() {
  const root = "#custom-content";
  const sliders = () => Array.from($$(`${root} [data-lab-w]`));   // Array, not NodeList — need .filter/.reduce
  const readAll = () => { const w = {}; sliders().forEach((s) => { w[s.dataset.labW] = Math.max(0, Math.min(100, Number(s.value) || 0)); }); return w; };
  const updateTotal = () => {
    const total = sliders().reduce((a, s) => a + (Number(s.value) || 0), 0);
    const chip = $(`${root} [data-lab-total]`);
    if (!chip) return;
    chip.textContent = `Total ${total} / 100`;
    const ok = total === 100;
    chip.className = `text-[11px] font-bold tabular-nums px-2 py-0.5 rounded-lg ring-1 ${ok ? "text-emerald-700 bg-emerald-50 ring-emerald-200" : "text-amber-700 bg-amber-50 ring-amber-200"}`;
  };
  sliders().forEach((inp) => {
    inp.addEventListener("input", () => {
      // Cap this pillar so the five weights can never sum past 100 (client:
      // "don't let me go beyond 100"). Each pillar is set directly — no silent
      // rescaling — so the number shown always equals the number set.
      const key = inp.dataset.labW;
      const others = sliders().filter((s) => s.dataset.labW !== key).reduce((a, s) => a + (Number(s.value) || 0), 0);
      const maxAllowed = Math.max(0, 100 - others);
      let v = Math.max(0, Math.min(100, Number(inp.value) || 0));
      if (v > maxAllowed) { v = maxAllowed; inp.value = String(v); }
      const lbl = $(`${root} [data-lab-wval="${key}"]`);
      if (lbl) lbl.textContent = v;
      updateTotal();
    });
    inp.addEventListener("change", () => {
      const w = readAll();
      state.labWeights = w; saveLabWeights(w); refreshWeightLab();
    });
  });
  $$(`${root} [data-lab-ai]`).forEach((btn) => btn.addEventListener("click", () => {
    btn.disabled = true; btn.textContent = "🤖 Searching ~1,000 mixes…";
    setTimeout(() => {
      const snaps = state.cache.history?.snapshots;
      if (!snaps?.length) return;
      const ctx = buildLabContext(snaps, snaps[0].date);
      const best = findBestWeights(ctx);
      // AI's mix lands in its OWN card only — the analyst's sliders stay put
      // (client: "when I hit AI find best weights, MY weights also changed —
      // keep my original weights").
      if (best) state.labAiBest = best;
      refreshWeightLab();
    }, 30);
  }));
  $(`${root} [data-lab-apply]`)?.addEventListener("click", () => {
    const w = normalizeWeights(state.labWeights || composite.PILLAR_WEIGHTS);
    state.pillarWeights = w; savePillarWeights(w);
    state.labWeights = { ...w };   // the applied weights are now the current/live anchor
    delete state.cache.composite; delete state.cache.topPicks; state.compositeBySlug.clear();
    alert(`Applied to AI Basket: F${w.fundamentals} · T${w.technicals} · M${w.macro} · S${w.sentiment} · L${w.liquidity}. Open the AI Basket tab to see the re-scored basket.`);
    refreshWeightLab();   // re-render so the "Current SPIP" card reflects the new live weights
  });
  $(`${root} [data-lab-reset]`)?.addEventListener("click", () => {
    // Reset = the framework default (F40/T35/M15/S5/L5), not whatever the live
    // SPIP weights may have drifted to (client: "reset to default").
    state.labWeights = { ...composite.PILLAR_WEIGHTS };
    saveLabWeights(state.labWeights); state.labAiBest = null; refreshWeightLab();
  });
}

// ============ 1-Year Technical Back-test (real OHLC history) ============
// Replays a pure-technical rotation over ~1 year of real daily prices
// (data/history-technical.json, built by scrape-history-ohlc.mjs): rank the
// universe by the technical score each day, hold the top N, exit on target /
// stop-loss / drop-out, rebalance on a cadence. This is Bharat's "test the
// technicals over a year before running live" — separate from the live product.
let techHist = null;
async function loadTechHistory() {
  if (techHist !== null) return techHist;
  try { techHist = await fetch("data/history-technical.json").then((r) => (r.ok ? r.json() : false)); }
  catch { techHist = false; }
  return techHist;
}

// Pre-rank each day once (score desc) so grid search stays fast.
function buildTechContext(hist, windowDays) {
  const start = Math.max(0, hist.dates.length - windowDays);
  const dates = hist.dates.slice(start);
  const T = Object.keys(hist.tickers);
  const ranked = dates.map((_, di) => {
    const i = start + di, row = [];
    for (const t of T) { const sc = hist.tickers[t].score[i], cl = hist.tickers[t].close[i]; if (sc != null && cl != null) row.push({ t, score: sc, close: cl }); }
    row.sort((a, b) => b.score - a.score);
    return row;
  });
  const closeAt = (t, di) => hist.tickers[t].close[start + di];
  return { start, dates, T, ranked, closeAt, hist };
}

// chg = a charger from charges.js. Applied on every buy AND sell so heavy
// churn is penalised — pass ZERO_CHARGER for the gross (frictionless) run.
// The book is normalised to 1.0, so notionals are scaled to the real
// capital before charging: the DP fee is flat, and a fee computed on a
// notional of 0.14 rather than ₹7,143 would round to nothing.
function runTechBacktest(ctx, p, chg = ZERO_CHARGER, capital = 50000) {
  const { dates, ranked, closeAt, hist } = ctx;
  const N = Math.max(1, p.basketSize), thr = p.threshold, tgt = p.targetPct / 100, sl = p.slPct / 100, reb = Math.max(1, p.rebalanceDays);
  let cash = 1, lastReb = 0, trades = 0, charges = 0;
  const holds = new Map();
  const curve = [];
  for (let di = 0; di < dates.length; di++) {
    const isReb = di === 0 || (di - lastReb) >= reb;
    if (isReb) lastReb = di;
    const qual = ranked[di].filter((r) => r.score >= thr);
    const topN = new Set(qual.slice(0, N).map((r) => r.t));
    for (const [t, pos] of [...holds]) {
      const px = closeAt(t, di); if (px == null) continue;
      const g = px / pos.entryPrice - 1;
      if (g >= tgt || g <= -sl || (isReb && !topN.has(t))) {
        const proceeds = pos.units * px, fee = chg.sell(proceeds * capital) / capital;
        cash += proceeds - fee; charges += fee; holds.delete(t); trades++;
      }
    }
    if (holds.size < N) {
      for (const r of qual) {
        if (holds.size >= N) break;
        if (holds.has(r.t)) continue;
        let nav = cash; for (const [tk, pos] of holds) nav += pos.units * (closeAt(tk, di) ?? pos.entryPrice);
        const buy = Math.min(nav / N, cash / (1 + buyRate(chg.prefs || {})));   // leave room for the buy-side fee
        if (buy < 1e-6) break;
        const fee = chg.buy(buy * capital) / capital;
        holds.set(r.t, { units: buy / r.close, entryPrice: r.close }); cash -= buy + fee; charges += fee; trades++;
      }
    }
    let val = cash; for (const [t, pos] of holds) val += pos.units * (closeAt(t, di) ?? pos.entryPrice);
    curve.push({ date: dates[di], retPct: (val - 1) * 100 });
  }
  const finalReturn = curve.length ? curve[curve.length - 1].retPct : 0;
  let peak = -Infinity, maxDD = 0;
  for (const pt of curve) { const f = 1 + pt.retPct / 100; if (f > peak) peak = f; const dd = (f / peak - 1) * 100; if (dd < maxDD) maxDD = dd; }
  const last = ranked[ranked.length - 1] || [];
  const picksNow = last.filter((r) => r.score >= thr).slice(0, N)
    .map((r) => ({ ticker: r.t, score: r.score, sector: hist.tickers[r.t].sector }));
  return { finalReturn, maxDD, curve, trades, picksNow, days: curve.length, chargesPct: charges * 100 };
}

// Run one window at the current params — gross (no cost) and net (with cost)
// — and derive the founder's columns: gross / cost / net / annualized (XIRR)
// / max drawdown. Annualized = CAGR of the NET return over the window's years
// (= XIRR for a single-capital backtest). 252 trading days ≈ 1 year.
function techWindowRow(p, windowDays, label) {
  const ctx = buildTechContext(techHist, windowDays);
  const cap = simPrefs.capital ?? 50000;
  const gross = runTechBacktest(ctx, p, ZERO_CHARGER, cap);
  const net = runTechBacktest(ctx, p, makeCharger(simPrefs), cap);
  // Annualise over the ACTUAL trading days simulated, not the nominal window —
  // buildTechContext clamps `start` to the available history, so net.days can
  // be < windowDays (e.g. a shorter regenerated file). Using windowDays there
  // would divide a ~1-year 2Y row by 2 years and understate the annual figure.
  const years = Math.max(net.days / 252, 1e-6);
  const annualized = (Math.pow(1 + net.finalReturn / 100, 1 / years) - 1) * 100;
  return {
    label, windowDays,
    gross: gross.finalReturn, net: net.finalReturn,
    cost: gross.finalReturn - net.finalReturn,
    annualized, maxDD: net.maxDD, trades: net.trades, days: net.days,
    curve: net.curve, picksNow: net.picksNow,
    start: ctx.dates[0], end: ctx.dates[ctx.dates.length - 1],
  };
}

// Grid-search the strategy params for the best risk-adjusted score.
function findBestTechStrategy(ctx) {
  const GRID = { basketSize: [5, 7, 10], rebalanceDays: [5, 10, 20], targetPct: [8, 12, 20], slPct: [5, 8, 12], threshold: [50, 55, 60] };
  let best = null, tried = 0;
  const chg = makeCharger(simPrefs);   // score on NET return so the optimiser can't win by churning
  for (const basketSize of GRID.basketSize)
    for (const rebalanceDays of GRID.rebalanceDays)
      for (const targetPct of GRID.targetPct)
        for (const slPct of GRID.slPct)
          for (const threshold of GRID.threshold) {
            const params = { basketSize, rebalanceDays, targetPct, slPct, threshold };
            const r = runTechBacktest(ctx, params, chg, simPrefs.capital ?? 50000); tried++;
            const score = r.finalReturn + 0.4 * r.maxDD;
            if (!best || score > best.score) best = { params, r, score };
          }
  return best ? { ...best, tried } : null;
}

const TECH_FIELDS = [
  { key: "basketSize",    label: "Basket size",   min: 3,  max: 15, step: 1, suffix: "" },
  { key: "rebalanceDays", label: "Rebalance",     min: 1,  max: 30, step: 1, suffix: "d" },
  { key: "targetPct",     label: "Target",        min: 3,  max: 40, step: 1, suffix: "%" },
  { key: "slPct",         label: "Stop-loss",     min: 2,  max: 25, step: 1, suffix: "%" },
  { key: "threshold",     label: "Min strength",  min: 0,  max: 90, step: 5, suffix: "/100" },
];

function renderTechBacktest() {
  if (techHist === false || !techHist) {
    return `<section id="tech-backtest" class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5">
      <div class="flex items-center gap-2 mb-1"><span class="text-base">🧪</span><h3 class="font-display font-bold text-slate-900 text-base">1-Year Technical Back-test</h3></div>
      <div class="text-sm text-slate-500">History not loaded yet — run <code class="bg-slate-100 px-1 rounded text-[11px]">scrape-history-ohlc.mjs</code> to build <code class="bg-slate-100 px-1 rounded text-[11px]">history-technical.json</code>.</div>
    </section>`;
  }
  const p = state.techParams;
  const cs = chargeSummary(simPrefs.capital ?? 50000, state.techParams.basketSize || 7, simPrefs);
  // Compute all three windows at once so the client sees 6M / 1Y / 2Y side by
  // side (founder ask) instead of one number behind a window toggle.
  const rows = [techWindowRow(p, 126, "6M"), techWindowRow(p, 252, "1Y"), techWindowRow(p, 504, "2Y")];
  const sel = rows.find((x) => x.windowDays === p.windowDays) || rows[rows.length - 1];
  const full = rows[rows.length - 1];

  const controls = TECH_FIELDS.map((f) => `
    <label class="flex items-center gap-2">
      <span class="text-[11px] font-semibold text-slate-500 w-[70px] flex-shrink-0">${f.label}</span>
      <input type="range" min="${f.min}" max="${f.max}" step="${f.step}" data-tech-p="${f.key}" value="${p[f.key]}" class="flex-1 accent-indigo-600">
      <span class="text-xs font-bold tabular-nums text-slate-800 w-14 text-right" data-tech-pval="${f.key}">${p[f.key]}${f.suffix}</span>
    </label>`).join("");

  const sign = (v) => `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
  const cls = (v) => v >= 0 ? "text-emerald-600" : "text-rose-600";
  const tableRows = rows.map((x) => {
    const on = x.windowDays === p.windowDays;
    return `
      <tr data-tech-window="${x.windowDays}" class="border-t border-slate-100 cursor-pointer transition ${on ? "bg-indigo-50/60" : "hover:bg-slate-50"}" title="Show this window's curve">
        <td class="py-2.5 pl-1 text-left font-bold text-slate-800">${x.label}${on ? ` <span class="text-[9px] font-semibold text-indigo-500 align-middle">● curve</span>` : ""}</td>
        <td class="py-2.5 px-2 text-right ${cls(x.gross)}">${sign(x.gross)}</td>
        <td class="py-2.5 px-2 text-right text-slate-400">−${x.cost.toFixed(1)}pp</td>
        <td class="py-2.5 px-2 text-right font-extrabold ${cls(x.net)}">${sign(x.net)}</td>
        <td class="py-2.5 px-2 text-right ${cls(x.annualized)}">${sign(x.annualized)}<span class="text-[9px] text-slate-400">/yr</span></td>
        <td class="py-2.5 px-2 text-right text-rose-600">${x.maxDD.toFixed(1)}%</td>
        <td class="py-2.5 pl-2 pr-1 text-right text-slate-400 tabular-nums">${x.trades}</td>
      </tr>`;
  }).join("");

  const picks = sel.picksNow.map((x) => `
    <button type="button" data-cohort-row data-cohort-side="ai" data-ticker="${escapeHtml(x.ticker)}" title="Open ${escapeHtml(x.ticker)} chart" class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white ring-1 ring-slate-200 hover:ring-indigo-300 hover:bg-indigo-50 hover:-translate-y-px transition text-xs">
      <span class="font-semibold text-slate-800">${escapeHtml(x.ticker)}</span>
      <span class="text-[10px] text-slate-400 tabular-nums">${x.score}</span>
    </button>`).join("");

  return `
    <section id="tech-backtest" class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 overflow-hidden">
      <div class="px-4 sm:px-5 pt-4 sm:pt-5">
        <div class="flex items-center gap-2"><span class="text-base">🧪</span><h3 class="font-display font-bold text-slate-900 text-base">Technical Back-test</h3></div>
        <div class="text-[12px] text-slate-600 mt-1.5 leading-snug">Buy the top <b>${p.basketSize}</b> stocks with technical strength <b>≥ ${p.threshold}/100</b> · sell each at <b>+${p.targetPct}%</b> or <b>−${p.slPct}%</b> · rebuild every <b>${p.rebalanceDays} days</b>.</div>
        <div class="text-[11px] text-slate-400 mt-0.5 tabular-nums">Real daily prices · ${fmtDateDMY(full.start)} → ${fmtDateDMY(full.end)} · <b>Net</b> is after real charges — <b>${cs.roundTripPct.toFixed(2)}%</b> a round trip at ₹${Math.round(cs.perPosition).toLocaleString("en-IN")} a position.</div>
      </div>

      <div class="px-4 sm:px-5 mt-3">
        <div class="overflow-x-auto rounded-xl ring-1 ring-slate-100">
          <table class="w-full text-sm tabular-nums">
            <thead>
              <tr class="text-[10px] font-bold uppercase tracking-wider text-slate-400 bg-slate-50/70">
                <th class="text-left py-2 pl-1">Window</th>
                <th class="text-right py-2 px-2" title="Return before transaction costs">Gross</th>
                <th class="text-right py-2 px-2" title="Charges drag = gross − net">Cost</th>
                <th class="text-right py-2 px-2" title="What you actually keep after costs">Net</th>
                <th class="text-right py-2 px-2" title="Annualised (CAGR / XIRR) of the net return">Annual</th>
                <th class="text-right py-2 px-2" title="Worst peak-to-trough decline">Max DD</th>
                <th class="text-right py-2 pl-2 pr-1">Trades</th>
              </tr>
            </thead>
            <tbody>${tableRows}</tbody>
          </table>
        </div>
        <div class="text-[10px] text-slate-400 mt-1.5">Gross before costs · Cost = charges drag · <b>Net</b> = what you keep · Annual = per-year (XIRR). Tap a row to plot its curve.</div>
      </div>

      ${renderTechCurve(sel.curve)}

      <div class="px-4 sm:px-5 pb-4 sm:pb-5">
        <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">Adjust the strategy</div>
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-2.5">${controls}</div>
        <div class="flex items-center gap-2 mt-4 flex-wrap">
          <button type="button" data-tech-ai class="px-3.5 py-2 rounded-lg bg-slate-900 text-white text-xs font-semibold hover:bg-slate-800">🤖 Find best settings</button>
          <button type="button" data-tech-reset class="px-3 py-2 rounded-lg text-slate-500 text-xs font-semibold hover:bg-slate-100">↺ Reset to default</button>        </div>
        <div class="mt-4 pt-3 border-t border-slate-100">
          <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">${sel.label} top picks now</div>
          <div class="flex flex-wrap gap-1.5">${picks || '<span class="text-xs text-slate-400">none above threshold</span>'}</div>
        </div>
      </div>
    </section>`;
}

// Big, full-width equity curve — the hero visual. Area fill under the line,
// dashed zero baseline, colour by up/down. SVG stretches to the card width.
// Hover crosshair + tooltip are wired in wireTechBacktest off techCurvePts.
let techCurvePts = [];
function renderTechCurve(curve) {
  const pts = (curve || []).filter((p) => p.retPct != null);
  techCurvePts = pts;
  if (pts.length < 2) return `<div class="px-5 py-10 text-center text-[11px] text-slate-400">Not enough data to plot.</div>`;
  const W = 1000, H = 210, M = { l: 4, r: 4, t: 16, b: 16 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const vals = pts.map((p) => p.retPct);
  let lo = Math.min(0, ...vals), hi = Math.max(0, ...vals); if (hi === lo) hi = lo + 1;
  const x = (i) => M.l + (i / (pts.length - 1)) * iw;
  const y = (v) => M.t + (1 - (v - lo) / (hi - lo)) * ih;
  const line = pts.map((p, i) => `${i ? "L" : "M"}${x(i).toFixed(1)},${y(p.retPct).toFixed(1)}`).join(" ");
  const area = `${line} L${x(pts.length - 1).toFixed(1)},${y(lo).toFixed(1)} L${x(0).toFixed(1)},${y(lo).toFixed(1)} Z`;
  const zeroY = y(0).toFixed(1);
  const up = pts[pts.length - 1].retPct >= 0;
  const col = up ? "#059669" : "#e11d48";
  return `<div id="tech-chart" class="relative select-none" style="cursor:crosshair">
    <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" style="width:100%;height:210px;display:block">
      <defs><linearGradient id="tech-area" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="${col}" stop-opacity="0.22"/><stop offset="100%" stop-color="${col}" stop-opacity="0.02"/>
      </linearGradient></defs>
      <line x1="${M.l}" y1="${zeroY}" x2="${W - M.r}" y2="${zeroY}" stroke="#cbd5e1" stroke-width="1" stroke-dasharray="4 5"/>
      <path d="${area}" fill="url(#tech-area)"/>
      <path d="${line}" fill="none" stroke="${col}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" vector-effect="non-scaling-stroke"/>
    </svg>
    <div data-tech-cross class="absolute top-0 bottom-0 w-px bg-slate-400/70 pointer-events-none hidden"></div>
    <div data-tech-dot class="absolute w-2 h-2 rounded-full pointer-events-none hidden" style="background:${col};transform:translate(-50%,-50%)"></div>
    <div data-tech-tip class="absolute pointer-events-none hidden px-2 py-1 rounded-md bg-slate-900 text-white text-[11px] font-semibold tabular-nums shadow-lg z-10 whitespace-nowrap" style="top:6px"></div>
  </div>`;
}

function wireTechBacktest() {
  const root = "#custom-content";
  $$(`${root} [data-tech-p]`).forEach((inp) => {
    inp.addEventListener("input", () => { const s = $(`${root} [data-tech-pval="${inp.dataset.techP}"]`); const f = TECH_FIELDS.find((x) => x.key === inp.dataset.techP); if (s) s.textContent = inp.value + (f ? f.suffix : ""); });
    inp.addEventListener("change", () => { state.techParams = { ...state.techParams, [inp.dataset.techP]: Number(inp.value) }; saveTechParams(state.techParams); refreshTechBacktest(); });
  });
  $$(`${root} [data-tech-window]`).forEach((b) => b.addEventListener("click", () => { state.techParams = { ...state.techParams, windowDays: Number(b.dataset.techWindow) }; saveTechParams(state.techParams); refreshTechBacktest(); }));
  const ai = $(`${root} [data-tech-ai]`);
  if (ai) ai.addEventListener("click", () => {
    ai.disabled = true; ai.textContent = "🤖 Searching…";
    setTimeout(() => {
      const ctx = buildTechContext(techHist, state.techParams.windowDays);
      const best = findBestTechStrategy(ctx);
      state.techAiBest = best;
      if (best) { state.techParams = { ...state.techParams, ...best.params }; saveTechParams(state.techParams); }
      refreshTechBacktest();
    }, 30);
  });
  $(`${root} [data-tech-reset]`)?.addEventListener("click", () => { state.techParams = { ...TECH_DEFAULTS }; saveTechParams(state.techParams); state.techAiBest = null; refreshTechBacktest(); });
  // Clickable picks — open the stock's drill chart (re-wired after each refresh).
  $$(`${root} #tech-backtest [data-cohort-row]`).forEach((el) => el.addEventListener("click", () => {
    const t = el.dataset.ticker; if (!t) return;
    const pick = buildCohortClickPick(t, "ai", null);
    if (pick) openHistoryDrill(pick);
  }));
  // Chart hover — crosshair + dot + date/return tooltip.
  const chart = document.querySelector(`${root} #tech-chart`);
  if (chart && techCurvePts.length > 1) {
    const cross = chart.querySelector("[data-tech-cross]"), dot = chart.querySelector("[data-tech-dot]"), tip = chart.querySelector("[data-tech-tip]");
    const W = 1000, H = 210, M = { l: 4, r: 4, t: 16, b: 16 }, iw = W - 8, ih = H - 32, n = techCurvePts.length;
    const vals = techCurvePts.map((p) => p.retPct);
    let lo = Math.min(0, ...vals), hi = Math.max(0, ...vals); if (hi === lo) hi = lo + 1;
    chart.addEventListener("mousemove", (e) => {
      const rect = chart.getBoundingClientRect();
      const idx = Math.max(0, Math.min(n - 1, Math.round((e.clientX - rect.left) / rect.width * (n - 1))));
      const pt = techCurvePts[idx]; if (!pt) return;
      const px = ((M.l + (idx / (n - 1)) * iw) / W) * rect.width;
      const py = ((M.t + (1 - (pt.retPct - lo) / (hi - lo)) * ih) / H) * rect.height;
      cross.style.left = `${px}px`; cross.classList.remove("hidden");
      dot.style.left = `${px}px`; dot.style.top = `${py}px`; dot.classList.remove("hidden");
      tip.textContent = `${fmtDateDMY(pt.date)} · ${pt.retPct >= 0 ? "+" : ""}${pt.retPct.toFixed(1)}%`;
      tip.style.left = `${Math.max(4, Math.min(rect.width - 120, px - 45))}px`; tip.classList.remove("hidden");
    });
    chart.addEventListener("mouseleave", () => { cross.classList.add("hidden"); dot.classList.add("hidden"); tip.classList.add("hidden"); });
  }
}

function refreshTechBacktest() {
  const el = document.querySelector("#custom-content #tech-backtest");
  if (!el) return;
  el.outerHTML = renderTechBacktest();
  wireTechBacktest();
}

function renderCustomOverview(views) {
  // Focus mode: the Custom tab is just the Weight Lab (the rotation strategy
  // cards + comparison chart are hidden). Flip SHOW_ROTATION_STRATEGIES to
  // bring the full Strategy Lab back.
  if (!SHOW_ROTATION_STRATEGIES) {
    return `<div class="space-y-5">${renderWeightLab()}${renderTechBacktest()}</div>`;
  }
  const withColor = stratColors(views);
  const series = withColor.filter((v) => v.view).map((v) => ({ label: v.strat.name, color: v.color, curve: v.view.equityCurve }));
  const firstNifty = withColor.find((v) => v.view?.niftyCurve?.length);
  if (firstNifty) series.push({ label: "Smallcap 250", color: "#94a3b8", curve: firstNifty.view.niftyCurve, dash: "5 4" });
  const aiViews = withColor.filter((v) => v.strat.origin === "ai");
  const userViews = withColor.filter((v) => v.strat.origin !== "ai");
  const grid = (arr) => `<div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">${arr.map((v) => renderStrategyCard(v.strat, v.view, v.color)).join("")}</div>`;
  return `
    <div class="space-y-5">
      <div class="bg-gradient-to-br from-indigo-50 via-white to-purple-50 rounded-2xl ring-1 ring-indigo-100 p-5">
        <div class="flex items-start justify-between gap-3 flex-wrap">
          <div class="flex items-center gap-2"><h2 class="font-display font-bold text-xl text-slate-900">Custom Strategy Lab</h2><span class="text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-700 ring-1 ring-indigo-200">Lab</span></div>
          <button type="button" data-perf-open class="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-white text-indigo-700 ring-1 ring-indigo-200 text-sm font-semibold shadow-sm hover:bg-indigo-50 whitespace-nowrap">📊 Performance</button>
        </div>
        <div class="text-sm text-slate-600 mt-1 max-w-2xl"><strong>A "what if I'd traded like this?" tester.</strong> Set buy &amp; sell rules and see how much money each plan <em>would have</em> made on past data — practice, not live trading.</div>
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-2 mt-3">
          ${howStep("1", "Open a strategy", "Click a card to set its rules with sliders.")}
          ${howStep("2", "It tests itself", "Every change instantly replays history and redraws the line.")}
          ${howStep("3", "Keep the winner", "The chart stacks all plans — keep the best, delete the rest.")}
        </div>
      </div>
      ${renderWeightLab()}
      ${renderTechBacktest()}
      ${renderMultiCurveChart(series, "Which plan performed best?", "Each line is one plan's growth over time (after charges). Higher = better.")}
      <section>
        <div class="flex items-center gap-2 mb-1"><span class="text-base">🤖</span><h3 class="font-display font-bold text-slate-900 text-base">AI Generated top strategies</h3></div>        ${grid(aiViews)}
      </section>
      <section>
        <div class="flex items-center justify-between gap-2 mb-3 flex-wrap">
          <div class="flex items-center gap-2"><span class="text-base">👤</span><h3 class="font-display font-bold text-slate-900 text-base">My strategies</h3></div>
          <button type="button" data-strat-new class="px-4 py-2 rounded-xl bg-indigo-600 text-white text-sm font-semibold shadow-sm hover:bg-indigo-700 whitespace-nowrap">+ New strategy</button>
        </div>
        ${userViews.length ? grid(userViews) : `<div class="bg-white rounded-2xl ring-1 ring-dashed ring-slate-300 p-6 text-center text-sm text-slate-500">No strategies of your own yet. Click <strong>+ New strategy</strong>, or open an AI one above and hit <strong>Duplicate</strong>.</div>`}
      </section>
      <details class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
        <summary class="cursor-pointer text-sm font-semibold text-slate-700 select-none">⚙ Shared capital &amp; charges <span class="font-normal text-slate-400">(applies to every plan)</span></summary>
        <div class="mt-3">${renderSimPanel(withColor.find((v) => v.view)?.view || {})}</div>
      </details>
    </div>`;
}

// Performance comparison — its own page inside the Custom Lab (reached via
// the "📊 Performance" button at the top). One row per strategy (AI presets
// + the user's own), rebuilt from `views` on every render so a strategy
// added/saved shows up immediately. Click a row to open its deep-dive.
function renderCustomPerformancePage(views) {
  const withColor = stratColors(views);
  return `
    <div class="space-y-4">
      <button type="button" data-perf-close class="text-sm font-semibold text-indigo-600 hover:text-indigo-700">← All strategies</button>
      <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
        <div class="flex items-center gap-2 mb-1"><span class="text-base">📊</span><h2 class="font-display font-bold text-slate-900 text-lg">Performance comparison</h2><span class="text-[10px] text-slate-400">${withColor.length} ${withColor.length === 1 ? "strategy" : "strategies"}</span></div>        ${renderCustomPerfTable(withColor)}
      </div>
    </div>`;
}

function renderCustomPerfTable(views) {
  const rows = views.map((v) => {
    const c = v.view?.equityCurve;
    return {
      s: v.strat, color: v.color,
      inception: v.view?.finalReturn ?? null,
      d1: c ? oneDayReturnPct(c) : null,
      w1: c ? trailingReturnPct(c, 7) : null,
      m1: c ? trailingReturnPct(c, 30) : null,
      dd: c ? curveMaxDrawdown(c) : null,
      alpha: v.view?.alpha ?? null,
      hit: v.view?.hitSummary?.hitRate ?? null,
      trades: v.view?.tradeCount ?? null,
    };
  });
  rows.sort((a, b) => (b.inception ?? -Infinity) - (a.inception ?? -Infinity));
  const startDate = views.find((v) => v.view)?.view?.startDate;
  const pc = (v, digits = 2) => v == null
    ? `<span class="text-slate-300">—</span>`
    : `<span class="tabular-nums font-semibold ${v >= 0 ? "text-emerald-700" : "text-rose-700"}">${v >= 0 ? "+" : ""}${v.toFixed(digits)}%</span>`;
  const body = rows.map((r) => `
    <tr class="border-t border-slate-100 hover:bg-slate-50 cursor-pointer" data-perf-row="${r.s.id}">
      <td class="py-2 pr-2 min-w-0">
        <div class="flex items-center gap-2 min-w-0">
          <span class="w-2 h-2 rounded-full flex-shrink-0" style="background:${r.color}"></span>
          <span class="font-semibold text-slate-800 text-xs truncate" title="${escapeHtml(r.s.name)}">${escapeHtml(r.s.name)}</span>
          <span class="text-[8px] font-bold uppercase tracking-wider px-1 py-0.5 rounded flex-shrink-0 ${r.s.origin === "ai" ? "bg-indigo-50 text-indigo-600 ring-1 ring-indigo-100" : "bg-slate-100 text-slate-500"}">${r.s.origin === "ai" ? "AI" : "Mine"}</span>
        </div>
      </td>
      <td class="py-2 px-2 text-right text-sm">${pc(r.inception)}</td>
      <td class="py-2 px-2 text-right text-xs">${pc(r.d1)}</td>
      <td class="py-2 px-2 text-right text-xs">${pc(r.w1)}</td>
      <td class="py-2 px-2 text-right text-xs">${pc(r.m1)}</td>
      <td class="py-2 px-2 text-right text-xs"><span class="tabular-nums ${r.dd == null ? "text-slate-300" : "text-rose-600 font-semibold"}">${r.dd == null ? "—" : r.dd.toFixed(1) + "%"}</span></td>
      <td class="py-2 px-2 text-right text-xs">${pc(r.alpha)}</td>
      <td class="py-2 px-2 text-right text-xs"><span class="tabular-nums ${r.hit == null ? "text-slate-300" : "text-slate-600 font-semibold"}">${r.hit == null ? "—" : r.hit.toFixed(0) + "%"}</span></td>
      <td class="py-2 pl-2 text-right text-xs tabular-nums text-slate-500">${r.trades == null ? "—" : r.trades}</td>
    </tr>`).join("");
  return `
    <div class="overflow-x-auto mt-3">
      <table class="w-full text-sm">
        <thead>
          <tr class="text-[10px] font-bold uppercase tracking-wider text-slate-500">
            <th class="text-left pb-1 pr-2">Strategy</th>
            <th class="text-right pb-1 px-2">Since inception</th>
            <th class="text-right pb-1 px-2">1D</th>
            <th class="text-right pb-1 px-2">1W</th>
            <th class="text-right pb-1 px-2">1M</th>
            <th class="text-right pb-1 px-2">Max DD</th>
            <th class="text-right pb-1 px-2">vs Nifty</th>
            <th class="text-right pb-1 px-2">Hit rate</th>
            <th class="text-right pb-1 pl-2">Trades</th>
          </tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
      <div class="text-[10px] text-slate-400 mt-2 leading-snug">All returns are net of charges${startDate ? `, backtested from ${fmtDateDMY(startDate)}` : ""}. <strong>Since inception</strong> = total return · <strong>1D/1W/1M</strong> = trailing windows (— when history is shorter than the window) · <strong>vs benchmark</strong> = return above the Nifty Smallcap 250 over the same period · <strong>Hit rate</strong> = share of all picks that hit target. Sorted by since-inception return; click a row to open it.</div>
    </div>`;
}

function howStep(n, title, body) {
  return `<div class="bg-white/70 rounded-xl ring-1 ring-slate-100 px-3 py-2"><div class="text-xs font-bold text-indigo-600">${n}. ${escapeHtml(title)}</div><div class="text-[11px] text-slate-500 mt-0.5">${escapeHtml(body)}</div></div>`;
}

function renderCustomDeepDive(entry, todaysNames) {
  const { strat, view } = entry;
  const todaysBlock = todaysNames ? renderTodaysNames(todaysNames, strat) : "";
  if (!view) {
    return `
      <div class="space-y-4">
        <button type="button" data-strat-back class="text-sm font-semibold text-indigo-600 hover:text-indigo-700">← All strategies</button>
        ${renderStrategyConfig(strat)}
        ${todaysBlock}
        <div class="bg-white rounded-2xl ring-1 ring-slate-200 p-6 text-sm text-slate-500">No qualifying picks for this strategy over the available history. Loosen the composite threshold or widen the bands.</div>
      </div>`;
  }
  const series = [
    { label: strat.name, color: "#6366f1", curve: view.equityCurve, width: 2.4 },
    { label: "Manual", color: "#f59e0b", curve: view.manualCurve || [] },
    { label: "Smallcap 250", color: "#94a3b8", curve: view.niftyCurve || [], dash: "5 4" },
  ];
  return `
    <div class="space-y-4">
      <div class="flex items-center justify-between gap-2 flex-wrap">
        <button type="button" data-strat-back class="text-sm font-semibold text-indigo-600 hover:text-indigo-700">← All strategies</button>
        <div class="text-[11px] text-slate-500">${view.tradeCount} trades · ${view.liveHoldings} held now</div>
      </div>
      ${renderPlainResult(view)}
      ${renderStrategyConfig(strat)}
      ${todaysBlock}
      ${renderMultiCurveChart(series, `${strat.name} vs Manual vs Nifty`, "Blue = your plan. Compared against the manual basket and Nifty 50.")}
      <details class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
        <summary class="cursor-pointer text-sm font-semibold text-slate-700 select-none">📊 Show detailed breakdown <span class="font-normal text-slate-400">(money in/out, risk, every trade, sector timing)</span></summary>
        <div class="space-y-4 mt-3">
          ${renderSimPanel(view)}
          ${renderStrategyKpis(view)}
          ${renderActivePickRowsSplit(view)}
          ${renderSectorTiming(view)}
        </div>
      </details>
    </div>`;
}

// Plain-English headline for a strategy — the one thing most people want
// to know: "would this plan have made money?"
function renderPlainResult(view) {
  const money = (n) => `₹${Math.round(n).toLocaleString("en-IN")}`;
  const net = view.finalReturn;
  const up = curveMaxUpside(view.equityCurve) ?? 0;
  const dd = curveMaxDrawdown(view.equityCurve) ?? 0;
  const cls = net >= 0 ? "text-emerald-700" : "text-rose-700";
  const bg = net >= 0 ? "from-emerald-50 to-teal-50 ring-emerald-200" : "from-rose-50 to-pink-50 ring-rose-200";
  return `
    <div class="rounded-2xl ring-1 bg-gradient-to-br ${bg} p-4 sm:p-5">
      <div class="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1">In plain words</div>
      <div class="text-sm sm:text-base text-slate-800 leading-relaxed">
        Run this plan with <strong>${money(view.startCapital)}</strong> and you'd have <strong class="${cls}">${money(view.finalValue)}</strong> today —
        a <strong class="${cls}">${net >= 0 ? "+" : ""}${net.toFixed(2)}%</strong> return after all charges (${fmtDateDMY(view.startDate)} → ${fmtDateDMY(view.endDate)}).
        Best it reached along the way: <strong class="text-emerald-700">+${up.toFixed(1)}%</strong>. Worst dip: <strong class="text-rose-700">${dd.toFixed(1)}%</strong>.
      </div>
    </div>`;
}

async function renderCustom() {
  const host = $("#custom-content");
  if (!host) return;
  host.innerHTML = `<div class="bg-white rounded-2xl ring-1 ring-slate-200 p-10 text-center text-slate-500 text-sm">Loading Custom Lab…</div>`;
  try {
    try { await ensureHistoryCache(); } catch (e) { host.innerHTML = renderHistoryEmpty(e.message); return; }
    await loadTechHistory();
    const { snapshots, benchmark, lkp } = state.cache.history;
    if (!snapshots.length) { host.innerHTML = renderHistoryEmpty("No snapshots loaded."); return; }
    // Open the Weight Lab anchored on the CURRENT live weights each time the
    // tab is opened (founder ask) — discard any prior sandbox tinkering so the
    // sliders always start from what the AI basket actually uses.
    state.labWeights = { ...(state.pillarWeights || composite.PILLAR_WEIGHTS) };
    state.labAiBest = null;
    const anchorDate = lkpAnchorDate(lkp, snapshots);
    const todayDate = snapshots[snapshots.length - 1].date;
    const niftyClosesByDate = benchmark?.indices?.["NIFTYSMLCAP250.NS"]?.closes || null;
    const niftyDatesSorted = niftyClosesByDate ? Object.keys(niftyClosesByDate).sort() : null;
    const niftyOn = (date) => { if (!niftyClosesByDate) return null; if (niftyClosesByDate[date] != null) return niftyClosesByDate[date]; let last = null; for (const d of niftyDatesSorted) { if (d <= date) last = niftyClosesByDate[d]; else break; } return last; };
    const lkpResolved = lkpOverride() || lkp;
    const anchorMonth = anchorDate?.slice(0, 7) || null;
    const mostRecentMonth = snapshots[snapshots.length - 1].date.slice(0, 7);
    const manualPicks = lkpResolved ? (lkpPicksForMonth(lkpResolved, anchorMonth, mostRecentMonth) || lkpResolved.picks || []) : [];

    // Granular technicals (today) — fetched once, joined by ticker, so
    // the indicator picker can evaluate live pass/fail per stock.
    if (!customTechByTicker) {
      try {
        const tj = await fetch("data/technicals.json").then((r) => r.json());
        customTechByTicker = new Map((tj.companies || []).map((c) => [c.ticker, c]));
      } catch { customTechByTicker = new Map(); }
    }

    const views = customStrategies.map((s) => ({ strat: s, view: buildCustomView(snapshots, anchorDate, todayDate, s, niftyOn, manualPicks) }));
    if (customPerfOpen) {
      host.innerHTML = renderCustomPerformancePage(views);
    } else if (customSelectedId && views.some((v) => v.strat.id === customSelectedId)) {
      const entry = views.find((v) => v.strat.id === customSelectedId);
      const todaysNames = buildTodaysNames(entry.strat, snapshots[snapshots.length - 1]);
      host.innerHTML = renderCustomDeepDive(entry, todaysNames);
    } else {
      customSelectedId = null;
      host.innerHTML = renderCustomOverview(views);
    }
    wireCustomTab();
    wireWeightLab();
    wireTechBacktest();
    wireMultiCurveHover("#custom-content");
    $$("#custom-content [data-cohort-row]").forEach((el) => el.addEventListener("click", () => {
      const ticker = el.dataset.ticker; if (!ticker) return;
      const pick = buildCohortClickPick(ticker, el.dataset.cohortSide || "ai", el.dataset.segAnchor || null);
      if (pick) openHistoryDrill(pick);
    }));
    $$("#custom-content [data-pick-toggle]").forEach((btn) => btn.addEventListener("click", () => {
      const side = btn.dataset.pickToggle; const list = $(`#custom-content [data-pick-list="${side}"]`); if (!list) return;
      const extra = list.querySelectorAll(".pick-extra-row"); const expand = btn.dataset.expanded !== "1";
      extra.forEach((el) => { el.style.display = expand ? "" : "none"; });
      btn.dataset.expanded = expand ? "1" : "0";
      btn.textContent = expand ? "Show less" : `+ ${extra.length} more`;
    }));
  } catch (e) {
    console.error("renderCustom failed:", e);
    host.innerHTML = `<div class="bg-white rounded-2xl ring-1 ring-rose-200 p-6"><div class="text-rose-600 font-bold text-sm">Custom Lab failed to render</div><pre class="text-[10px] text-slate-500 mt-2 whitespace-pre-wrap overflow-x-auto">${escapeHtml(e?.stack || String(e))}</pre></div>`;
  }
}

function wireCustomTab() {
  const root = "#custom-content";
  $$(`${root} [data-strat-new]`).forEach((b) => b.addEventListener("click", () => {
    const s = defaultStrategy("Strategy " + String.fromCharCode(65 + customStrategies.length));
    customStrategies.push(s); saveStrategies(customStrategies); customSelectedId = s.id; renderCustom();
  }));
  $$(`${root} [data-strat-open]`).forEach((b) => b.addEventListener("click", () => { customSelectedId = b.dataset.stratOpen; renderCustom(); }));
  $$(`${root} [data-strat-back]`).forEach((b) => b.addEventListener("click", () => { customSelectedId = null; renderCustom(); }));
  // Performance page — open from the top button, back to overview, or open
  // a strategy's deep-dive from a row (which leaves the performance page).
  $$(`${root} [data-perf-open]`).forEach((b) => b.addEventListener("click", () => { customPerfOpen = true; customSelectedId = null; renderCustom(); }));
  $$(`${root} [data-perf-close]`).forEach((b) => b.addEventListener("click", () => { customPerfOpen = false; renderCustom(); }));
  $$(`${root} [data-perf-row]`).forEach((tr) => tr.addEventListener("click", () => { customPerfOpen = false; customSelectedId = tr.dataset.perfRow; renderCustom(); }));
  $$(`${root} [data-strat-dup]`).forEach((b) => b.addEventListener("click", () => {
    const src = customStrategies.find((s) => s.id === b.dataset.stratDup); if (!src) return;
    const copy = { ...src, id: newStratId(), name: src.name.replace(/^★ /, "") + " (copy)", origin: "user" };
    customStrategies.push(copy); saveStrategies(customStrategies); customSelectedId = copy.id; renderCustom();
  }));
  $$(`${root} [data-strat-del]`).forEach((b) => b.addEventListener("click", () => {
    if (customStrategies.length <= 1) { alert("Keep at least one strategy."); return; }
    if (!confirm("Delete this strategy?")) return;
    const id = b.dataset.stratDel;
    customStrategies = customStrategies.filter((s) => s.id !== id); saveStrategies(customStrategies);
    if (customSelectedId === id) customSelectedId = null; renderCustom();
  }));
  $$(`${root} [data-strat-name]`).forEach((inp) => inp.addEventListener("change", () => {
    const s = customStrategies.find((x) => x.id === customSelectedId); if (!s) return;
    s.name = inp.value.trim() || s.name; saveStrategies(customStrategies); renderCustom();
  }));
  $$(`${root} [data-strat-field]`).forEach((inp) => {
    inp.addEventListener("input", () => {
      const out = $(`${root} [data-strat-out="${inp.dataset.stratField}"]`);
      if (out) { const f = STRAT_FIELDS.find((x) => x.key === inp.dataset.stratField); out.textContent = inp.value + (f ? f.suffix : ""); }
    });
    inp.addEventListener("change", () => {
      const s = customStrategies.find((x) => x.id === customSelectedId); if (!s) return;
      s[inp.dataset.stratField] = parseFloat(inp.value); saveStrategies(customStrategies); renderCustom();
    });
  });
  // Parameter-pillar dropdown — switch which pillar's params are shown
  // (Technicals / Fundamentals / Sentiment / Macro / Liquidity).
  $$(`${root} [data-param-pillar]`).forEach((sel) => sel.addEventListener("change", () => {
    paramPickerPillar = sel.value; paramAdvancedOpen = true; rerenderKeepingScroll(renderCustom);
  }));
  // Keep the Advanced <details> open/closed state across re-renders.
  $$(`${root} [data-param-advanced]`).forEach((el) => el.addEventListener("toggle", () => {
    paramAdvancedOpen = el.open;
  }));
  // Parameter toggles — flip on/off, seed defaults, re-run.
  $$(`${root} [data-param-on]`).forEach((inp) => inp.addEventListener("change", () => {
    const s = customStrategies.find((x) => x.id === customSelectedId); if (!s) return;
    const key = inp.dataset.paramOn;
    const def = STRAT_PARAMS.find((p) => p.key === key);
    s.params = s.params || {};
    if (inp.checked) s.params[key] = { on: true, ...(def?.def || {}), ...(s.params[key] || {}), on: true };
    else s.params[key] = { ...(s.params[key] || {}), on: false };
    paramAdvancedOpen = true;
    saveStrategies(customStrategies); rerenderKeepingScroll(renderCustom);
  }));
  // Parameter value tweaks (min / max / val).
  $$(`${root} [data-param-input]`).forEach((inp) => inp.addEventListener("change", () => {
    const s = customStrategies.find((x) => x.id === customSelectedId); if (!s) return;
    const key = inp.dataset.paramInput, which = inp.dataset.paramWhich;
    const num = parseFloat(inp.value); if (!Number.isFinite(num)) return;
    s.params = s.params || {};
    s.params[key] = { ...(s.params[key] || { on: true }), [which]: num };
    paramAdvancedOpen = true;
    saveStrategies(customStrategies); rerenderKeepingScroll(renderCustom);
  }));
  // Shared capital & charges (reused renderSimPanel) — re-run all strategies.
  $$(`${root} [data-sim-field]`).forEach((inp) => inp.addEventListener("change", () => {
    const key = inp.dataset.simField; const num = parseFloat(inp.value);
    if (!Number.isFinite(num) || num < 0) { inp.value = simPrefs[key]; return; }
    simPrefs = { ...simPrefs, [key]: num }; saveSimPrefs(simPrefs); renderCustom();
  }));
  const sr = $(`${root} #sim-reset`);
  if (sr) sr.addEventListener("click", () => { simPrefs = { ...SIM_DEFAULTS }; saveSimPrefs(simPrefs); renderCustom(); });
  // Sector ↔ Industry toggle on the rebalance-timing table.
  wireSectorTimingToggle(root, renderCustom);
}

// ============================================================
// ALERTS (Phase 5)
// ============================================================
// Configurable alert center: outcome alerts (target / SL hits today on
// the active strategy) and signal alerts (volume / momentum / proximity
// / composite cross) evaluated against the latest snapshot + technicals.
// Each rule has an on/off toggle and (where relevant) a threshold. All
// adjustable, persisted locally. A nav badge shows the live count.
const ALERT_PREFS_KEY = "klpdash-alert-prefs-v1";
// Only basket-outcome alerts: a pick hit its target (upside) or its
// stop-loss (downside) today. The old signal alerts (volume, RSI, near
// 52-week high, ADX, composite-crossed) were noise for this product and
// were removed.
const ALERT_DEFS = [
  { key: "targetHit",      label: "🎯 Target hit today",       desc: "A basket pick crossed its profit target today" },
  { key: "slHit",          label: "⚠ Stop-loss hit today",     desc: "A basket pick hit its stop-loss today" },
];
const ALERT_DEFAULTS = {
  minComposite: 60,
  rules: {
    targetHit: { on: true },
    slHit: { on: true },
  },
};
function loadAlertPrefs() {
  try {
    const raw = JSON.parse(localStorage.getItem(ALERT_PREFS_KEY));
    if (raw && raw.rules) return { ...ALERT_DEFAULTS, ...raw, rules: { ...ALERT_DEFAULTS.rules, ...raw.rules } };
  } catch {}
  return JSON.parse(JSON.stringify(ALERT_DEFAULTS));
}
function saveAlertPrefs(p) { try { localStorage.setItem(ALERT_PREFS_KEY, JSON.stringify(p)); } catch {} }
let alertPrefs = loadAlertPrefs();

// Evaluate every enabled rule against the latest data. Returns grouped
// alert lists [{ key, label, items:[{ticker,name,sector,composite,detail}], total }].
function evaluateAlerts(snapshots, techByTicker, strategyHits, prefs) {
  const R = prefs.rules;
  const out = [];

  if (R.targetHit?.on) {
    const items = (strategyHits?.todayHits || []).filter((h) => h.status === "TARGET_HIT")
      .map((h) => ({ ticker: h.ticker, name: h.name || h.ticker, sector: h.sector, composite: null, detail: `${h.basket} · hit ₹${formatPrice(h.exitPrice)}` }));
    if (items.length) out.push({ key: "targetHit", label: "🎯 Target hit today", items, total: items.length });
  }
  if (R.slHit?.on) {
    const items = (strategyHits?.todayHits || []).filter((h) => h.status === "SL_HIT")
      .map((h) => ({ ticker: h.ticker, name: h.name || h.ticker, sector: h.sector, composite: null, detail: `${h.basket} · hit ₹${formatPrice(h.exitPrice)}` }));
    if (items.length) out.push({ key: "slHit", label: "⚠ Stop-loss hit today", items, total: items.length });
  }
  return out;
}

function renderAlertConfig(prefs) {
  const rows = ALERT_DEFS.map((d) => {
    const r = prefs.rules[d.key] || { on: false };
    const thr = d.unit
      ? `<input type="number" data-alert-threshold="${d.key}" value="${r.threshold ?? ""}" step="${d.step}" class="w-16 rounded-lg ring-1 ring-slate-200 px-2 py-1 text-xs tabular-nums outline-none focus:ring-2 focus:ring-indigo-300" /><span class="text-[10px] text-slate-400 whitespace-nowrap">${d.unit}</span>`
      : "";
    return `
      <div class="flex items-center justify-between gap-3 py-2 border-t border-slate-100">
        <label class="flex items-center gap-2 cursor-pointer min-w-0">
          <input type="checkbox" data-alert-on="${d.key}" ${r.on ? "checked" : ""} class="accent-indigo-600 w-4 h-4 flex-shrink-0" />
          <span class="min-w-0"><span class="text-sm font-semibold text-slate-800">${d.label}</span><span class="block text-[10px] text-slate-400 truncate">${d.desc}</span></span>
        </label>
        <div class="flex items-center gap-1.5 flex-shrink-0">${thr}</div>
      </div>`;
  }).join("");
  return `
    <details class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <summary class="cursor-pointer flex items-center gap-2 select-none">
        <span class="text-base">⚙</span><span class="font-display font-bold text-slate-900 text-sm">Alert settings</span><span class="text-[11px] text-slate-400 font-normal">— turn rules on/off, set thresholds</span>
      </summary>
      <div class="mt-3">
        <div class="flex items-center justify-between gap-3 py-2">
          <span class="text-xs text-slate-600">Only alert on stocks scoring ≥</span>
          <div class="flex items-center gap-2"><input type="number" data-alert-mincomposite value="${prefs.minComposite}" step="1" class="w-16 rounded-lg ring-1 ring-slate-200 px-2 py-1 text-xs tabular-nums outline-none focus:ring-2 focus:ring-indigo-300" /><button type="button" id="alert-reset" class="text-[11px] font-semibold text-slate-500 hover:text-indigo-600">Reset</button></div>
        </div>
        ${rows}
        <div class="text-[10px] text-slate-400 mt-2 leading-snug">Checked = on. The number next to each rule is its trigger level (e.g. RSI ≥ 70). Evaluated on the latest data; click any result to open that stock's chart.</div>
      </div>
    </details>`;
}

function renderAlertFeed(alerts) {
  if (!alerts.length) {
    return `<div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-8 text-center text-sm text-slate-500">No alerts firing right now. Loosen a threshold or enable more rules above.</div>`;
  }
  return `<div class="grid grid-cols-1 md:grid-cols-2 gap-3">${alerts.map((g) => {
    const shown = g.items.slice(0, 6);
    return `
    <div class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4">
      <div class="flex items-center justify-between mb-2">
        <h3 class="font-display font-bold text-slate-900 text-sm">${g.label}</h3>
        <span class="text-[11px] font-semibold text-slate-500">${g.total} match${g.total === 1 ? "" : "es"}</span>
      </div>
      <div class="space-y-0.5">
        ${shown.map((it) => `<button type="button" data-cohort-row data-cohort-side="ai" data-ticker="${escapeHtml(it.ticker)}" class="w-full text-left flex items-center justify-between gap-2 py-1.5 px-2 rounded-lg hover:bg-indigo-50/50 transition">
          <div class="min-w-0"><div class="text-xs font-semibold text-slate-800 truncate">${escapeHtml(it.name)}</div><div class="text-[10px] text-slate-400 truncate">${escapeHtml(it.sector || "")}</div></div>
          <div class="text-right flex-shrink-0"><div class="text-[11px] font-bold tabular-nums text-indigo-700">${escapeHtml(it.detail)}</div>${it.composite != null ? `<div class="text-[9px] text-slate-400">score ${it.composite.toFixed(1)}</div>` : ""}</div>
        </button>`).join("")}
        ${g.total > shown.length ? `<div class="text-[10px] text-slate-400 px-2 pt-1">+ ${g.total - shown.length} more</div>` : ""}
      </div>
    </div>`; }).join("")}</div>`;
}

// Alerts as a compact, collapsible section inside the Strategy tab
// (no longer a standalone tab). Collapsed by default with a live count
// in the summary; expands to the feed + settings.
function renderAlertsSection(alerts, prefs) {
  const total = alerts.reduce((a, g) => a + g.total, 0);
  return `
    <details class="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-4 sm:p-5">
      <summary class="cursor-pointer flex items-center gap-2 select-none flex-wrap">
        <span class="text-base">🔔</span>
        <span class="font-display font-bold text-slate-900 text-sm">Alerts — basket picks that hit today</span>
        <span class="text-[10px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded ${total ? "bg-rose-100 text-rose-700 ring-1 ring-rose-200" : "bg-slate-100 text-slate-500 ring-1 ring-slate-200"}">${total} live</span>
        <span class="text-[11px] text-slate-400 font-normal hidden sm:inline">target &amp; stop-loss hits</span>
      </summary>
      <div class="mt-3 space-y-3">
        ${renderAlertFeed(alerts)}
        ${renderAlertConfig(prefs)}
      </div>
    </details>`;
}

// Wire the alert toggles / thresholds within a given root, re-rendering
// via the supplied callback (the Strategy tab passes renderActive).
function wireAlertsInputs(root, rerender) {
  $$(`${root} [data-alert-on]`).forEach((inp) => inp.addEventListener("change", () => {
    const key = inp.dataset.alertOn;
    alertPrefs.rules[key] = { ...(alertPrefs.rules[key] || {}), on: inp.checked };
    saveAlertPrefs(alertPrefs); rerender();
  }));
  $$(`${root} [data-alert-threshold]`).forEach((inp) => inp.addEventListener("change", () => {
    const key = inp.dataset.alertThreshold; const num = parseFloat(inp.value);
    if (!Number.isFinite(num)) { inp.value = alertPrefs.rules[key]?.threshold ?? ""; return; }
    alertPrefs.rules[key] = { ...(alertPrefs.rules[key] || {}), threshold: num };
    saveAlertPrefs(alertPrefs); rerender();
  }));
  const mc = $(`${root} [data-alert-mincomposite]`);
  if (mc) mc.addEventListener("change", () => {
    const num = parseFloat(mc.value);
    if (Number.isFinite(num)) { alertPrefs.minComposite = num; saveAlertPrefs(alertPrefs); rerender(); }
  });
  const reset = $(`${root} #alert-reset`);
  if (reset) reset.addEventListener("click", () => { alertPrefs = JSON.parse(JSON.stringify(ALERT_DEFAULTS)); saveAlertPrefs(alertPrefs); rerender(); });
}

wire();
switchTab("fundamentals");
