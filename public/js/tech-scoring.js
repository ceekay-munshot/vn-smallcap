// Technicals scoring — rules taken verbatim from the client's scoring sheet.
// Same {points, max, status, value, note} shape as fundamentals scoring.js
// so the shared dashboard renderer can consume either.

const NA = { points: 0, max: 0, status: "na", value: null, note: "Data not available" };

function fmtPct(n, d = 1) { return n == null ? "—" : `${(n * 100).toFixed(d)}%`; }
function fmtNum(n, d = 1) { return n == null ? "—" : Number(n).toFixed(d); }

// ---- rules (11 active; 5 deferred) ----

function rulePriceAbove50EMA(c) {
  if (c.cmp == null || c.ema50 == null) return { ...NA, max: 2 };
  const val = `CMP ${fmtNum(c.cmp,0)} vs 50 EMA ${fmtNum(c.ema50,0)}`;
  if (c.above_50ema) return { points: 2, max: 2, status: "pass", value: val, note: "Confirms short-to-medium term uptrend." };
  return { points: 0, max: 2, status: "fail", value: val, note: "Below 50 EMA — short-to-medium term weakness." };
}

function rulePriceAbove200DMA(c) {
  if (c.cmp == null || c.sma200 == null) return { ...NA, max: 2, note: "Less than 200 trading days of history — primary trend cannot be evaluated." };
  const val = `CMP ${fmtNum(c.cmp,0)} vs 200 DMA ${fmtNum(c.sma200,0)}`;
  if (c.above_200dma) return { points: 2, max: 2, status: "pass", value: val, note: "Above 200 DMA — primary trend filter passes." };
  return { points: 0, max: 2, status: "hard_fail", value: val, note: "Below 200 DMA — primary trend filter fail. Client framework: 'Price < 200 DMA = immediate fail — stock exits pipeline.' Excluded from SPIP basket." };
}

function ruleGoldenCross(c) {
  if (c.sma50 == null || c.sma200 == null) return { ...NA, max: 1, note: "Need both 50 SMA and 200 SMA — insufficient history." };
  const val = `50 DMA ${fmtNum(c.sma50,0)} vs 200 DMA ${fmtNum(c.sma200,0)}`;
  if (c.golden_cross) return { points: 1, max: 1, status: "pass", value: val, note: "Golden Cross active (50 DMA > 200 DMA)." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Death Cross — 50 DMA below 200 DMA." };
}

function ruleRSI(c) {
  const v = c.rsi14;
  if (v == null) return { ...NA, max: 2 };
  const val = `RSI ${fmtNum(v,1)}`;
  if (v >= 55 && v <= 75) return { points: 2, max: 2, status: "pass", value: val, note: "RSI in ideal momentum zone (55–75)." };
  if (v > 75 && v <= 80) return { points: 1, max: 2, status: "partial", value: val, note: "RSI 75–80 — strong but approaching overbought." };
  if (v > 80) return { points: 1, max: 2, status: "partial", value: val, note: "RSI > 80 — overbought zone." };
  if (v >= 40 && v < 55) return { points: 1, max: 2, status: "partial", value: val, note: "RSI 40–55 — neutral / consolidating." };
  return { points: 0, max: 2, status: "fail", value: val, note: "RSI < 40 — weak / failing momentum." };
}

function ruleMACD(c) {
  if (!c.macd) return { ...NA, max: 2 };
  const { line, signal, above_zero, positive_crossover } = c.macd;
  const val = `Line ${fmtNum(line,2)} | Signal ${fmtNum(signal,2)}`;
  if (line > signal && above_zero) {
    const cross = positive_crossover ? " (fresh positive crossover)" : "";
    return { points: 2, max: 2, status: "pass", value: val, note: `MACD line above signal and above zero${cross}.` };
  }
  if (line > signal) return { points: 1, max: 2, status: "partial", value: val, note: "MACD above signal but still below zero." };
  return { points: 0, max: 2, status: "fail", value: val, note: "Negative crossover — MACD line below signal." };
}

function ruleADX(c) {
  const v = c.adx14;
  if (v == null) return { ...NA, max: 1 };
  const val = `ADX ${fmtNum(v,1)}`;
  if (v > 25) return { points: 1, max: 1, status: "pass", value: val, note: "ADX > 25 — strong trend." };
  if (v >= 20) return { points: 0.5, max: 1, status: "partial", value: val, note: "ADX 20–25 — developing trend (0.5 pt per client framework)." };
  return { points: 0, max: 1, status: "fail", value: val, note: "ADX < 20 — choppy / no clear trend." };
}

function ruleRelativeStrength(c) {
  const rs = c.relative_strength_6m;
  const stockR = c.return_6m;
  const indexR = c.return_6m_index;
  if (rs == null) return { ...NA, max: 2 };
  const val = `6M return ${fmtPct(stockR,1)} vs Nifty 500 ${fmtPct(indexR,1)} (RS ${rs>=0?"+":""}${fmtPct(rs,1)})`;
  if (rs > 0) return { points: 2, max: 2, status: "pass", value: val, note: "6M price return ahead of Nifty 500 — outperforming." };
  return { points: 1, max: 2, status: "partial", value: val, note: "Underperforming benchmark." };
}

function ruleVolumeBreakout(c) {
  const r = c.volume_ratio_today;
  if (r == null) return { ...NA, max: 2 };
  const val = `Today vs 20-day avg: ${fmtNum(r,2)}×`;
  if (r >= 1.5) return { points: 2, max: 2, status: "pass", value: val, note: "Today's volume ≥ 1.5× 20-day average — strong participation." };
  if (r >= 1.0) return { points: 1, max: 2, status: "partial", value: val, note: "Slightly above average — low-volume move (suspect)." };
  return { points: 0, max: 2, status: "fail", value: val, note: "Below average volume — weak conviction." };
}

function rule52WHighProximity(c) {
  const p = c.high_proximity_pct;
  if (p == null) return { ...NA, max: 2 };
  const distance = (1 - p) * 100;
  const val = `${fmtNum(distance,1)}% below 52W high (CMP ${fmtNum(c.cmp,0)} / 52W ${fmtNum(c.high_52w,0)})`;
  if (distance <= 10) return { points: 2, max: 2, status: "pass", value: val, note: "Within 10% of 52-week high — strong structure." };
  if (distance <= 20) return { points: 1, max: 2, status: "partial", value: val, note: "10–20% off 52W high — still constructive." };
  return { points: 0, max: 2, status: "fail", value: val, note: ">20% below 52W high — weak structure." };
}

function ruleBeta(c) {
  const b = c.beta_1y;
  if (b == null) return { ...NA, max: 1 };
  const val = `Beta ${fmtNum(b,2)}`;
  if (b >= 0.7 && b <= 1.3) return { points: 1, max: 1, status: "pass", value: val, note: "Beta 0.7–1.3 — moderate market sensitivity." };
  if (b > 1.5) return { points: 0, max: 1, status: "fail", value: val, note: "Beta > 1.5 — high risk, position-size penalty." };
  if (b < 0.5) return { points: 0, max: 1, status: "fail", value: val, note: "Beta < 0.5 — drag risk (low responsiveness)." };
  return { points: 1, max: 1, status: "partial", value: val, note: "Beta in transitional range." };
}

function ruleATRStability(c) {
  // Per client framework: PASS if 14-day ATR% is declining or stable AND
  // < 2.5% for large-caps. We check both the absolute level + the trend
  // (using atr_history accumulator built up over recent daily scrapes).
  const a = c.atr14_pct;
  if (a == null) return { ...NA, max: 1 };
  const hist = Array.isArray(c.atr_history) ? c.atr_history : [];
  const val = `ATR(14) = ${fmtNum(a,2)}% of price`;
  // Trend assessment from accumulator (when enough history available)
  let trend = null;     // "declining" | "stable" | "rising" | null (unknown)
  let trendNote = "";
  if (hist.length >= 10) {
    const half = Math.floor(hist.length / 2);
    const recent = hist.slice(-half).map((e) => e.atr_pct);
    const older  = hist.slice(0, half).map((e) => e.atr_pct);
    const avgR = recent.reduce((x, y) => x + y, 0) / recent.length;
    const avgO = older.reduce((x, y) => x + y, 0) / older.length;
    const diff = avgR - avgO;
    if (diff < -0.1) trend = "declining";
    else if (diff > 0.2) trend = "rising";
    else trend = "stable";
    trendNote = ` · ${hist.length}-day trend ${trend} (${avgO.toFixed(2)}% → ${avgR.toFixed(2)}%)`;
  } else if (hist.length > 0) {
    trendNote = ` · only ${hist.length}-day history yet (need ≥ 10 for trend)`;
  } else {
    trendNote = " · trend pending (accumulator building from today)";
  }
  // Score: absolute level first, then trend modifier
  if (a < 2.5 && (trend === "declining" || trend === "stable" || trend == null)) {
    return { points: 1, max: 1, status: "pass", value: val + trendNote, note: "ATR% under 2.5% and not rising — controlled volatility." };
  }
  if (a < 2.5 && trend === "rising") {
    return { points: 0.5, max: 1, status: "partial", value: val + trendNote, note: "ATR% absolutely OK (<2.5%) but trend rising — position-size flag per client framework." };
  }
  if (a < 4 && trend === "declining") {
    return { points: 1, max: 1, status: "pass", value: val + trendNote, note: "ATR% 2.5–4% but trending down — improving volatility profile." };
  }
  if (a < 4) {
    return { points: 0.5, max: 1, status: "partial", value: val + trendNote, note: "ATR% 2.5–4% — elevated but tolerable." };
  }
  return { points: 0, max: 1, status: "fail", value: val + trendNote, note: "ATR% > 4% — high volatility, position-size flag." };
}

function ruleInstitutionalActivity(c) {
  const fii = c.chg_fii_hold;
  const dii = c.chg_dii_hold;
  if (fii == null && dii == null) {
    return { ...NA, max: 1, note: "FII / DII holding change not available for this company in the latest Screener shareholding data." };
  }
  const sum = (fii ?? 0) + (dii ?? 0);
  const val = `Chg FII ${fii == null ? "—" : (fii > 0 ? "+" : "") + fmtNum(fii, 2) + "%"} | Chg DII ${dii == null ? "—" : (dii > 0 ? "+" : "") + fmtNum(dii, 2) + "%"}`;
  if (sum > 0) {
    if (fii != null && fii < -2) {
      return { points: 1, max: 1, status: "partial", value: val, note: "Combined institutional buying, but FII exiting sharply (>2%) — caution." };
    }
    return { points: 1, max: 1, status: "pass", value: val, note: "Net FII + DII buying in latest period — institutional accumulation." };
  }
  return { points: 0, max: 1, status: "fail", value: val, note: "Net institutional selling in latest period." };
}

function ruleDeliveryPercentage(c) {
  if (c.delivery_trend_diff == null) {
    if (c.error || c.delivery_days_count === 0) {
      return { ...NA, max: 1, note: "Delivery % data not available — NSE bhavcopy did not include this ticker, or files could not be fetched." };
    }
    return { ...NA, max: 1, note: "Insufficient delivery % history (fewer than 6 trading days returned by NSE)." };
  }
  const diff = c.delivery_trend_diff;
  const val = `Recent ~15d avg ${fmtNum(c.delivery_avg_recent,1)}% vs preceding ~15d ${fmtNum(c.delivery_avg_older,1)}% (Δ ${diff > 0 ? "+" : ""}${fmtNum(diff,1)} pp)`;
  if (diff > 1) return { points: 1, max: 1, status: "pass", value: val, note: "Delivery % clearly rising over the last 30 trading days — stronger holding conviction." };
  if (diff > 0) return { points: 1, max: 1, status: "partial", value: val, note: "Delivery % slightly higher — mildly improving conviction." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Delivery % flat or declining — weaker holding conviction." };
}

function ruleHHHL(c) {
  const p = c?.higher_highs_lows;
  if (!p) return { ...NA, max: 1, note: "Insufficient history (need ~6 months of bars) to evaluate HH-HL pattern." };
  const val = `Recent 3m hi ${fmtNum(p.recent_high, 0)} / lo ${fmtNum(p.recent_low, 0)} vs prior 3m hi ${fmtNum(p.prior_high, 0)} / lo ${fmtNum(p.prior_low, 0)}`;
  if (p.pattern_present) return { points: 1, max: 1, status: "pass", value: val, note: "Higher highs + higher lows over last ~6 months." };
  if (p.higher_high || p.higher_low) return { points: 1, max: 1, status: "partial", value: val, note: "Only one side is higher (either highs or lows but not both)." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Neither highs nor lows are advancing — no HH-HL structure." };
}

function ruleConsolidationBreakout(c) {
  const p = c?.consolidation_breakout;
  if (!p) return { ...NA, max: 2, note: "Insufficient history (need 6+ weeks) to evaluate consolidation breakout." };
  const val = `Base range ${fmtNum(p.base_range_pct, 1)}% · today ₹${fmtNum(p.today_close, 0)} vs base hi ₹${fmtNum(p.base_max, 0)}${p.today_volume_ratio != null ? ` · vol ${fmtNum(p.today_volume_ratio, 1)}×` : ""}`;
  if (p.quality === "strong")    return { points: 2, max: 2, status: "pass", value: val, note: "Tight 6-week base + breakout + volume confirmation." };
  if (p.quality === "weak_base") return { points: 1, max: 2, status: "partial", value: val, note: "Breakout with volume but base wasn't tight." };
  if (p.quality === "low_volume") return { points: 1, max: 2, status: "partial", value: val, note: "Breakout above base high but on low volume — suspect." };
  return { points: 0, max: 2, status: "fail", value: val, note: "No breakout — price still inside (or below) the recent range." };
}

function ruleBaseFormation(c) {
  const p = c?.base_formation;
  if (!p) return { ...NA, max: 1, note: "Insufficient history (need ~12 weeks) to evaluate base health." };
  const val = `Drawdown ${fmtNum(p.drawdown_pct, 1)}% · closing-range tightness ${fmtNum(p.tightness_pct, 1)}%`;
  if (p.healthy_base) return { points: 1, max: 1, status: "pass", value: val, note: "Healthy 12-week base — controlled drawdown (<15%) and tight closing range (<4%)." };
  if (p.drawdown_pct < 20) return { points: 1, max: 1, status: "partial", value: val, note: "Acceptable base but either drawdown or closing range looser than ideal." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Drawdown > 20% or closing range too wide — no clean base." };
}

// ---- deferred (0 — all rules now active) ----
const DEFERRED = [];

const ACTIVE_RULES = [
  { key: "ema50",    label: "Price Above 50 EMA",       category: "Trend Strength", criteria: "CMP > 50 EMA",          fn: rulePriceAbove50EMA },
  { key: "dma200",   label: "Price Above 200 DMA",      category: "Trend Strength", criteria: "CMP > 200 DMA",         fn: rulePriceAbove200DMA },
  { key: "gold",     label: "Golden Cross",             category: "Trend Strength", criteria: "50 DMA > 200 DMA",      fn: ruleGoldenCross },
  { key: "hhhl",     label: "Higher Highs–Higher Lows", category: "Trend Strength", criteria: "Pattern over ~6 months", fn: ruleHHHL },
  { key: "rsi",      label: "RSI (14)",                 category: "Momentum",       criteria: "55–75",                 fn: ruleRSI },
  { key: "macd",     label: "MACD",                     category: "Momentum",       criteria: "Positive crossover",    fn: ruleMACD },
  { key: "adx",      label: "ADX (14)",                 category: "Momentum",       criteria: "> 25",                  fn: ruleADX },
  { key: "rs",       label: "Relative Strength vs Nifty 500", category: "Momentum", criteria: "Outperforming index",   fn: ruleRelativeStrength },
  { key: "volbo",    label: "Volume Breakout",          category: "Volume",         criteria: "≥ 1.5× 20-day avg",     fn: ruleVolumeBreakout },
  { key: "delivery", label: "Delivery Percentage",      category: "Volume",         criteria: "Rising over 30 days",   fn: ruleDeliveryPercentage },
  { key: "instact",  label: "Institutional Activity",   category: "Volume",         criteria: "Net FII + DII buying",  fn: ruleInstitutionalActivity },
  { key: "near52w",  label: "52-Week High Proximity",   category: "Breakout",       criteria: "Within 10%",            fn: rule52WHighProximity },
  { key: "consolidation", label: "Breakout from Consolidation", category: "Breakout", criteria: "6-week base + volume", fn: ruleConsolidationBreakout },
  { key: "base",     label: "Base Formation",           category: "Breakout",       criteria: "Drawdown <15% + tight closes", fn: ruleBaseFormation },
  { key: "beta",     label: "Beta",                     category: "Risk",           criteria: "0.7 – 1.3",             fn: ruleBeta },
  { key: "atr",      label: "ATR Stability",            category: "Risk",           criteria: "< 2.5% of price",       fn: ruleATRStability },
];

export function scoreCompany(c) {
  if (c.error) {
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

// Compact per-indicator value record for the Custom Lab's tweakable
// parameter filters — raw numbers + boolean flags, so a strategy can
// test ANY threshold (RSI ≥ 65, within 5% of the high, Beta 0.7–1.3…)
// live and historically. Shared by the snapshot writer and the app so
// both evaluate identically. null when there's no usable tech row.
export function techVals(c) {
  if (!c || c.cmp == null) return null;
  const macdPos = !!(c.macd && (c.macd.positive_crossover || c.macd.above_zero));
  const num = (v) => (typeof v === "number" ? v : null);
  return {
    rsi:  num(c.rsi14),
    adx:  num(c.adx14),
    vol:  num(c.volume_ratio_today),
    d52:  c.high_proximity_pct != null ? Number(((1 - c.high_proximity_pct) * 100).toFixed(2)) : null,
    beta: num(c.beta_1y),
    atr:  num(c.atr14_pct),
    e50:  !!c.above_50ema,
    d200: !!c.above_200dma,
    gc:   !!c.golden_cross,
    hh:   !!c.higher_highs_lows,
    macd: macdPos,
    rs:   c.relative_strength_6m != null ? c.relative_strength_6m > 0 : null,
    dlv:  c.delivery_trend_diff != null ? c.delivery_trend_diff > 0 : null,
    inst: (c.chg_fii_hold != null || c.chg_dii_hold != null) ? ((c.chg_fii_hold || 0) + (c.chg_dii_hold || 0)) > 0 : null,
    cons: !!c.consolidation_breakout,
    base: !!c.base_formation,
  };
}

export { ACTIVE_RULES, DEFERRED };
