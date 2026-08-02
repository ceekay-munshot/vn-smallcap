// ============================================================
// THE FIVE STRATEGIES
//
// One ranking, five sets of rules on top of it. Each strategy adds a
// layer to the one before it, and the Compare tab has to prove that
// helped -- nothing here asserts that a higher number is better.
//
// Two execution paths share these definitions:
//
//   LIVE  -- selectBasket() runs against a daily snapshot, where every
//            indicator already exists in stock.techVals.
//   BACK  -- backtest() runs against data/history-technical.json, which
//            carries only { close, score } per ticker per day. The same
//            indicators are recomputed from the close series so the two
//            paths gate on like for like.
//
// Honest limits, surfaced in the UI rather than buried:
//   - The history file is today's universe replayed backwards, so names
//     that delisted or left the market-cap band are absent (survivorship).
//   - It has no volume, so the liquidity gate cannot be backtested. It
//     applies live only.
//   - It has no high/low, so ATR is approximated from close-to-close
//     moves and will read slightly tighter than a true ATR.
// ============================================================

import { makeCharger, ZERO_CHARGER } from "./charges.js";

// ---------------- definitions ----------------

let DEFS = null;

export async function loadStrategies() {
  if (DEFS) return DEFS;
  try {
    const j = await fetch("data/strategies.json").then((r) => (r.ok ? r.json() : null));
    if (j?.strategies?.length) DEFS = j;
  } catch {}
  if (!DEFS) DEFS = { version: 0, primaryId: "s5", strategies: [] };
  return DEFS;
}

// Inject definitions directly — used by the offline node harness, which
// has no fetch and reads the JSON off disk instead.
export function setStrategies(defs) { DEFS = defs; return DEFS; }

export function allStrategies() { return DEFS?.strategies || []; }
export function strategyById(id) { return (DEFS?.strategies || []).find((s) => s.id === id) || null; }
export function primaryId() { return DEFS?.primaryId || "s5"; }
export function setPrimaryId(id) { if (DEFS) DEFS.primaryId = id; }

// The file the user is told to commit after changing #1.
export function strategiesFileJson() {
  return JSON.stringify(DEFS, null, 2) + "\n";
}

// ---------------- indicator maths (backtest path) ----------------

function sma(arr, end, n) {
  if (end + 1 < n) return null;
  let s = 0;
  for (let i = end - n + 1; i <= end; i++) { const v = arr[i]; if (v == null) return null; s += v; }
  return s / n;
}

function ema(arr, end, n) {
  const start = Math.max(0, end - n * 3);
  let e = null;
  const k = 2 / (n + 1);
  for (let i = start; i <= end; i++) {
    const v = arr[i]; if (v == null) continue;
    e = e == null ? v : v * k + e * (1 - k);
  }
  return e;
}

// Wilder RSI over the last `n` closes ending at `end`.
function rsi(arr, end, n = 14) {
  if (end < n) return null;
  let gain = 0, loss = 0;
  for (let i = end - n + 1; i <= end; i++) {
    const a = arr[i - 1], b = arr[i];
    if (a == null || b == null) return null;
    const d = b - a;
    if (d >= 0) gain += d; else loss -= d;
  }
  if (gain + loss === 0) return 50;
  const rs = (gain / n) / ((loss / n) || 1e-9);
  return 100 - 100 / (1 + rs);
}

// Mean absolute daily move as a percentage -- a close-only stand-in for
// ATR%. Slightly tighter than a true ATR because it never sees the
// intraday range, so stops derived from it are marginally optimistic.
function atrPct(arr, end, n = 14) {
  if (end < n) return null;
  let s = 0, c = 0;
  for (let i = end - n + 1; i <= end; i++) {
    const a = arr[i - 1], b = arr[i];
    if (a == null || b == null || a === 0) continue;
    s += Math.abs(b / a - 1); c++;
  }
  return c ? (s / c) * 100 * 1.35 : null;   // 1.35 bridges close-to-close -> true range
}

// How far below its trailing 1-year high, as a positive percentage.
function offHigh(arr, end, n = 252) {
  const start = Math.max(0, end - n + 1);
  let hi = -Infinity;
  for (let i = start; i <= end; i++) { const v = arr[i]; if (v != null && v > hi) hi = v; }
  const c = arr[end];
  if (!Number.isFinite(hi) || hi <= 0 || c == null) return null;
  return ((hi - c) / hi) * 100;
}

// Mean of the last `n` daily scores — nulls skipped, not treated as zero.
function avgScore(arr, end, n) {
  let s = 0, c = 0;
  for (let i = Math.max(0, end - n + 1); i <= end; i++) { const v = arr[i]; if (v != null) { s += v; c++; } }
  return c ? s / c : null;
}

function retOver(arr, end, n) {
  const a = arr[end - n], b = arr[end];
  if (a == null || b == null || a === 0) return null;
  return (b / a - 1) * 100;
}

// Everything the gates need, computed once per ticker per day.
function indicatorsAt(closes, i) {
  const c = closes[i];
  if (c == null) return null;
  const ma50 = sma(closes, i, 50), ma200 = sma(closes, i, 200);
  const e12 = ema(closes, i, 12), e26 = ema(closes, i, 26);
  return {
    close: c,
    rsi: rsi(closes, i),
    atr: atrPct(closes, i),
    d52: offHigh(closes, i),
    aboveMa50: ma50 != null ? c > ma50 : null,
    aboveMa200: ma200 != null ? c > ma200 : null,
    macdPositive: e12 != null && e26 != null ? e12 > e26 : null,
    ret3m: retOver(closes, i, 63),
  };
}

// ---------------- gates ----------------
// A gate returns false only when it can prove the stock fails. A missing
// input passes, so a data hole never silently empties the basket -- the
// alternative (fail closed) would make the backtest look artificially
// selective in exactly the periods where data is thin.

function passesGates(ind, gates, ctx) {
  if (!gates) return true;
  if (gates.maxRsi != null && ind.rsi != null && ind.rsi > gates.maxRsi) return false;
  if (gates.minOffHigh != null && ind.d52 != null && ind.d52 < gates.minOffHigh) return false;
  if (gates.aboveMa50 && ind.aboveMa50 === false) return false;
  if (gates.aboveMa200 && ind.aboveMa200 === false) return false;
  if (gates.macdPositive && ind.macdPositive === false) return false;
  if (gates.relativeStrength && ind.ret3m != null && ctx?.medianRet3m != null && ind.ret3m <= ctx.medianRet3m) return false;
  // Liquidity is live-only: the history file carries no volume.
  if (gates.minAdtv != null && ind.adtv != null && ind.adtv < gates.minAdtv) return false;
  return true;
}

// Sector cap, applied while walking a ranked list.
function sectorCapped(picked, sector, gates) {
  if (!gates?.maxPerSector || !sector) return false;
  return picked.filter((p) => p.sector === sector).length >= gates.maxPerSector;
}

// ---------------- regime ----------------
// Full size when the index is above its 200 day average and VIX is calm;
// half when one turns; nothing when both do. Returns 1 / 0.5 / 0.

export function regimeExposure(strategy, ctx) {
  if (!strategy?.regime) return { exposure: 1, reason: "No regime overlay" };
  const trendOk = ctx?.indexAboveMa200;
  const vixOk = ctx?.vix == null ? true : ctx.vix < (strategy.regimeVixHigh ?? 20);
  if (trendOk == null) return { exposure: 1, reason: "Index history unavailable" };
  if (trendOk && vixOk) return { exposure: 1, reason: "Index above 200 DMA, VIX calm" };
  if (!trendOk && !vixOk) return { exposure: 0, reason: "Index below 200 DMA and VIX elevated" };
  return { exposure: 0.5, reason: trendOk ? "VIX elevated" : "Index below 200 DMA" };
}

// ---------------- selection ----------------
// candidates: [{ ticker, name, sector, score, ind }] already sorted by
// score descending. Returns core + buffer + why each rejection happened.

export function selectFromRanked(candidates, strategy, ctx = {}) {
  const picked = [], rejected = [], qualified = [];
  const want = strategy.basketSize || 7;
  const wantBuffer = strategy.bufferSize || 3;
  // How deep a name may sit and still be worth holding. Selling a stock
  // that slipped from 6th to 9th and buying it back next month costs a
  // full round trip for no change in exposure, so the hold list is
  // deliberately wider than the buy list.
  const holdDepth = strategy.holdRank || want * 2;

  for (const c of candidates) {
    if (strategy.minScore && c.score != null && c.score < strategy.minScore) {
      rejected.push({ ...c, why: `Score ${Math.round(c.score)} below ${strategy.minScore}` });
      continue;
    }
    if (!passesGates(c.ind || {}, strategy.gates, ctx)) {
      rejected.push({ ...c, why: gateReason(c.ind || {}, strategy.gates, ctx) });
      continue;
    }
    // Everything past the gates is holdable; the sector cap only limits
    // what we newly BUY, so a name already held is never force-sold by it.
    if (qualified.length < holdDepth) qualified.push(c);
    if (picked.length >= want + wantBuffer) continue;
    if (sectorCapped(picked, c.sector, strategy.gates)) {
      rejected.push({ ...c, why: `Already ${strategy.gates.maxPerSector} in ${c.sector}` });
      continue;
    }
    picked.push(c);
  }
  return {
    core: picked.slice(0, want),
    buffer: picked.slice(want, want + wantBuffer),
    qualified,
    holdable: new Set(qualified.map((c) => c.ticker)),
    rejected,
  };
}

function gateReason(ind, gates, ctx) {
  if (!gates) return "Filtered out";
  if (gates.maxRsi != null && ind.rsi != null && ind.rsi > gates.maxRsi) return `RSI ${Math.round(ind.rsi)} — already run`;
  if (gates.minOffHigh != null && ind.d52 != null && ind.d52 < gates.minOffHigh) return `At its 1-year high`;
  if (gates.aboveMa50 && ind.aboveMa50 === false) return "Below 50 DMA";
  if (gates.aboveMa200 && ind.aboveMa200 === false) return "Below 200 DMA";
  if (gates.macdPositive && ind.macdPositive === false) return "MACD negative";
  if (gates.relativeStrength && ind.ret3m != null && ctx?.medianRet3m != null && ind.ret3m <= ctx.medianRet3m) return "Lagging the universe";
  if (gates.minAdtv != null && ind.adtv != null && ind.adtv < gates.minAdtv) return `Thin — ₹${ind.adtv} Cr a day`;
  return "Filtered out";
}

// Consensus: count how many of the four base strategies picked each name.
export function consensusPicks(baseSelections, strategy) {
  const votes = new Map();
  for (const sel of baseSelections) {
    // A vote means "this strategy would hold the name", so it counts the
    // shortlist (basket + bench), not just the basket. Otherwise a stock
    // ranked 8th collects fewer votes purely because two of the four
    // strategies stop at 7 -- that is a difference in basket size, not a
    // disagreement about the stock, and it was quietly penalising exactly
    // the names the wider baskets were added to capture.
    for (const p of [...sel.core, ...(sel.buffer || [])]) {
      const cur = votes.get(p.ticker) || { ...p, votes: 0 };
      cur.votes++;
      votes.set(p.ticker, cur);
    }
  }
  const full = strategy.consensusFull ?? 4;
  const half = strategy.consensusHalf ?? 3;
  const ranked = [...votes.values()]
    .filter((v) => v.votes >= half)
    .sort((a, b) => b.votes - a.votes || (b.score ?? 0) - (a.score ?? 0));
  const core = [], buffer = [];
  for (const v of ranked) {
    v.convictionWeight = v.votes >= full ? 1 : 0.5;
    if (core.length < (strategy.basketSize || 7) && !sectorCapped(core, v.sector, strategy.gates)) core.push(v);
    else if (buffer.length < (strategy.bufferSize || 3)) buffer.push(v);
  }
  // Anything still carrying the minimum votes stays holdable, so a name
  // that slips from 4 votes to 3 is not sold and rebought.
  return { core, buffer, qualified: ranked, holdable: new Set(ranked.map((v) => v.ticker)), rejected: [] };
}

// ---------------- backtest ----------------
// Walks the history file day by day. Rebalances on a cadence, exits on
// stop / trailing stop / falling out of the basket, and charges every
// buy and sell through the charger so turnover is genuinely penalised.

export function buildBacktestContext(hist, benchmark = null) {
  const dates = hist.dates;
  const tickers = Object.keys(hist.tickers);
  const closesBy = {};
  for (const t of tickers) closesBy[t] = hist.tickers[t].close;

  // Indicators for every ticker on every day, computed once.
  const indBy = {};
  for (const t of tickers) {
    const cl = closesBy[t];
    const arr = new Array(dates.length).fill(null);
    for (let i = 0; i < dates.length; i++) arr[i] = indicatorsAt(cl, i);
    indBy[t] = arr;
  }

  // Equal-weight index built from the universe itself. The benchmark file
  // does not reach back far enough to cover the whole window, and this is
  // a truer read of "is our own universe trending" anyway.
  const idx = [];
  for (let i = 0; i < dates.length; i++) {
    let s = 0, n = 0;
    for (const t of tickers) {
      const a = closesBy[t][i - 1], b = closesBy[t][i];
      if (a != null && b != null && a > 0) { s += b / a - 1; n++; }
    }
    idx.push(i === 0 ? 100 : idx[i - 1] * (1 + (n ? s / n : 0)));
  }
  const idxAboveMa200 = idx.map((_, i) => {
    const m = sma(idx, i, 200);
    return m == null ? null : idx[i] > m;
  });

  // Median 3-month return across the universe each day -- the bar the
  // relative-strength gate has to clear.
  const medianRet3m = dates.map((_, i) => {
    const vals = [];
    for (const t of tickers) { const r = indBy[t][i]?.ret3m; if (r != null) vals.push(r); }
    if (!vals.length) return null;
    vals.sort((a, b) => a - b);
    return vals[Math.floor(vals.length / 2)];
  });

  // India VIX aligned to our dates, when the benchmark file carries it.
  const vixCloses = benchmark?.indices?.["^INDIAVIX"]?.closes || null;
  const vixSorted = vixCloses ? Object.keys(vixCloses).sort() : null;
  const vixOn = dates.map((d) => {
    if (!vixCloses) return null;
    if (vixCloses[d] != null) return vixCloses[d];
    let last = null;
    for (const k of vixSorted) { if (k <= d) last = vixCloses[k]; else break; }
    return last;
  });

  return { hist, dates, tickers, closesBy, indBy, idx, idxAboveMa200, medianRet3m, vixOn };
}

// One strategy over the whole window. rebalanceDays 21 ~= monthly.
export function backtest(ctx, strategy, charger = ZERO_CHARGER, opts = {}) {
  const { dates, tickers, closesBy, indBy, idxAboveMa200, medianRet3m, vixOn } = ctx;
  const startCapital = opts.capital ?? 50000;
  const rebalanceDays = opts.rebalanceDays ?? 21;
  const stopAtr = strategy.stopAtr ?? 2.5;
  const trailAtr = strategy.trailAtr ?? 3.0;
  const want = strategy.basketSize || 7;

  let cash = startCapital, charges = 0, lastRebal = -1;
  const holdings = new Map();   // ticker -> { units, entry, entryDate, peak, atr }
  const trades = [];
  const curve = [];
  let daysInvested = 0;

  // Consensus needs the four base selections each rebalance day.
  const baseDefs = strategy.consensus
    ? (DEFS?.strategies || []).filter((s) => !s.consensus)
    : null;

  for (let i = 0; i < dates.length; i++) {
    const date = dates[i];
    const dayCtx = { medianRet3m: medianRet3m[i], indexAboveMa200: idxAboveMa200[i], vix: vixOn[i] };
    const isRebal = lastRebal < 0 || (i - lastRebal) >= rebalanceDays;

    // --- exits: stop, trailing stop, or dropped out of the basket ---
    let selection = null;
    if (isRebal) {
      selection = selectForDay(ctx, strategy, i, dayCtx, baseDefs);
      lastRebal = i;
    }
    const inBasket = selection ? new Set(selection.core.map((p) => p.ticker)) : null;

    for (const [t, pos] of [...holdings]) {
      const px = closesBy[t][i];
      if (px == null) continue;
      if (px > pos.peak) pos.peak = px;
      let reason = null;
      const stopPx = pos.entry * (1 - (stopAtr * pos.atr) / 100);
      // The trail only arms once the position has actually made money —
      // an immediately-armed trail is just a tighter stop and cuts the
      // few large winners that pay for everything else.
      const armed = strategy.trailStop && pos.peak >= pos.entry * (1 + ((strategy.trailArmAtr ?? 1.5) * pos.atr) / 100);
      const trailPx = pos.peak * (1 - (trailAtr * pos.atr) / 100);
      if (px <= stopPx) reason = "SL";
      else if (armed && px <= trailPx) reason = "TRAIL";
      else if (isRebal && selection) {
        // Carry-over: keep a name while it still qualifies and sits inside
        // the hold list, even if it is no longer a top-7 buy. Without it,
        // a slip from 6th to 9th costs a full round trip for nothing.
        if (strategy.carryOver) { if (!selection.holdable.has(t)) reason = "DROPPED"; }
        else if (inBasket && !inBasket.has(t)) reason = "REBAL";
      }
      if (!reason) continue;
      const gross = pos.units * px;
      const fee = charger.sell(gross);
      cash += gross - fee; charges += fee;
      trades.push({ ticker: t, entryDate: pos.entryDate, exitDate: date, entry: pos.entry, exit: px,
        retPct: (px / pos.entry - 1) * 100, days: i - pos.dayIdx, reason });
      holdings.delete(t);
    }

    // --- entries ---
    if (isRebal && selection) {
      const { exposure } = regimeExposure(strategy, dayCtx);
      let nav = cash;
      for (const [t, p] of holdings) nav += p.units * (closesBy[t][i] ?? p.entry);
      const deployable = nav * exposure;
      // Inverse-ATR weights: a name that swings 5% a day gets a smaller
      // position than one that swings 2%, so every holding puts roughly
      // the same rupees at risk. Equal rupees would let the wildest name
      // dominate the month. Falls back to equal slots when disabled.
      const weightOf = invAtrWeights(selection.core, strategy);
      for (const pick of selection.core) {
        if (holdings.size >= want) break;
        if (holdings.has(pick.ticker)) continue;
        const px = closesBy[pick.ticker][i];
        if (px == null || px <= 0) continue;
        const convict = pick.convictionWeight ?? 1;
        const slot = deployable * weightOf(pick) * convict;
        const spend = Math.min(slot, Math.max(0, cash) * 0.98);
        if (spend < px) continue;                       // cannot afford one share
        const units = Math.floor(spend / px);
        if (units < 1) continue;
        const cost = units * px;
        const fee = charger.buy(cost);
        if (cost + fee > cash) continue;
        cash -= cost + fee; charges += fee;
        holdings.set(pick.ticker, {
          units, entry: px, entryDate: date, dayIdx: i, peak: px,
          atr: pick.ind?.atr ?? 3.5,
        });
      }
    }

    let value = cash;
    for (const [t, p] of holdings) value += p.units * (closesBy[t][i] ?? p.entry);
    if (holdings.size) daysInvested++;
    curve.push({ date, retPct: (value / startCapital - 1) * 100, holdings: holdings.size, cash });
  }

  return { curve, trades, charges, startCapital, daysInvested, strategy };
}

// Returns pick -> fraction of deployable capital. Inverse-ATR when the
// strategy asks for it, otherwise a flat 1/basketSize. A missing ATR
// falls back to the basket median so one data gap cannot hand a single
// name an enormous position.
export function invAtrWeights(picks, strategy) {
  const want = strategy.basketSize || 7;
  if (!picks?.length) return () => 1 / want;

  const atrs = picks.map((p) => p.ind?.atr).filter((v) => Number.isFinite(v) && v > 0).sort((a, b) => a - b);
  const med = atrs.length ? atrs[Math.floor(atrs.length / 2)] : 3.5;
  const atrOf = (p) => (Number.isFinite(p.ind?.atr) && p.ind.atr > 0 ? p.ind.atr : med);

  // Raw weight = risk parity (or a flat slot) times conviction.
  const raw = new Map();
  for (const p of picks) {
    const base = strategy.invAtrSizing ? 1 / atrOf(p) : 1;
    raw.set(p.ticker, base * (p.convictionWeight ?? 1));
  }
  const total = [...raw.values()].reduce((a, b) => a + b, 0);
  if (total <= 0) return () => 1 / want;

  let w = new Map([...raw].map(([k, v]) => [k, v / total]));

  // Cap any single name at twice its equal share. Risk parity alone will
  // happily put a third of the book into one unusually quiet stock (an
  // InvIT or a utility), which is a concentration risk arriving through
  // the back door of a risk-control rule. Excess is redistributed across
  // the uncapped names; two passes settles it.
  const cap = strategy.maxWeight ?? 2 / want;
  for (let pass = 0; pass < 3; pass++) {
    let excess = 0;
    const under = [];
    for (const [k, v] of w) {
      if (v > cap) { excess += v - cap; w.set(k, cap); }
      else under.push(k);
    }
    if (excess <= 1e-9 || !under.length) break;
    const room = under.reduce((a, k) => a + (cap - w.get(k)), 0);
    if (room <= 1e-9) break;
    for (const k of under) w.set(k, w.get(k) + excess * ((cap - w.get(k)) / room));
  }

  // A full basket is deployed completely, just tilted -- half-weighting a
  // name should shift money to the higher-conviction names, not park it in
  // cash. A SHORT basket keeps the shortfall in cash on purpose: if only
  // five names qualify, holding five and sitting on the rest is the point.
  const scale = picks.length >= want ? 1 : picks.length / want;
  return (p) => (w.get(p.ticker) || 0) * scale;
}

function selectForDay(ctx, strategy, i, dayCtx, baseDefs) {
  const { tickers, indBy, hist } = ctx;
  // Ranking on a smoothed score rather than today's reading. The raw
  // score reshuffles ~90% of its top 7 within a month, so ranking on a
  // single day's value means buying yesterday's spike and paying a round
  // trip to do it. Averaging over `scoreSmoothDays` ranks on the trend
  // in the score instead of its noise.
  const n = Math.max(1, strategy.scoreSmoothDays || 1);
  const ranked = [];
  for (const t of tickers) {
    const ind = indBy[t][i];
    if (!ind) continue;
    const score = n === 1 ? hist.tickers[t].score[i] : avgScore(hist.tickers[t].score, i, n);
    if (score == null) continue;
    ranked.push({ ticker: t, name: t, sector: hist.tickers[t].sector, score, ind });
  }
  ranked.sort((a, b) => b.score - a.score);
  if (strategy.consensus && baseDefs?.length) {
    const base = baseDefs.map((s) => selectFromRanked(ranked, s, dayCtx));
    return consensusPicks(base, strategy);
  }
  return selectFromRanked(ranked, strategy, dayCtx);
}

// ---------------- metrics ----------------

function monthKey(d) { return d.slice(0, 7); }

export function metrics(result, benchCurve = null) {
  const { curve, trades, charges, startCapital, daysInvested } = result;
  if (!curve?.length) return null;
  const last = curve[curve.length - 1].retPct;
  const years = Math.max(curve.length / 252, 1e-6);
  const cagr = (Math.pow(1 + last / 100, 1 / years) - 1) * 100;

  // Drawdown and how long it took to get back.
  let peak = -Infinity, maxDD = 0, peakIdx = 0, ddStart = 0, recoveryDays = null, inDD = false;
  curve.forEach((pt, i) => {
    const f = 1 + pt.retPct / 100;
    if (f >= peak) {
      if (inDD && recoveryDays == null) recoveryDays = i - ddStart;
      peak = f; peakIdx = i; inDD = false;
    } else {
      const dd = (f / peak - 1) * 100;
      if (dd < maxDD) { maxDD = dd; ddStart = peakIdx; inDD = true; }
    }
  });

  // Daily volatility -> annualised.
  const rets = [];
  for (let i = 1; i < curve.length; i++) {
    const a = 1 + curve[i - 1].retPct / 100, b = 1 + curve[i].retPct / 100;
    if (a > 0) rets.push(b / a - 1);
  }
  const mean = rets.length ? rets.reduce((x, y) => x + y, 0) / rets.length : 0;
  const varr = rets.length > 1 ? rets.reduce((x, y) => x + (y - mean) ** 2, 0) / (rets.length - 1) : 0;
  const vol = Math.sqrt(varr) * Math.sqrt(252) * 100;

  // Monthly returns, compounded off the curve.
  const byMonth = new Map();
  for (const pt of curve) {
    const k = monthKey(pt.date);
    if (!byMonth.has(k)) byMonth.set(k, { first: pt.retPct, last: pt.retPct });
    else byMonth.get(k).last = pt.retPct;
  }
  const months = [...byMonth.entries()].map(([m, v]) => ({
    month: m,
    retPct: ((1 + v.last / 100) / (1 + v.first / 100) - 1) * 100,
  }));
  const wins = months.filter((m) => m.retPct > 0).length;
  const best = months.reduce((a, b) => (!a || b.retPct > a.retPct ? b : a), null);
  const worst = months.reduce((a, b) => (!a || b.retPct < a.retPct ? b : a), null);

  // Trade stats.
  const closed = trades.filter((t) => t.retPct != null);
  const winners = closed.filter((t) => t.retPct > 0);
  const losers = closed.filter((t) => t.retPct <= 0);
  const grossWin = winners.reduce((a, t) => a + t.retPct, 0);
  const grossLoss = Math.abs(losers.reduce((a, t) => a + t.retPct, 0));
  const bestTrade = closed.reduce((a, b) => (!a || b.retPct > a.retPct ? b : a), null);
  const worstTrade = closed.reduce((a, b) => (!a || b.retPct < a.retPct ? b : a), null);

  const benchRet = benchCurve?.length ? benchCurve[benchCurve.length - 1].retPct : null;

  return {
    netReturn: last, cagr, maxDD, recoveryDays, vol,
    riskAdj: vol > 0 ? cagr / vol : null,
    months, monthlyWinRate: months.length ? (wins / months.length) * 100 : null,
    bestMonth: best, worstMonth: worst,
    trades: closed.length,
    winRate: closed.length ? (winners.length / closed.length) * 100 : null,
    avgWin: winners.length ? grossWin / winners.length : null,
    avgLoss: losers.length ? -grossLoss / losers.length : null,
    profitFactor: grossLoss > 0 ? grossWin / grossLoss : null,
    avgHoldDays: closed.length ? closed.reduce((a, t) => a + t.days, 0) / closed.length : null,
    bestTrade, worstTrade,
    charges, chargesPct: (charges / startCapital) * 100,
    exposurePct: curve.length ? (daysInvested / curve.length) * 100 : null,
    alpha: benchRet != null ? last - benchRet : null,
    benchReturn: benchRet,
    start: curve[0].date, end: curve[curve.length - 1].date, days: curve.length,
  };
}

// Run all five and return [{ strategy, result, metrics, gross }].
export function runAll(ctx, chargePrefs, opts = {}) {
  const charger = makeCharger(chargePrefs);
  const bench = buildUniverseCurve(ctx);
  return (DEFS?.strategies || []).map((s) => {
    const result = backtest(ctx, s, charger, opts);
    const gross = backtest(ctx, s, ZERO_CHARGER, opts);
    return {
      strategy: s,
      result,
      gross,
      metrics: metrics(result, bench),
      grossReturn: gross.curve.length ? gross.curve[gross.curve.length - 1].retPct : null,
    };
  });
}

// Equal-weight universe curve -- the "did you beat just owning it all"
// benchmark for the backtest window.
export function buildUniverseCurve(ctx) {
  return ctx.idx.map((v, i) => ({ date: ctx.dates[i], retPct: (v / ctx.idx[0] - 1) * 100 }));
}
