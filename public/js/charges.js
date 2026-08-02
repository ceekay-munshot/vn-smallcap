// ============================================================
// TRANSACTION CHARGES — NSE equity delivery, Zerodha rate card.
//
// The old model was a single "% per side" applied symmetrically. Real
// charges are not symmetric and are not all percentages:
//
//   - STT 0.10%          both sides
//   - Stamp duty 0.015%  BUY side only
//   - NSE txn 0.00307%   both sides
//   - SEBI Rs 10/crore   both sides
//   - GST 18%            on (brokerage + txn + SEBI)
//   - Brokerage          Rs 0 on delivery
//   - DP charge Rs 15.34 FLAT per scrip, SELL side only, incl GST
//
// The DP fee is the one that matters. Being flat, its percentage cost
// depends entirely on position size: Rs 15.34 is 0.43% of a Rs 3,571
// position and 0.11% of a Rs 14,286 one. A model that ignores it
// understates the cost of small tickets and of high turnover, because
// it is charged per SELL -- holding a name across a rebalance is free.
//
// Verified against the Zerodha calculator: Rs 5,000 buy = Rs 5.94,
// Rs 5,000 sell = Rs 20.53, round trip Rs 26.47.
// ============================================================

export const CHARGE_DEFAULTS = {
  brokeragePct: 0,        // % per side — Zerodha delivery is free
  sttPct: 0.10,           // % — both sides
  stampPct: 0.015,        // % — buy side only
  exchangePct: 0.00307,   // % — both sides (NSE)
  sebiPct: 0.0001,        // % — both sides (Rs 10 per crore)
  gstPct: 18,             // % — on brokerage + exchange + SEBI
  dpFee: 15.34,           // Rs flat per scrip sold, incl GST
};

// Editable fields for the settings panel. Order = display order.
export const CHARGE_FIELDS = [
  { key: "brokeragePct", label: "Brokerage", suffix: "%/side", step: 0.01, note: "0 on delivery" },
  { key: "sttPct",       label: "STT",       suffix: "%/side", step: 0.005, note: "both sides" },
  { key: "stampPct",     label: "Stamp duty", suffix: "%",     step: 0.005, note: "buy only" },
  { key: "exchangePct",  label: "NSE txn",   suffix: "%/side", step: 0.001, note: "both sides" },
  { key: "sebiPct",      label: "SEBI",      suffix: "%/side", step: 0.0001, note: "Rs 10/crore" },
  { key: "gstPct",       label: "GST",       suffix: "%",      step: 1, note: "on brokerage + txn + SEBI" },
  { key: "dpFee",        label: "DP charge", suffix: "Rs",     step: 1, note: "flat, per scrip sold" },
];

const pct = (v) => (Number(v) || 0) / 100;

// Percentage component of a BUY, as a fraction of trade value.
export function buyRate(p = CHARGE_DEFAULTS) {
  const brok = pct(p.brokeragePct), stt = pct(p.sttPct), stamp = pct(p.stampPct);
  const exch = pct(p.exchangePct), sebi = pct(p.sebiPct), gst = pct(p.gstPct);
  return brok + stt + stamp + exch + sebi + (brok + exch + sebi) * gst;
}

// Percentage component of a SELL. The flat DP fee is deliberately NOT
// folded in here — it does not scale with notional, so callers must add
// it per position via sellCharge().
export function sellRate(p = CHARGE_DEFAULTS) {
  const brok = pct(p.brokeragePct), stt = pct(p.sttPct);
  const exch = pct(p.exchangePct), sebi = pct(p.sebiPct), gst = pct(p.gstPct);
  return brok + stt + exch + sebi + (brok + exch + sebi) * gst;
}

export function buyCharge(notional, p = CHARGE_DEFAULTS) {
  return Math.max(0, notional) * buyRate(p);
}

// One scrip sold = one DP fee, regardless of quantity.
export function sellCharge(notional, p = CHARGE_DEFAULTS) {
  const n = Math.max(0, notional);
  return n * sellRate(p) + (n > 0 ? (Number(p.dpFee) || 0) : 0);
}

// Full round trip on a position of this size, in rupees and as a % of it.
export function roundTrip(notional, p = CHARGE_DEFAULTS) {
  const n = Math.max(0, notional);
  const rupees = buyCharge(n, p) + sellCharge(n, p);
  return { rupees, pct: n > 0 ? (rupees / n) * 100 : 0 };
}

// What a full basket rotation costs, annualised at a given frequency.
// rotationsPerYear 12 = replace every holding every month.
export function annualDrag(perPositionNotional, rotationsPerYear, p = CHARGE_DEFAULTS) {
  return roundTrip(perPositionNotional, p).pct * rotationsPerYear;
}

// ---- Charger objects — what the backtest engines consume -------------
// An engine calls charger.buy(notional) / charger.sell(notional) and gets
// rupees back. Passing ZERO_CHARGER produces the frictionless twin curve
// used to show how much the charges actually cost.

export function makeCharger(p = CHARGE_DEFAULTS) {
  return {
    zero: false,
    buy: (notional) => buyCharge(notional, p),
    sell: (notional) => sellCharge(notional, p),
    prefs: p,
  };
}

export const ZERO_CHARGER = {
  zero: true,
  buy: () => 0,
  sell: () => 0,
  prefs: null,
};

// Human-readable one-liner for the settings panel: what a round trip
// costs at the user's actual ticket size.
export function chargeSummary(capital, basketSize, p = CHARGE_DEFAULTS) {
  const per = basketSize > 0 ? capital / basketSize : 0;
  const rt = roundTrip(per, p);
  return {
    perPosition: per,
    roundTripPct: rt.pct,
    roundTripRs: rt.rupees,
    monthlyRotationAnnualPct: rt.pct * 12,
    dpSharePct: rt.rupees > 0 ? ((Number(p.dpFee) || 0) / rt.rupees) * 100 : 0,
  };
}
