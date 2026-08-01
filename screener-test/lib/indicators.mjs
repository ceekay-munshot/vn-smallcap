// Technical indicators computed from a daily OHLC series. Every function
// returns a full-length array aligned to the input (null during the warm-up
// window) so a backtest can read each indicator's value on any given day.
// Pure functions, no I/O — unit-tested in verify-indicators-history.mjs.

export function sma(values, period) {
  const out = new Array(values.length).fill(null);
  let sum = 0, count = 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v == null) { out[i] = null; continue; }
    sum += v; count++;
    if (i >= period) { const old = values[i - period]; if (old != null) { sum -= old; count--; } }
    if (count >= period) out[i] = sum / period;
  }
  return out;
}

export function ema(values, period) {
  const out = new Array(values.length).fill(null);
  const k = 2 / (period + 1);
  let prev = null;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v == null) { out[i] = prev; continue; }
    if (prev == null) {
      // seed with SMA once we have `period` values
      if (i >= period - 1) {
        let s = 0, ok = true;
        for (let j = i - period + 1; j <= i; j++) { if (values[j] == null) { ok = false; break; } s += values[j]; }
        if (ok) { prev = s / period; out[i] = prev; }
      }
    } else {
      prev = v * k + prev * (1 - k);
      out[i] = prev;
    }
  }
  return out;
}

// Wilder's RSI(14).
export function rsi(closes, period = 14) {
  const out = new Array(closes.length).fill(null);
  let avgGain = 0, avgLoss = 0;
  for (let i = 1; i < closes.length; i++) {
    if (closes[i] == null || closes[i - 1] == null) continue;
    const ch = closes[i] - closes[i - 1];
    const gain = Math.max(0, ch), loss = Math.max(0, -ch);
    if (i <= period) {
      avgGain += gain / period; avgLoss += loss / period;
      if (i === period) out[i] = rsiFrom(avgGain, avgLoss);
    } else {
      avgGain = (avgGain * (period - 1) + gain) / period;
      avgLoss = (avgLoss * (period - 1) + loss) / period;
      out[i] = rsiFrom(avgGain, avgLoss);
    }
  }
  return out;
}
function rsiFrom(avgGain, avgLoss) {
  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - 100 / (1 + rs);
}

// Rate of change over `period` days, in percent.
export function roc(closes, period) {
  const out = new Array(closes.length).fill(null);
  for (let i = period; i < closes.length; i++) {
    if (closes[i] == null || closes[i - period] == null || closes[i - period] === 0) continue;
    out[i] = (closes[i] / closes[i - period] - 1) * 100;
  }
  return out;
}

// MACD: {macd, signal, hist} series (EMA12 − EMA26, signal EMA9 of macd).
export function macd(closes, fast = 12, slow = 26, sig = 9) {
  const emaFast = ema(closes, fast), emaSlow = ema(closes, slow);
  const line = closes.map((_, i) => (emaFast[i] != null && emaSlow[i] != null) ? emaFast[i] - emaSlow[i] : null);
  const signal = ema(line, sig);
  const hist = line.map((v, i) => (v != null && signal[i] != null) ? v - signal[i] : null);
  return { macd: line, signal, hist };
}

// Wilder's ATR(14) and ADX(14) (trend strength).
export function adx(highs, lows, closes, period = 14) {
  const n = closes.length;
  const tr = new Array(n).fill(null), plusDM = new Array(n).fill(null), minusDM = new Array(n).fill(null);
  for (let i = 1; i < n; i++) {
    if ([highs[i], lows[i], closes[i - 1], highs[i - 1], lows[i - 1]].some((v) => v == null)) continue;
    tr[i] = Math.max(highs[i] - lows[i], Math.abs(highs[i] - closes[i - 1]), Math.abs(lows[i] - closes[i - 1]));
    const up = highs[i] - highs[i - 1], down = lows[i - 1] - lows[i];
    plusDM[i] = (up > down && up > 0) ? up : 0;
    minusDM[i] = (down > up && down > 0) ? down : 0;
  }
  const out = new Array(n).fill(null);
  let atr = null, pDI = null, mDI = null, adxVal = null, dxCount = 0, dxSum = 0;
  for (let i = 1; i < n; i++) {
    if (tr[i] == null) continue;
    if (atr == null) {
      // seed after `period` valid TRs
      let s = 0, sp = 0, sm = 0, ok = 0;
      for (let j = i - period + 1; j <= i; j++) { if (tr[j] == null) { ok = -1; break; } s += tr[j]; sp += plusDM[j]; sm += minusDM[j]; ok++; }
      if (ok < period) continue;
      atr = s / period; pDI = sp / period; mDI = sm / period;
    } else {
      atr = (atr * (period - 1) + tr[i]) / period;
      pDI = (pDI * (period - 1) + plusDM[i]) / period;
      mDI = (mDI * (period - 1) + minusDM[i]) / period;
    }
    if (atr === 0) continue;
    const plusDI = 100 * pDI / atr, minusDI = 100 * mDI / atr;
    const dx = (plusDI + minusDI === 0) ? 0 : 100 * Math.abs(plusDI - minusDI) / (plusDI + minusDI);
    dxCount++; dxSum += dx;
    if (dxCount < period) continue;
    if (adxVal == null) adxVal = dxSum / period;
    else adxVal = (adxVal * (period - 1) + dx) / period;
    out[i] = adxVal;
  }
  return out;
}

// Rolling 52-week (252-day) high proximity: 1.0 = at the high, lower = below.
export function highProximity(closes, window = 252) {
  const out = new Array(closes.length).fill(null);
  for (let i = 0; i < closes.length; i++) {
    if (closes[i] == null) continue;
    let hi = -Infinity;
    for (let j = Math.max(0, i - window + 1); j <= i; j++) if (closes[j] != null && closes[j] > hi) hi = closes[j];
    if (hi > 0) out[i] = closes[i] / hi;
  }
  return out;
}

// Transparent 0–100 technical score for one day from its indicator values.
// Trend (price vs 50/200 MAs) + momentum (RSI zone, MACD, ROC) + trend
// strength (ADX) + being near the 52-week high. Mirrors the desk's language
// (RSI bands, moving-average stacking) so the backtest is legible.
export function technicalScore(d) {
  let pts = 0, max = 0;
  const add = (cond, w) => { max += w; if (cond) pts += w; };
  const part = (val, w) => { max += w; if (val != null) pts += w * Math.max(0, Math.min(1, val)); };

  // Trend structure (35)
  add(d.close != null && d.sma50 != null && d.close > d.sma50, 12);
  add(d.close != null && d.sma200 != null && d.close > d.sma200, 12);
  add(d.sma50 != null && d.sma200 != null && d.sma50 > d.sma200, 11);
  // Momentum (35)
  if (d.rsi != null) {
    max += 15;
    if (d.rsi >= 50 && d.rsi <= 70) pts += 15;            // healthy bullish zone
    else if (d.rsi > 70 && d.rsi <= 80) pts += 8;          // strong but extended
    else if (d.rsi >= 45 && d.rsi < 50) pts += 8;
    else if (d.rsi > 80) pts += 3;                         // overbought
  }
  add(d.macdHist != null && d.macdHist > 0, 10);
  add(d.roc20 != null && d.roc20 > 0, 10);
  // Trend strength (15) — ADX scaled 20→40
  part(d.adx != null ? (d.adx - 20) / 20 : null, 15);
  // Near 52-week high (15)
  part(d.highProx != null ? (d.highProx - 0.7) / 0.3 : null, 15);

  return max > 0 ? Math.round((pts / max) * 1000) / 10 : null;
}
