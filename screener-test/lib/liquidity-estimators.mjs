// Liquidity estimators from daily OHLCV — used to compute Bid-Ask Spread
// and Impact Cost without needing NSE Level-1 quote data.
//
// References:
//   - Abdi & Ranaldo (2017) — close/high/low spread proxy.
//     RFS 30(12):4437–4480. Closed form:
//       S = 2 * sqrt(mean[(c_t - eta_t) * (c_t - eta_{t+1})])
//     where c_t = ln(close), eta_t = (ln(H_t) + ln(L_t)) / 2.
//     Best of the classic low-frequency estimators (~0.74 monthly cross-
//     sectional correlation with TAQ effective spread).
//   - Amihud (2002) — daily illiquidity ratio.
//     ILLIQ_t = |return_t| / rupee_volume_t
//     Goyenko-Holden-Trzcinka (JFE 2009) found Amihud to be the best
//     low-frequency proxy for price impact (vs. spread).
//
// IMPORTANT: these are statistical *estimates* derived from end-of-day
// data. They produce reliable cross-sectional rankings (~0.7-0.9
// correlation with true values) but the absolute levels are noisy,
// especially for the most liquid large caps. Notes in the consuming
// rules make this caveat explicit.

// Abdi-Ranaldo (CHL) spread estimator. Returns spread as % of mid-price
// (e.g. 0.23 means 0.23%). Inputs: array of bar objects with .high,
// .low, .close — oldest to newest. Returns null if not enough data.
//
// Corporate-action guard: pairs of consecutive days where the close-to-
// close return exceeds ±15% are skipped. Indian circuit-breaker rules
// cap normal intraday moves at 10% (20% for index stocks on extreme
// days), so anything outside ±15% is almost certainly a dividend, stock
// split, demerger, or bonus issue — those create overnight close gaps
// that aren't real spread. Without this guard the estimator misreads
// the gap as a wide spread, producing values like 5%+ on otherwise
// highly-liquid stocks (e.g. Vedanta during its 2024 capital reduction).
export function abdiRanaldoSpreadPct(bars, windowSize = 30) {
  if (!Array.isArray(bars) || bars.length < 5) return null;
  const slice = bars.slice(-Math.max(windowSize + 1, 6));
  const products = [];
  const CORP_ACTION_RETURN_THRESHOLD = 0.15;
  for (let t = 0; t < slice.length - 1; t++) {
    const b1 = slice[t];
    const b2 = slice[t + 1];
    if (!b1 || !b2) continue;
    const h1 = Number(b1.high), l1 = Number(b1.low), c1 = Number(b1.close);
    const h2 = Number(b2.high), l2 = Number(b2.low), c2 = Number(b2.close);
    if (!(h1 > 0) || !(l1 > 0) || !(c1 > 0) || !(h2 > 0) || !(l2 > 0) || !(c2 > 0)) continue;
    // Skip pairs that straddle a corporate action.
    const c2c = Math.abs((c2 - c1) / c1);
    if (c2c > CORP_ACTION_RETURN_THRESHOLD) continue;
    const lnC1 = Math.log(c1);
    const eta1 = (Math.log(h1) + Math.log(l1)) / 2;
    const eta2 = (Math.log(h2) + Math.log(l2)) / 2;
    products.push((lnC1 - eta1) * (lnC1 - eta2));
  }
  if (products.length < 5) return null;
  const meanProd = products.reduce((s, x) => s + x, 0) / products.length;
  // Standard practice: clamp negatives to zero (spread can't be < 0;
  // negative estimates arise from noise on very liquid stocks).
  if (meanProd <= 0) return 0;
  const spreadProportion = 2 * Math.sqrt(meanProd);
  // Return as % of mid-price, rounded to 3 decimal places.
  return Math.round(spreadProportion * 100 * 1000) / 1000;
}

// Amihud illiquidity ratio expressed as the expected % price impact of
// an order of size `orderSizeRupees`. ILLIQ = mean(|return| / rupee_volume)
// across the window; impact% = ILLIQ * orderSizeRupees * 100. Returns
// null when data is too sparse or volumes are zero. Default order size
// is ₹5 crore — the size where the client's three-band scoring
// (≤0.3 pass / 0.3-0.5 partial / >0.5 fail) actually discriminates
// across the Nifty 500 universe; ₹1 cr is too small (98% pass).
export function amihudImpactPct(bars, windowSize = 30, orderSizeRupees = 5e7) {
  if (!Array.isArray(bars) || bars.length < 6) return null;
  const slice = bars.slice(-Math.max(windowSize + 1, 6));
  const ratios = [];
  for (let t = 1; t < slice.length; t++) {
    const cPrev = Number(slice[t - 1].close);
    const cNow  = Number(slice[t].close);
    const vol   = Number(slice[t].volume);
    if (!(cPrev > 0) || !(cNow > 0) || !(vol > 0)) continue;
    const r = (cNow - cPrev) / cPrev;
    // Same corporate-action guard as the spread estimator — a 30%
    // dividend-driven overnight gap shouldn't be misread as a real
    // price move that consumed liquidity.
    if (Math.abs(r) > 0.15) continue;
    const rupeeVolume = cNow * vol;
    if (rupeeVolume <= 0) continue;
    ratios.push(Math.abs(r) / rupeeVolume);
  }
  if (ratios.length < 5) return null;
  const meanRatio = ratios.reduce((s, x) => s + x, 0) / ratios.length;
  const impactPct = meanRatio * orderSizeRupees * 100;
  return Math.round(impactPct * 1000) / 1000;
}

// Liquidity tier from 20-day ADTV in ₹ crores. Calibrated against NSE's
// own published bands: Group I = mean impact cost ≤ 1% (roughly ADTV
// above ~₹5 cr), Group II = > 1%. Nifty-50-grade liquidity sits at
// ADTV ≥ ~₹100 cr.
export function liquidityTier(adtv_cr) {
  if (adtv_cr == null || !Number.isFinite(adtv_cr)) return null;
  if (adtv_cr >= 100) return "highly_liquid";   // Nifty 50 / large caps
  if (adtv_cr >= 25)  return "liquid";          // Solid Group I
  if (adtv_cr >= 5)   return "moderate";        // Group I borderline
  return "illiquid";                             // Group II
}
