// Fundamentals scoring — rules taken verbatim from the client's scoring sheet.
// Each rule returns { points, max, status: 'pass'|'partial'|'fail'|'na'|'hard_fail',
//                     value, note }
// status 'hard_fail' = a red flag — still scored, but flagged visually.
// status 'na'        = data not available, this parameter is skipped from the total.

const NA = { points: 0, max: 0, status: "na", value: null, note: "Data not available" };

// Sector-aware N/A: produce a contextual note instead of generic "Data not available".
function naWithReason(c, ruleKey, max) {
  const broad = (c["Broad Sector"] || "").trim();
  const sec = (c["Sector"] || "").trim();
  const ind = (c["Broad Industry"] || "").trim();
  const isFinancial = broad === "Financial Services";
  const isInsurance = /insurance/i.test(ind) || /insurance/i.test(sec);
  const isBank = /bank/i.test(ind);
  const isREIT = /reit|invit|real estate investment/i.test(ind);
  let note = "Data not available";
  if (ruleKey === "de" && isFinancial) {
    note = "Not applicable — client framework excludes the financial sector from Debt/Equity scoring (banks, NBFCs, insurance and AMCs are highly leveraged by design).";
  } else if (ruleKey === "de") {
    note = "D/E not reported — likely negative net worth, or capital structure too small/atypical to report. Treat as a caution.";
  } else if (ruleKey === "icr" && (isFinancial || isInsurance || isBank)) {
    note = `Not applicable — ${isInsurance ? "insurance" : isBank ? "banking" : "financial-sector"} companies report interest as core revenue/expense, not as a coverage ratio.`;
  } else if (ruleKey === "icr") {
    note = "Interest Coverage not reported — usually means the company is debt-free or had no interest expense in the period (which is positive).";
  } else if (ruleKey === "cfo" && (isFinancial || isInsurance || isREIT)) {
    note = `Not directly comparable — ${isREIT ? "REITs/InvITs" : isInsurance ? "insurance companies" : "financial-sector companies"} report cash flows on a different framework.`;
  } else if (ruleKey === "cfo") {
    note = "Operating cash flow not reported yet (often the case for recently-listed companies before their first full annual report).";
  } else if (ruleKey === "ebitda" && (isFinancial || isBank)) {
    note = `Not applicable — ${isBank ? "banks" : "financial-sector companies"} don't report "operating margin" in the conventional sense.`;
  } else if (ruleKey === "npm" && (isFinancial || isBank)) {
    note = `Not applicable to ${isBank ? "banks" : "financial-sector companies"} as reported.`;
  } else if ((ruleKey === "rev3y" || ruleKey === "pat3y" || ruleKey === "eps") && c["Company"]) {
    note = "Insufficient history — likely a recently-listed company. Will populate once 3+ years of data are available.";
  }
  return { points: 0, max, status: "na", value: null, note };
}

// ---- helpers ----
const parsePercent = (v) => {
  if (v == null) return null;
  const s = String(v).replace(/%/g, "").replace(/,/g, "").trim();
  if (!s || s === "-") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
};
// Parse a pipe-delimited series like "23.4|24.1|24.8|26.2" into a number[]
const parseSeries = (v) => {
  if (!v) return [];
  return String(v).split("|").map((s) => parseFloat(s)).filter((n) => Number.isFinite(n));
};
const parseNumber = (v) => {
  if (v == null) return null;
  const s = String(v).replace(/[,\s]/g, "").replace(/Cr\.?/i, "").trim();
  if (!s || s === "-") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
};
const fmtPct = (n) => (n == null ? "—" : `${n}%`);

// ---- rules (16 active; 3 deferred) ----

function ruleROE(c) {
  const v = parsePercent(c["ROE"]);
  if (v == null) return naWithReason(c, "roe", 2);
  if (v > 12) return { points: 2, max: 2, status: "pass", value: fmtPct(v), note: "ROE > 12% — strong capital efficiency vs cost of equity." };
  if (v >= 8) return { points: 1, max: 2, status: "partial", value: fmtPct(v), note: "ROE 8–12% — borderline." };
  return { points: 0, max: 2, status: "fail", value: fmtPct(v), note: "ROE < 8% — weak capital efficiency." };
}

function ruleROCE(c) {
  const v = parsePercent(c["ROCE"]);
  if (v == null) return naWithReason(c, "roce", 2);
  if (v > 15) return { points: 2, max: 2, status: "pass", value: fmtPct(v), note: "ROCE > 15% — returns exceed cost of capital deployed." };
  if (v >= 10) return { points: 1, max: 2, status: "partial", value: fmtPct(v), note: "ROCE 10–15% — borderline." };
  return { points: 0, max: 2, status: "fail", value: fmtPct(v), note: "ROCE < 10% — capital not being deployed efficiently." };
}

function ruleEBITDAMargin(c) {
  // Per client framework: PASS if EBITDA margin improving YoY; FAIL if
  // contracting 2+ consecutive years. We use OPM (Operating Profit Margin)
  // as the EBITDA-margin proxy — Screener exposes OPM directly.
  // OPM Series is pipe-delimited oldest → newest, up to 12 years.
  const series = parseSeries(c["OPM Series"]);
  const last = parsePercent(c["OPM last year"]);
  const prev = parsePercent(c["OPM preceding year"]);
  if (series.length < 3) {
    // Not enough history for the multi-year rule — fall back to the 2-year
    // comparison and flag the limitation.
    if (last == null || prev == null) return naWithReason(c, "ebitda", 1);
    const val2 = `${fmtPct(prev)} → ${fmtPct(last)}`;
    if (last > prev) return { points: 1, max: 1, status: "pass", value: val2, note: "Operating margin improving YoY (only 2 years of history available)." };
    if (last >= prev - 0.5) return { points: 1, max: 1, status: "partial", value: val2, note: "Margin roughly stable (only 2 years of history)." };
    return { points: 0, max: 1, status: "fail", value: val2, note: "Operating margin contracting (only 2 years of history)." };
  }
  const recent = series.slice(-Math.min(series.length, 6));   // up to last 6 years
  const yrsShown = recent.length;
  const val = `OPM history (${yrsShown}y): ${recent.map(n => n.toFixed(0) + "%").join(" → ")}`;
  // Count the longest run of YoY contractions in the recent window
  let longestContraction = 0, currentRun = 0;
  for (let i = 1; i < recent.length; i++) {
    if (recent[i] < recent[i - 1]) { currentRun++; longestContraction = Math.max(longestContraction, currentRun); }
    else currentRun = 0;
  }
  const latestImproving = recent.at(-1) > recent.at(-2);
  if (longestContraction >= 2) {
    return { points: 0, max: 1, status: "fail", value: val, note: `OPM contracting ${longestContraction} consecutive years — fails per client framework.` };
  }
  if (latestImproving) {
    return { points: 1, max: 1, status: "pass", value: val, note: "Operating margin improving in latest year; no 2-year contraction streak in recent history." };
  }
  return { points: 1, max: 1, status: "partial", value: val, note: "Margin not in 2+ year contraction but latest year not improving — partial credit." };
}

function ruleNPM(c) {
  const last = parsePercent(c["NPM last year"]);
  const prev = parsePercent(c["NPM preceding year"]);
  if (last == null || prev == null) return naWithReason(c, "npm", 1);
  const val = `${fmtPct(prev)} → ${fmtPct(last)}`;
  const compressionBps = (prev - last) * 100;
  if (last >= prev) return { points: 1, max: 1, status: "pass", value: val, note: "Net profit margin stable or expanding." };
  if (compressionBps <= 200) return { points: 1, max: 1, status: "partial", value: val, note: "Slight margin compression (within 200 bps)." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Margin compressed > 200 bps in latest year." };
}

function ruleCFO(c) {
  // Per client framework: PASS if positive CFO for 10 consecutive years;
  // HARD FAIL if any year is negative. Series is oldest → newest.
  // Financial-sector exception (same rationale as D/E): banks/NBFCs consume
  // cash as they grow the loan book, so negative operating cash flow is
  // structural, not a red flag. Non-financials keep the strict test.
  if ((c["Broad Sector"] || "").trim() === "Financial Services") return naWithReason(c, "cfo", 2);
  const series = parseSeries(c["CF Operations Series"]);
  const ly = parseNumber(c["CF Operations LY"]);
  const py = parseNumber(c["CF Operations PY"]);
  if (series.length === 0 && ly == null && py == null) return naWithReason(c, "cfo", 2);
  if (series.length === 0) {
    // No history series — fall back to 2-year check with a note
    const val0 = `LY ${ly ?? "—"} | PY ${py ?? "—"}`;
    if (ly != null && py != null && ly > 0 && py > 0) {
      return { points: 2, max: 2, status: "pass", value: val0, note: "Both reported years positive (only 2 years of CFO history available — full 10-year check pending)." };
    }
    if (ly != null && ly < 0) {
      return { points: 0, max: 2, status: "hard_fail", value: val0, note: "Negative CFO in latest year — hard fail regardless of PAT." };
    }
    return { points: 0, max: 2, status: "fail", value: val0, note: "CFO not positive." };
  }
  const recent = series.slice(-10);    // up to last 10 years
  const yrsCovered = recent.length;
  const negativeYears = recent.filter((n) => n < 0).length;
  const latestNeg = recent.at(-1) < 0;
  const val = `${yrsCovered}y of CFO: ${recent.map((n) => Math.round(n).toLocaleString()).join(" → ")}`;
  if (latestNeg) {
    return { points: 0, max: 2, status: "hard_fail", value: val, note: `Negative CFO in latest year (₹${Math.round(recent.at(-1)).toLocaleString()} Cr) — hard fail regardless of PAT.` };
  }
  if (negativeYears === 0 && yrsCovered >= 10) {
    return { points: 2, max: 2, status: "pass", value: val, note: "Positive CFO for the full 10-year window — matches client's most stringent test." };
  }
  if (negativeYears === 0) {
    return { points: 2, max: 2, status: "pass", value: val, note: `Positive CFO across all ${yrsCovered} reported years (10-year window not yet available).` };
  }
  if (negativeYears <= 2) {
    return { points: 1, max: 2, status: "partial", value: val, note: `${negativeYears} of last ${yrsCovered} years had negative CFO — partial credit but watch closely.` };
  }
  return { points: 0, max: 2, status: "fail", value: val, note: `${negativeYears} of last ${yrsCovered} years had negative CFO — pattern of weak cash generation.` };
}

// Defensive sectors get a lower 3Y revenue-growth threshold per client
// framework. Typical Indian-market defensive classification — staples,
// pharma, utilities. Match against Screener's "Broad Industry" or "Sector".
const DEFENSIVE_SECTORS = new Set([
  // Consumer staples
  "Diversified FMCG", "Personal Products", "Food Products", "Beverages",
  "Cigarettes & Tobacco Products", "Agricultural Food & other Products",
  // Healthcare
  "Pharmaceuticals & Biotechnology", "Healthcare Services",
  "Healthcare Equipment & Supplies",
  // Utilities
  "Power", "Gas",
]);

function isDefensiveSector(c) {
  const ind = (c["Broad Industry"] || "").trim();
  const sec = (c["Sector"] || "").trim();
  return DEFENSIVE_SECTORS.has(ind) || DEFENSIVE_SECTORS.has(sec);
}

function ruleRevenueGrowth(c) {
  const v = parsePercent(c["Sales growth 3Years"]);
  if (v == null) return naWithReason(c, "rev3y", 2);
  const defensive = isDefensiveSector(c);
  const passT = defensive ? 8 : 12;       // defensives need only 8%+ to PASS
  const partialT = defensive ? 5 : 8;     // partial range tightens accordingly
  const sectorNote = defensive ? " (defensive sector — lower threshold applied per client framework)" : "";
  if (v >= passT) return { points: 2, max: 2, status: "pass", value: fmtPct(v), note: `3Y revenue CAGR ≥ ${passT}%${sectorNote}.` };
  if (v >= partialT) return { points: 1, max: 2, status: "partial", value: fmtPct(v), note: `3Y revenue CAGR ${partialT}–${passT}%${sectorNote}.` };
  return { points: 0, max: 2, status: "fail", value: fmtPct(v), note: `3Y revenue CAGR < ${partialT}%${sectorNote}.` };
}

function rulePATGrowth(c) {
  const v = parsePercent(c["Profit Var 3Yrs"]);
  if (v == null) return naWithReason(c, "pat3y", 2);
  if (v >= 15) return { points: 2, max: 2, status: "pass", value: fmtPct(v), note: "3Y PAT CAGR ≥ 15% — strongest profit signal." };
  if (v >= 8) return { points: 1, max: 2, status: "partial", value: fmtPct(v), note: "3Y PAT CAGR 8–15%." };
  return { points: 0, max: 2, status: "fail", value: fmtPct(v), note: "3Y PAT CAGR < 8%." };
}

function ruleEPSGrowth(c) {
  // Per client framework: PASS if EPS positive AND growing consistently
  // over last 3–4 quarters. Erratic EPS = 0 pts.
  // EPS Qtr Series is up to last 8 quarters, oldest → newest.
  const series = parseSeries(c["EPS Qtr Series"]);
  const v3 = parsePercent(c["EPS growth 3Years"]);
  const v5 = parsePercent(c["EPS growth 5Years"]);
  if (series.length < 4 && v3 == null && v5 == null) return naWithReason(c, "eps", 1);
  if (series.length < 4) {
    // Fallback to the CAGR-based check
    const val = `3Y ${fmtPct(v3)} | 5Y ${fmtPct(v5)}`;
    if ((v3 ?? -1) > 0 && (v5 ?? -1) > 0) return { points: 1, max: 1, status: "pass", value: val, note: "EPS CAGR positive over 3Y + 5Y (only annual data available — quarterly history pending)." };
    if ((v3 ?? -1) > 0 || (v5 ?? -1) > 0) return { points: 1, max: 1, status: "partial", value: val, note: "EPS CAGR positive in one window (only annual data available)." };
    return { points: 0, max: 1, status: "fail", value: val, note: "EPS CAGR not positive in either annual window." };
  }
  const recent = series.slice(-4);  // latest 4 quarters
  const val = `EPS last 4q: ${recent.map((n) => n.toFixed(1)).join(" → ")}`;
  const allPositive = recent.every((n) => n > 0);
  let monotonicUp = 0;     // count of strictly increasing transitions
  for (let i = 1; i < recent.length; i++) if (recent[i] > recent[i - 1]) monotonicUp++;
  // "Consistently growing" = at least 2 of 3 quarter-on-quarter increases AND all positive
  if (allPositive && monotonicUp >= 2) {
    return { points: 1, max: 1, status: "pass", value: val, note: `EPS positive and growing over last 4 quarters (${monotonicUp}/3 QoQ increases).` };
  }
  if (allPositive) {
    return { points: 1, max: 1, status: "partial", value: val, note: "EPS positive across all 4 quarters but growth not consistently up." };
  }
  return { points: 0, max: 1, status: "fail", value: val, note: "EPS erratic or negative in one or more recent quarters." };
}

function ruleQuarterlyEarnings(c) {
  const q = [
    parseNumber(c["Net Profit Qtr (-3)"]),
    parseNumber(c["Net Profit Qtr (-2)"]),
    parseNumber(c["Net Profit Qtr (-1)"]),
    parseNumber(c["Net Profit Qtr (latest)"]),
  ];
  if (q.some((x) => x == null)) return naWithReason(c, "qoq", 1);
  const val = q.join(" → ");
  let declines = 0, maxConsecDeclines = 0;
  for (let i = 1; i < q.length; i++) {
    if (q[i] < q[i - 1]) { declines++; maxConsecDeclines = Math.max(maxConsecDeclines, declines); } else declines = 0;
  }
  if (q[3] > q[2] && q[2] > q[1]) return { points: 1, max: 1, status: "pass", value: val, note: "QoQ earnings improving last 3 quarters." };
  if (maxConsecDeclines >= 2) return { points: 0, max: 1, status: "fail", value: val, note: "Two or more consecutive QoQ declines." };
  return { points: 1, max: 1, status: "partial", value: val, note: "Mixed but no sustained decline." };
}

function ruleDebtEquity(c) {
  const v = parseNumber(c["Debt to equity"]);
  const broadSector = (c["Broad Sector"] || "").trim();
  // Apply client's sector exception: financial sector excluded from D/E scoring.
  if (broadSector === "Financial Services") return naWithReason(c, "de", 2);
  if (v == null) return naWithReason(c, "de", 2);
  if (v < 0.5) return { points: 2, max: 2, status: "pass", value: v.toString(), note: "D/E < 0.5 — comfortably low leverage." };
  if (v <= 1.0) return { points: 1, max: 2, status: "partial", value: v.toString(), note: "D/E 0.5–1.0 — moderate leverage." };
  if (v <= 2.0) return { points: 0, max: 2, status: "fail", value: v.toString(), note: "D/E > 1.0 — elevated leverage." };
  return { points: 0, max: 2, status: "hard_fail", value: v.toString(), note: "D/E > 2 — extreme leverage, hard fail per client framework (stock excluded from SPIP basket)." };
}

function ruleInterestCoverage(c) {
  const v = parseNumber(c["Int Coverage"]);
  // Financial-sector exception (same rationale as D/E): interest is a bank /
  // NBFC's raw material, so interest-coverage isn't a meaningful debt-service
  // test — excluding them stops ~36 blue-chip financials being red-flagged.
  if ((c["Broad Sector"] || "").trim() === "Financial Services") return naWithReason(c, "icr", 2);
  if (v == null) return naWithReason(c, "icr", 2);
  if (v > 3) return { points: 2, max: 2, status: "pass", value: v.toString(), note: "Interest coverage > 3 — debt comfortably serviced." };
  if (v >= 1.5) return { points: 1, max: 2, status: "partial", value: v.toString(), note: "Coverage 1.5–3 — tight." };
  return { points: 0, max: 2, status: "hard_fail", value: v.toString(), note: "Coverage < 1.5 — debt serviceability risk." };
}

function ruleCurrentRatio(c) {
  const v = parseNumber(c["Current ratio"]);
  if (v == null) return naWithReason(c, "cr", 1);
  if (v > 1.2) return { points: 1, max: 1, status: "pass", value: v.toString(), note: "Current ratio > 1.2." };
  if (v >= 1.0) return { points: 1, max: 1, status: "partial", value: v.toString(), note: "Current ratio 1.0–1.2." };
  return { points: 0, max: 1, status: "fail", value: v.toString(), note: "Current ratio < 1.0 — liquidity stress." };
}

function rulePledge(c) {
  const v = parsePercent(c["Pledged percentage"]);
  if (v == null) return naWithReason(c, "pledge", 2);
  // Rising-pledge red flag: client spec says rising trend = downgrade.
  // We only have a multi-quarter series when Screener exposed the pledge
  // sub-row (i.e. promoters have actually pledged). For 0% companies the
  // series is empty and the trend check is a no-op.
  const series = parseSeries(c["Pledge Series"]);
  let trendNote = "";
  let rising = false;
  if (series.length >= 2) {
    const first = series[0];
    const last = series[series.length - 1];
    if (last > first + 0.5) {
      rising = true;
      trendNote = ` Rising trend ${first}% → ${last}% across ${series.length} quarters — flagged per client framework.`;
    } else if (last < first - 0.5) {
      trendNote = ` Trend declining ${first}% → ${last}%.`;
    } else {
      trendNote = ` Trend roughly stable.`;
    }
  }
  if (v < 5) return { points: 2, max: 2, status: "pass", value: fmtPct(v), note: "Promoter pledge < 5%." + trendNote };
  if (v <= 20) {
    if (rising) return { points: 0, max: 2, status: "fail", value: fmtPct(v), note: `Promoter pledge ${fmtPct(v)} and rising — red flag per client framework.` + trendNote };
    return { points: 1, max: 2, status: "partial", value: fmtPct(v), note: "Promoter pledge 5–20%." + trendNote };
  }
  return { points: 0, max: 2, status: "hard_fail", value: fmtPct(v), note: "Promoter pledge > 20% — hard fail." + trendNote };
}

function rulePromoterHolding(c) {
  const v = parsePercent(c["Change in Prom Hold"]);
  if (v == null) return naWithReason(c, "ph", 1);
  // Always render the change with an explicit sign so the cell value
  // can't be misread as the absolute holding level. "0%" alone looks
  // like "no promoter holding"; "+0.00%" reads clearly as "no change".
  const signed = (v > 0 ? "+" : "") + fmtPct(v);
  const val = `Change ${signed} (vs previous quarter)`;
  if (v >= 0) return { points: 1, max: 1, status: "pass", value: val, note: "Promoter holding stable or increasing — no dilution this quarter." };
  if (v > -2) return { points: 1, max: 1, status: "partial", value: val, note: "Minor reduction in promoter holding (under 2 % in the quarter)." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Promoter holding fell by more than 2 % in the quarter — fails per client framework." };
}

function ruleFIIDII(c) {
  // Per client framework: PASS if FII + DII holding trending up over last
  // 4 quarters; 1 pt. Sharp FII exit = flag. PR E uses the new quarterly
  // history series; falls back to the single-quarter change if series is
  // too short (recent IPOs, mergers).
  const fiiSer = parseSeries(c["FII Holding Series"]);
  const diiSer = parseSeries(c["DII Holding Series"]);
  if (fiiSer.length >= 4 && diiSer.length >= 4) {
    const f = fiiSer.slice(-4);
    const d = diiSer.slice(-4);
    const combined = f.map((v, i) => v + d[i]);
    const trendUp = combined[3] > combined[0];                // 4-quarter net direction
    let upTransitions = 0;
    for (let i = 1; i < combined.length; i++) if (combined[i] > combined[i - 1]) upTransitions++;
    const sharpFIIExit = (f[3] - f[0]) < -2;                  // FII alone fell >2 pp across the 4 qtrs
    const val = `FII ${f.join(" → ")}  |  DII ${d.join(" → ")}`;
    if (trendUp && !sharpFIIExit && upTransitions >= 2) {
      return { points: 1, max: 1, status: "pass", value: val, note: "FII + DII combined holding trending up over last 4 quarters." };
    }
    if (trendUp && sharpFIIExit) {
      return { points: 1, max: 1, status: "partial", value: val, note: "Combined up but FII fell sharply (>2 pp over 4 quarters) — caution flag." };
    }
    if (trendUp) {
      return { points: 1, max: 1, status: "partial", value: val, note: "Net up over 4 quarters but inconsistent QoQ direction." };
    }
    return { points: 0, max: 1, status: "fail", value: val, note: "Institutional holding not trending up over last 4 quarters." };
  }
  // Fallback to the single-quarter change ribbon value
  const fii = parsePercent(c["Chg in FII Hold"]);
  const dii = parsePercent(c["Chg in DII Hold"]);
  if (fii == null && dii == null) return naWithReason(c, "fiidii", 1);
  const sum = (fii ?? 0) + (dii ?? 0);
  const val = `FII ${fmtPct(fii)} | DII ${fmtPct(dii)} (single-quarter change — 4-quarter history unavailable)`;
  if (sum > 0) {
    if (fii != null && fii < -2) return { points: 1, max: 1, status: "partial", value: val, note: "Combined positive but FII exiting sharply (>2%)." };
    return { points: 1, max: 1, status: "pass", value: val, note: "FII + DII holding trending up (single-quarter snapshot)." };
  }
  return { points: 0, max: 1, status: "fail", value: val, note: "Institutional holding not increasing." };
}

function ruleDividendConsistency(c) {
  // Per client framework: PASS if dividend paid in at least 3 of last 5 years.
  // PR E uses the new annual Dividend Series; counts years where payout > 0.
  const series = parseSeries(c["Dividend Series"]);
  if (series.length >= 5) {
    const last5 = series.slice(-5);
    const paidYears = last5.filter((v) => v > 0).length;
    const val = `Dividend payout % last 5 years: ${last5.map((n) => Math.round(n) + "%").join(" → ")}`;
    if (paidYears >= 4) return { points: 1, max: 1, status: "pass", value: val, note: `Paid in ${paidYears} of last 5 years — comfortably consistent.` };
    if (paidYears >= 3) return { points: 1, max: 1, status: "pass", value: val, note: "Paid in 3 of last 5 years — meets client threshold." };
    if (paidYears >= 1) return { points: 0, max: 1, status: "fail", value: val, note: `Paid only in ${paidYears} of last 5 years — erratic.` };
    return { points: 0, max: 1, status: "fail", value: val, note: "No dividend in any of last 5 years." };
  }
  // Fallback to ribbon values when history is too short
  const ly = parseNumber(c["Dividend last year"]);
  const py = parseNumber(c["Dividend Prev Ann"]);
  const d5 = parsePercent(c["Div 5Yrs"]);
  if (ly == null && py == null && d5 == null) return naWithReason(c, "div", 1);
  const val = `LY ${ly ?? "—"} | PY ${py ?? "—"} | 5Y avg ${fmtPct(d5)} (annual history too short)`;
  const lastTwoPaid = (ly ?? 0) > 0 && (py ?? 0) > 0;
  const fiveYrYield = d5 ?? 0;
  if (lastTwoPaid && fiveYrYield > 0) return { points: 1, max: 1, status: "pass", value: val, note: "Dividend paid consistently (ribbon-only proxy — full 5-yr history not available)." };
  if ((ly ?? 0) > 0 || fiveYrYield > 0) return { points: 1, max: 1, status: "partial", value: val, note: "Some dividend history." };
  return { points: 0, max: 1, status: "fail", value: val, note: "No / erratic dividend." };
}

function ruleInsiderBuying(c) {
  // The dashboard merges insider data from public/data/insider-trades.json
  // into each company under c.insider_*. If the merge didn't find a match,
  // those fields are absent → N/A with explanation.
  if (!c.insider_loaded) {
    return { ...NA, max: 1, note: "Insider trades data not loaded — NSE PIT feed unavailable. Will populate on next refresh." };
  }
  if (c.insider_transactions == null || c.insider_transactions === 0) {
    return { ...NA, max: 1, note: "No insider trades disclosed in the last 6 months for this company." };
  }
  const netSh = c.insider_net_shares ?? 0;
  const netVal = c.insider_net_value ?? 0;
  const buySh  = c.insider_buy_shares ?? 0;
  const sellSh = c.insider_sell_shares ?? 0;
  // Compute sell % of total shares using market cap / price as a rough proxy
  const mcap = parseNumber((c["Market Cap"] || "").replace(/Cr\.?/i, "").replace(/,/g, ""));
  const cmp = parseNumber((c["Current Price"] || "").replace(/,/g, ""));
  const totalShares = (mcap != null && cmp != null && cmp > 0) ? (mcap * 1e7) / cmp : null;
  const sellPctOfFloat = totalShares ? (sellSh / totalShares) * 100 : null;
  const sellHardFail = sellPctOfFloat != null && sellPctOfFloat > 1;

  const pledges = c.insider_pledges_excluded || 0;
  const pledgeNote = pledges > 0 ? ` · ${pledges} pledge${pledges===1?"":"s"} excluded` : "";
  const val = `Buy ${formatShares(buySh)} | Sell ${formatShares(sellSh)} | Net ${(netSh>=0?"+":"")}${formatShares(netSh)} (${c.insider_transactions} trade${c.insider_transactions===1?"":"s"}${pledgeNote})`;
  if (sellHardFail) {
    return { points: 0, max: 1, status: "fail", value: val, note: `Insider selling exceeded 1% of float (${sellPctOfFloat.toFixed(2)}%) — red flag per client framework.` };
  }
  // Primary score signal is the rupee net (netVal); but NSE PIT
  // disclosures frequently come without execution prices, in which
  // case netVal == 0 even though netSh is clearly positive (or
  // negative). Fall back to net SHARES so a stock with unambiguous
  // buying like 53.91 L shares net bought / 0 sold doesn't get
  // misclassified as "balanced".
  const signal = netVal !== 0 ? netVal : netSh;
  if (signal > 0) return { points: 1, max: 1, status: "pass", value: val, note: "Net insider buying over the last 6 months — positive signal." };
  if (signal === 0) return { points: 1, max: 1, status: "partial", value: val, note: "Insider buys and sells offset each other — no directional signal." };
  return { points: 0, max: 1, status: "fail", value: val, note: "Net insider selling over the last 6 months." };
}

function formatShares(n) {
  if (n == null || !Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  if (abs >= 1e7) return (n / 1e7).toFixed(2) + " Cr";
  if (abs >= 1e5) return (n / 1e5).toFixed(2) + " L";
  if (abs >= 1e3) return (n / 1e3).toFixed(1) + " K";
  return n.toString();
}

// ---- governance: SEBI orders + active proceedings ----
function ruleGovernanceIssues(c) {
  // c.governance_flag is set by app.js merging governance-flags.json
  // when the loaded file's flagged_companies map contains this ticker.
  if (!c?.governance_loaded) return naWithReason(c, "governance", 2);
  const flag = c?.governance_flag;
  if (!flag) {
    return { points: 2, max: 2, status: "pass", value: "No active SEBI proceedings",
      note: "No SEBI enforcement orders or press-release mentions naming this company as an active noticee." };
  }
  // Active proceeding flagged → hard fail per client spec.
  const typeLabel = flag.order_type ? ` (${flag.order_type})` : "";
  const snippet = flag.snippet ? ` — "${String(flag.snippet).slice(0, 120)}"` : "";
  return {
    points: 0, max: 2, status: "hard_fail",
    value: `Active SEBI proceeding${typeLabel}`,
    note: `${flag.primary_name || c.Company} flagged in SEBI ${flag.source_label || "enforcement listing"}${snippet}. Hard fail per client framework.`,
  };
}

// ---- auditor remarks: weekly routine extraction from BSE annual reports ----
// Reads each company's latest auditor opinion from public/data/auditor-opinions.json.
// That file is maintained by a weekly Claude.ai routine that processes the
// auto-generated public/data/auditor-todo.csv, extracts the auditor's report
// section from each company's latest BSE annual report PDF, and the
// auditor-merger workflow merges its output back into auditor-opinions.json.

// Substantive-detail suffix appended to the rule's note: auditor firm,
// report year, opinion date, Emphasis of Matter text, key concerns —
// drawn from the routine-extracted fields. Returns "" when no useful
// detail is present (e.g. legacy entries without routine enrichment).
function buildAuditorDetail(c) {
  const trunc = (s, n) => (s && s.length > n) ? s.slice(0, n).trimEnd() + "..." : s;
  const isEmpty = (s) => !s || /^(none|n\/?a|could not extract)$/i.test(String(s).trim());
  const meta = [];
  if (c.auditor_report_year) meta.push(c.auditor_report_year);
  if (c.auditor_firm) meta.push(`audited by ${c.auditor_firm}`);
  if (c.auditor_opinion_date && !/could not extract/i.test(c.auditor_opinion_date)) {
    meta.push(`opinion dated ${c.auditor_opinion_date}`);
  }
  const parts = [];
  if (meta.length) parts.push(meta.join(", "));
  if (!isEmpty(c.auditor_emphasis_of_matter)) {
    parts.push(`Emphasis of Matter: ${trunc(c.auditor_emphasis_of_matter, 280)}`);
  }
  if (!isEmpty(c.auditor_key_concerns)) {
    parts.push(`Key concerns: ${trunc(c.auditor_key_concerns, 280)}`);
  }
  // Strip a trailing period/ellipsis from each part so the joiner doesn't double-punctuate.
  return parts.length ? ` ${parts.map((p) => p.replace(/[.…]+\s*$/, "")).join(". ")}.` : "";
}

function ruleAuditorOpinion(c) {
  if (!c?.auditor_opinions_loaded) {
    return naWithReason(c, "auditorRemarks", 2);
  }
  const op = c.auditor_opinion;
  if (!op) {
    return { points: 0, max: 2, status: "na", value: "—",
      note: "Latest auditor opinion not yet extracted." };
  }
  const opLower = String(op).toLowerCase();
  // "Not disclosed" / "Not provided" / similar non-answers → N/A.
  if (/\bnot\s+(disclosed|provided|available|stated|known|reported|specified|mentioned)\b|^n\/?a$|^unknown$|^none$/.test(opLower)) {
    return { points: 0, max: 2, status: "na", value: op,
      note: "Auditor opinion not disclosed in the latest annual report." };
  }
  const detail = buildAuditorDetail(c);
  // Adverse opinion or disclaimer → hard fail per client framework — BUT only
  // when it's a genuine auditor verdict. The extraction agent DEFAULTS to
  // "Disclaimer of Opinion" whenever it can't find or read a company's annual
  // report, which wrongly red-flagged ~110 blue-chips (Colgate, P&G, ICICI AMC…).
  // If the source shows the extraction failed, treat it as N/A, not a red flag.
  if (/\b(adverse|disclaimer)\b/.test(opLower)) {
    const src = String(c.auditor_opinion_source || "").toLowerCase();
    const extractionFailed = /no annual report|annual report.{0,20}not (?:provided|identified|available)|not provided|no successful information|no auditor|could ?n.?.?t extract|no reference|not available|no company details|no information|not specified|no details|not disclosed|no data|not identified|no company name/.test(src);
    if (extractionFailed) {
      return { points: 0, max: 2, status: "na", value: op,
        note: "Auditor opinion couldn't be extracted (no annual report found for this company) — not counted as a red flag." };
    }
    return { points: 0, max: 2, status: "hard_fail", value: op,
      note: `${op} per the most recent annual report. Hard fail per client framework.${detail}` };
  }
  // Qualified opinion or emphasis-of-matter → 1 pt partial.
  if (/\bqualified\b|emphasis[- ]of[- ]matter|emphasis matter/.test(opLower) && !/un\s*-?\s*qualified/.test(opLower)) {
    return { points: 1, max: 2, status: "partial", value: op,
      note: `${op} flagged in the most recent annual report — 1 pt per client framework (qualification or emphasis-of-matter).${detail}` };
  }
  // Unqualified opinion → full 2 pts.
  if (/\bunqualified\b|\bclean\b/.test(opLower)) {
    return { points: 2, max: 2, status: "pass", value: op,
      note: `Clean ${op.toLowerCase().includes("unqualified") ? "unqualified opinion" : op} from the most recent annual report.${detail}` };
  }
  // Anything else: keep as N/A so we don't accidentally penalise a
  // misclassified-but-clean opinion.
  return { points: 0, max: 2, status: "na", value: op,
    note: `Auditor opinion text ("${op}") didn't match a known classification — review manually.` };
}

// ---- deferred parameters (data source pending) ----
const DEFERRED = [];

// ---- master ----
const ACTIVE_RULES = [
  { key: "roe", label: "ROE", category: "Profitability", criteria: "> 12%", fn: ruleROE },
  { key: "roce", label: "ROCE", category: "Profitability", criteria: "> 15%", fn: ruleROCE },
  { key: "ebitda", label: "EBITDA Margin (YoY)", category: "Profitability", criteria: "Improving", fn: ruleEBITDAMargin },
  { key: "npm", label: "Net Profit Margin", category: "Profitability", criteria: "Stable / expanding", fn: ruleNPM },
  { key: "cfo", label: "Operating Cash Flow", category: "Profitability", criteria: "Positive", fn: ruleCFO },
  { key: "rev3y", label: "Revenue Growth (3Y)", category: "Growth", criteria: "≥ 12–15%", fn: ruleRevenueGrowth },
  { key: "pat3y", label: "PAT Growth (3Y)", category: "Growth", criteria: "≥ 15%", fn: rulePATGrowth },
  { key: "eps", label: "EPS Growth", category: "Growth", criteria: "Positive & consistent", fn: ruleEPSGrowth },
  { key: "qoq", label: "Quarterly Earnings", category: "Growth", criteria: "QoQ improving 3 qtrs", fn: ruleQuarterlyEarnings },
  { key: "de", label: "Debt / Equity", category: "Balance Sheet", criteria: "< 0.5", fn: ruleDebtEquity },
  { key: "icr", label: "Interest Coverage", category: "Balance Sheet", criteria: "> 3", fn: ruleInterestCoverage },
  { key: "cr", label: "Current Ratio", category: "Balance Sheet", criteria: "> 1.2", fn: ruleCurrentRatio },
  { key: "pledge", label: "Promoter Pledge", category: "Balance Sheet", criteria: "< 5%", fn: rulePledge },
  { key: "ph", label: "Change in Promoter Holding", category: "Shareholding", criteria: "Stable or increasing (Δ ≥ 0% QoQ)", fn: rulePromoterHolding },
  { key: "fiidii", label: "FII + DII Holding", category: "Shareholding", criteria: "Increasing", fn: ruleFIIDII },
  { key: "insider", label: "Insider Buying", category: "Shareholding", criteria: "Net buying in last 6 mo", fn: ruleInsiderBuying },
  { key: "div", label: "Dividend Consistency", category: "Governance", criteria: "Positive", fn: ruleDividendConsistency },
  { key: "governance", label: "Corporate Governance Issues", category: "Governance", criteria: "No active SEBI proceedings", fn: ruleGovernanceIssues },
  { key: "auditorRemarks", label: "Auditor Remarks", category: "Governance", criteria: "Clean unqualified opinion", fn: ruleAuditorOpinion },
];

export function scoreCompany(company) {
  const breakdown = ACTIVE_RULES.map((r) => ({ ...r, ...r.fn(company) }));
  const totalPoints = breakdown.reduce((s, b) => s + b.points, 0);
  const totalMax = breakdown.reduce((s, b) => s + b.max, 0);
  const hardFails = breakdown.filter((b) => b.status === "hard_fail").map((b) => b.label);
  const naCount = breakdown.filter((b) => b.status === "na").length;
  return {
    company,
    breakdown,
    deferred: DEFERRED,
    totalPoints,
    totalMax,
    scorePct: totalMax ? Math.round((totalPoints / totalMax) * 100) : 0,
    hardFails,
    naCount,
  };
}

// Compact fundamental-ratio record for the Custom Lab's parameter picker
// (Fundamentals category). Parses the screener strings ("115 %", "0.00",
// "34.2") into numbers so a strategy can filter on ROE ≥, P/E ≤, etc.,
// live and — once snapshots log it — historically. null when no company.
export function fundVals(fc) {
  if (!fc) return null;
  const num = (v) => {
    if (v == null) return null;
    const n = parseFloat(String(v).replace(/[%,\s]/g, ""));
    return Number.isFinite(n) ? n : null;
  };
  return {
    roe:    num(fc.ROE),
    roce:   num(fc.ROCE),
    de:     num(fc["Debt to equity"]),
    pe:     num(fc["Stock P/E"]),
    salesG: num(fc["Sales growth 3Years"]),
    patG:   num(fc["Profit Var 3Yrs"]),
  };
}

export { ACTIVE_RULES, DEFERRED };
