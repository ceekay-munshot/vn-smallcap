// Macro scoring — rules taken verbatim from the client's scoring sheet.
// Each rule takes a company that has had c._macro merged in (the contents of
// public/data/macro.json) plus per-company convenience flags (c.in_pli,
// c.in_renewable). The Macro tab uses the same {points, max, status, value,
// note} shape as the other scoring modules.

const NA = { points: 0, max: 0, status: "na", value: null, note: "Data not available" };

function inTheme(c, themeKey) {
  const themes = c?._macro?.sector_themes || {};
  const list = themes[themeKey] || [];
  const ind = (c?.industry || c?.broadIndustry || c?.["Broad Industry"] || "").trim();
  const sec = (c?.sector || c?.["Sector"] || "").trim();
  return list.some((s) => s === ind || s === sec);
}

function getSector(c) {
  return (c?.industry || c?.broadIndustry || c?.["Broad Industry"] || c?.sector || c?.["Sector"] || "").trim();
}
function naIfNoSector(c, max) {
  if (getSector(c)) return null;
  return { points: 0, max, status: "na", value: null,
    note: "Sector data not yet populated for this company — Fundamentals workflow needs to run once after sector scraping was added." };
}

function fmtList(arr, max = 3) {
  if (!arr || !arr.length) return "—";
  const head = arr.slice(0, max).join(", ");
  return arr.length > max ? `${head} (+${arr.length - max} more)` : head;
}

// ----- Per-company annual-report extraction helpers ----------------------------
// company-revenue-mix.json is populated by the daily Claude.ai routine that
// reads each company's BSE annual report MD&A. When data exists for the
// company, the affected rules use per-company truth instead of the sector
// proxy. Companies still awaiting extraction fall through to the original
// sector-proxy notes — the user-facing dashboard reveals nothing about
// proxies or pipeline state.
function rmix(c) { return c?._revenue_mix || null; }
function truncEvidence(s, n = 220) {
  if (!s) return "";
  const str = String(s).trim();
  return str.length > n ? str.slice(0, n).trimEnd() + "..." : str;
}

// ---- Sector Rotation (7 pts) ----

function ruleInfraPush(c) {
  const m = c?._macro;
  if (!m) return { ...NA, max: 2 };
  const na = naIfNoSector(c, 2); if (na) return na;
  const active = m.regime?.gov_capex_active;
  const sector = getSector(c);

  // PRIMARY: per-company govt-capex revenue share from annual report.
  const rm = rmix(c);
  if (rm?.govt_capex?.revenue_pct != null) {
    const pct = rm.govt_capex.revenue_pct;
    const ev = truncEvidence(rm.govt_capex.evidence);
    if (!active) return { points: 0, max: 2, status: "fail", value: sector, note: "Government capex push not flagged as active in macro context." };
    if (pct >= 30) return { points: 2, max: 2, status: "pass", value: sector,
      note: `${pct}% of revenue from government / capex projects — material benefit from active capex push.${ev ? ` ${ev}` : ""}` };
    if (pct >= 10) return { points: 1, max: 2, status: "partial", value: sector,
      note: `${pct}% revenue from government / capex — partial exposure.${ev ? ` ${ev}` : ""}` };
    return { points: 1, max: 2, status: "partial", value: sector,
      note: `Only ${pct}% revenue from government / capex — not materially capex-driven.${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: original sector proxy.
  const isBenef = inTheme(c, "infra_push");
  if (!isBenef) return { points: 1, max: 2, status: "partial", value: sector, note: `Sector not on the infra-beneficiary list — 1 pt neutral per client framework.` };
  if (!active) return { points: 0, max: 2, status: "fail", value: sector, note: "Government capex push not flagged as active in macro context." };
  return { points: 2, max: 2, status: "pass", value: sector, note: m.regime?.gov_capex_note || "In a sector benefiting from active government capex push." };
}

function ruleRateCuts(c) {
  const m = c?._macro;
  if (!m) return { ...NA, max: 2 };
  const na = naIfNoSector(c, 2); if (na) return na;
  const inCycle = m.regime?.rate_cut_cycle;
  const sector = getSector(c);

  // PRIMARY: per-company floating-rate liability share from annual report.
  const rm = rmix(c);
  if (rm?.rate_sensitive?.floating_rate_liability_pct != null) {
    const pct = rm.rate_sensitive.floating_rate_liability_pct;
    const ev = truncEvidence(rm.rate_sensitive.evidence);
    if (pct < 15) return { points: 1, max: 2, status: "partial", value: sector,
      note: `Only ${pct}% of liabilities at floating rate — limited sensitivity to rate cycle.${ev ? ` ${ev}` : ""}` };
    if (!inCycle) return { points: 1, max: 2, status: "partial", value: sector,
      note: `${pct}% floating-rate liabilities but RBI not currently in rate-cut cycle.${ev ? ` ${ev}` : ""}` };
    if (pct >= 40) return { points: 2, max: 2, status: "pass", value: sector,
      note: `${pct}% of liabilities at floating rate — directly benefits from RBI rate-cut cycle.${ev ? ` ${ev}` : ""}` };
    return { points: 1, max: 2, status: "partial", value: sector,
      note: `${pct}% floating-rate liabilities — moderate benefit from rate-cut cycle.${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: original sector proxy.
  const isBenef = inTheme(c, "rate_sensitive");
  if (!isBenef) return { points: 1, max: 2, status: "partial", value: sector, note: `Sector not rate-sensitive — 1 pt neutral per client framework.` };
  if (!inCycle) return { points: 1, max: 2, status: "partial", value: sector, note: "RBI not in rate-cut cycle — neutral for rate-sensitive sectors." };
  return { points: 2, max: 2, status: "pass", value: sector, note: m.regime?.rate_cut_cycle_note || "Rate-sensitive sector benefiting from RBI rate-cut cycle." };
}

function ruleChinaPlusOne(c) {
  const m = c?._macro;
  if (!m) return { ...NA, max: 2 };
  const na = naIfNoSector(c, 2); if (na) return na;
  const active = m.regime?.china_plus_one_active;
  const sector = getSector(c);

  // PRIMARY: per-company truth from annual report MD&A (silently upgrades
  // the rule note when extraction data exists; otherwise falls through to
  // the original sector-proxy notes — the dashboard reveals nothing about
  // pipeline state).
  const rm = rmix(c);
  if (rm?.china_plus_one) {
    const exp = rm.geography?.export_pct ?? null;
    const strength = rm.china_plus_one.strength;
    const ev = truncEvidence(rm.china_plus_one.evidence);
    if (!active) {
      return { points: 0, max: 2, status: "fail", value: sector, note: "China+1 theme not flagged as active — even curated-list companies score 0." };
    }
    if (strength === "strong" && exp != null && exp >= 40) {
      return { points: 2, max: 2, status: "pass", value: sector,
        note: `Strong China+1 exposure with ${exp}% export revenue.${ev ? ` ${ev}` : ""}` };
    }
    if (strength === "moderate" || (strength === "strong" && exp != null && exp >= 25)) {
      return { points: 1, max: 2, status: "partial", value: sector,
        note: `Moderate China+1 exposure${exp != null ? ` (${exp}% export revenue)` : ""}.${ev ? ` ${ev}` : ""}` };
    }
    if (strength === "weak") {
      return { points: 0, max: 2, status: "fail", value: sector,
        note: `Weak China+1 narrative${exp != null ? ` (${exp}% export revenue)` : ""}.${ev ? ` ${ev}` : ""}` };
    }
    return { points: 0, max: 2, status: "fail", value: sector,
      note: `No material China+1 exposure${exp != null ? ` (${exp}% export revenue)` : ""}.${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: original sector / curated-list proxy notes.
  const onCompanyList = !!c?.in_china_plus_one;
  const inSectorList = inTheme(c, "china_plus_one");
  if (onCompanyList) {
    if (!active) return { points: 0, max: 2, status: "fail", value: sector, note: "China+1 theme not flagged as active — even curated-list companies score 0." };
    return { points: 2, max: 2, status: "pass", value: sector,
      note: `${m.regime?.china_plus_one_note || "Benefiting from China+1 reorientation."} On curated company list with material China+1 / EMS revenue.` };
  }
  if (inSectorList) {
    if (!active) return { points: 0, max: 2, status: "fail", value: sector, note: "China+1 theme not flagged as active — sector match alone scores 0." };
    return { points: 1, max: 2, status: "partial", value: sector,
      note: `Sector benefits from China+1 theme but company not on the curated >15%-revenue list — partial credit.` };
  }
  return { points: 1, max: 2, status: "partial", value: sector, note: `Neither on the China+1 curated company list nor in a benefiting sector — 1 pt neutral per client framework.` };
}

function ruleRuralRecovery(c) {
  const m = c?._macro;
  if (!m) return { ...NA, max: 1 };
  const na = naIfNoSector(c, 1); if (na) return na;
  const signal = m.regime?.rural_recovery_signal || "neutral";
  const sector = getSector(c);

  // PRIMARY: per-company rural revenue share from annual report.
  const rm = rmix(c);
  if (rm?.rural?.revenue_pct != null) {
    const pct = rm.rural.revenue_pct;
    const ev = truncEvidence(rm.rural.evidence);
    if (pct < 15) return { points: 1, max: 1, status: "pass", value: sector,
      note: `Only ${pct}% of revenue from rural markets — not materially exposed.${ev ? ` ${ev}` : ""}` };
    if (signal === "good") return { points: 1, max: 1, status: "pass", value: sector,
      note: `${pct}% of revenue from rural markets — favourable rural indicators.${ev ? ` ${ev}` : ""}` };
    if (signal === "neutral") return { points: 1, max: 1, status: "partial", value: sector,
      note: `${pct}% revenue from rural markets but indicators are mixed.${ev ? ` ${ev}` : ""}` };
    return { points: 0, max: 1, status: "fail", value: sector,
      note: `${pct}% revenue from rural markets — rural signal weak / drought risk flagged.${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: original sector proxy.
  const isBenef = inTheme(c, "rural_recovery");
  if (!isBenef) return { points: 1, max: 1, status: "pass", value: sector, note: `Sector not rural-facing — 1 pt neutral per client framework.` };
  if (signal === "good") return { points: 1, max: 1, status: "pass", value: sector, note: m.regime?.rural_recovery_note || "Rural indicators improving — favourable for FMCG/Tractors/Agro." };
  if (signal === "neutral") return { points: 1, max: 1, status: "partial", value: sector, note: "Rural-facing sector but indicators are mixed." };
  return { points: 0, max: 1, status: "fail", value: sector, note: "Rural signal weak / drought risk flagged." };
}

// ---- Economic Indicators (7 pts) ----

function ruleGDPGrowth(c) {
  const m = c?._macro;
  if (!m?.economic) return { ...NA, max: 2 };
  const v = m.economic.gdp_yoy;
  const up = m.economic.gdp_trending_up;
  const val = `GDP YoY ${v}%${up ? " (trending up)" : " (not trending up)"}`;
  if (v >= 6.5 && up) return { points: 2, max: 2, status: "pass", value: val, note: "GDP ≥ 6.5% and trending up — broad market re-rating tailwind." };
  if (v >= 6.0) return { points: 1, max: 2, status: "partial", value: val, note: "GDP between 6.0–6.5% — broad-market caution." };
  return { points: 1, max: 2, status: "partial", value: val, note: "GDP below 6% — broad market caution (per client framework: 1 pt caution flag)." };
}

function ruleInflation(c) {
  const m = c?._macro;
  if (!m?.economic) return { ...NA, max: 1 };
  const v = m.economic.cpi;
  const inBand = m.economic.cpi_within_rbi_band;
  const val = `CPI ${v}% (RBI target 4% ± 2%)`;
  if (inBand && v <= 5.5) return { points: 1, max: 1, status: "pass", value: val, note: "CPI within RBI band and stable — margin-friendly." };
  if (inBand) return { points: 1, max: 1, status: "partial", value: val, note: "CPI within band but elevated — margin pressure for margin-sensitive cos." };
  return { points: 0, max: 1, status: "fail", value: val, note: "CPI outside RBI band — compresses margins for input-cost-sensitive sectors." };
}

function ruleCrudeOil(c) {
  const m = c?._macro;
  if (!m?.live?.crude_brent?.latest) return { ...NA, max: 1, note: "Live crude price not yet fetched — run the Macro workflow once." };
  const na = naIfNoSector(c, 1); if (na) return na;
  const crude = m.live.crude_brent.latest;
  const trend = m.live.crude_brent.trend;
  const sector = getSector(c);
  const val = `Brent $${crude}/bbl (${trend})`;

  // PRIMARY: per-company crude-as-raw-material cost share from annual report.
  const rm = rmix(c);
  if (rm?.crude_exposure?.raw_material_pct != null) {
    const pct = rm.crude_exposure.raw_material_pct;
    const ev = truncEvidence(rm.crude_exposure.evidence);
    // Low crude cost → not materially crude-sensitive
    if (pct < 10) return { points: 1, max: 1, status: "pass", value: val, note: `Only ${pct}% of revenue from crude-linked raw materials — limited crude sensitivity.${ev ? ` ${ev}` : ""}` };
    // Material crude exposure: cheap crude helps, expensive crude hurts
    if (crude < 85 && pct >= 15) return { points: 1, max: 1, status: "pass", value: val, note: `${pct}% raw-material crude exposure — soft crude price is a margin tailwind.${ev ? ` ${ev}` : ""}` };
    if (crude > 90 && pct >= 15) return { points: 0, max: 1, status: "fail", value: val, note: `${pct}% raw-material crude exposure — elevated crude is a margin headwind.${ev ? ` ${ev}` : ""}` };
    return { points: 1, max: 1, status: "partial", value: val, note: `${pct}% raw-material crude exposure — neutral at current crude levels.${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: original sector proxy.
  const benefitsLow = inTheme(c, "crude_low_beneficiary");
  const hurtsHigh = inTheme(c, "crude_high_hurt");
  if (crude < 85 && benefitsLow) return { points: 1, max: 1, status: "pass", value: val, note: `Crude below $85 — favourable for ${sector}.` };
  if (crude > 90 && hurtsHigh) return { points: 0, max: 1, status: "fail", value: val, note: `Crude above $90 — input-cost headwind for ${sector}.` };
  return { points: 1, max: 1, status: "pass", value: val, note: `Sector not directly crude-sensitive — 1 pt neutral per client framework.` };
}

function ruleINRUSD(c) {
  const m = c?._macro;
  if (!m?.live?.usdinr?.latest) return { ...NA, max: 1, note: "Live USD/INR not yet fetched — run the Macro workflow once." };
  const na = naIfNoSector(c, 1); if (na) return na;
  const rate = m.live.usdinr.latest;
  const trend = m.live.usdinr.trend;
  const inrWeakening = trend === "rising";
  const inrStrengthening = trend === "falling";
  const sector = getSector(c);
  const val = `USD/INR ₹${rate} (INR ${inrWeakening ? "weakening" : inrStrengthening ? "strengthening" : "stable"})`;

  // PRIMARY: per-company USD revenue share from annual report.
  const rm = rmix(c);
  if (rm?.inr_exposure?.usd_revenue_pct != null) {
    const pct = rm.inr_exposure.usd_revenue_pct;
    const hedge = rm.inr_exposure.hedging_policy;
    const ev = truncEvidence(rm.inr_exposure.evidence);
    const hedgeNote = hedge ? ` Hedging: ${truncEvidence(hedge, 120)}` : "";
    if (pct < 10) return { points: 1, max: 1, status: "pass", value: val, note: `Only ${pct}% of revenue in USD — limited INR sensitivity.${ev ? ` ${ev}` : ""}` };
    if (inrWeakening && pct >= 25) return { points: 1, max: 1, status: "pass", value: val,
      note: `${pct}% USD revenue — INR weakening adds margin tailwind.${hedgeNote}${ev ? ` ${ev}` : ""}` };
    if (inrStrengthening && pct >= 25) return { points: 0, max: 1, status: "fail", value: val,
      note: `${pct}% USD revenue — INR strengthening compresses export margins.${hedgeNote}${ev ? ` ${ev}` : ""}` };
    return { points: 1, max: 1, status: "partial", value: val,
      note: `${pct}% USD revenue — neutral at current INR trend.${hedgeNote}${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: original sector proxy.
  const exporter = inTheme(c, "inr_weakening_benefit");
  const importer = inTheme(c, "inr_weakening_hurt");
  if (inrWeakening && exporter) return { points: 1, max: 1, status: "pass", value: val, note: `INR weakening trend — adds ~70 bps margin tailwind for ${sector} exporters.` };
  if (inrStrengthening && exporter) return { points: 0, max: 1, status: "fail", value: val, note: `Strong INR hurts ${sector} exporters.` };
  if (inrWeakening && importer) return { points: 0, max: 1, status: "fail", value: val, note: `INR weakening hurts ${sector} importers.` };
  return { points: 1, max: 1, status: "pass", value: val, note: `Sector not directly INR-sensitive — 1 pt neutral per client framework.` };
}

function ruleBondYields(c) {
  const m = c?._macro;
  if (!m?.economic) return { ...NA, max: 1 };
  const na = naIfNoSector(c, 1); if (na) return na;
  const y = m.economic.gsec10y;
  const trend = m.economic.gsec10y_trend;
  const src = m.economic.gsec10y_source || "macro-context.json";
  const sector = getSector(c);
  const benefitsFall = inTheme(c, "bond_falling_benefit");
  const val = `10Y G-Sec ${y}% (${trend})`;
  const srcNote = src.toLowerCase().includes("macro-context") ? "" : ` Source: ${src}.`;
  if (trend === "declining" && benefitsFall) return { points: 1, max: 1, status: "pass", value: val, note: "Falling 10Y yield — spread compression eases for Banks/NBFCs/Realty." + srcNote };
  if (trend === "rising" && benefitsFall) return { points: 0, max: 1, status: "fail", value: val, note: "Rising yields compress NIMs for rate-sensitive financials." + srcNote };
  // Non-bond-yield-sensitive sector: 1 pt neutral per client framework.
  return { points: 1, max: 1, status: "pass", value: val, note: `Sector not directly bond-yield-sensitive — 1 pt neutral per client framework.` };
}

// ---- Government Policy (4 pts) ----

function rulePLIBeneficiary(c) {
  const m = c?._macro;
  if (!m?.pli_companies) return { ...NA, max: 2 };

  // PRIMARY: explicit PLI mention + revenue share from annual report.
  const rm = rmix(c);
  if (rm?.pli_scheme?.is_beneficiary != null) {
    const isBenef = rm.pli_scheme.is_beneficiary;
    const pct = rm.pli_scheme.revenue_pct;
    const ev = truncEvidence(rm.pli_scheme.evidence);
    if (!isBenef) return { points: 0, max: 2, status: "fail", value: "Not a PLI beneficiary",
      note: `Annual report does not cite PLI-scheme participation.${ev ? ` ${ev}` : ""}` };
    if (pct != null && pct >= 5) return { points: 2, max: 2, status: "pass", value: `${pct}% PLI revenue`,
      note: `${pct}% of revenue from PLI-eligible products.${ev ? ` ${ev}` : ""}` };
    return { points: 2, max: 2, status: "pass", value: "On PLI scheme",
      note: `Confirmed PLI beneficiary in annual report.${ev ? ` ${ev}` : ""}` };
  }

  // FALLBACK: curated company list.
  const inList = !!c?.in_pli;
  if (inList) return { points: 2, max: 2, status: "pass", value: "On PLI list", note: "Company has approved PLI allocation or confirmed PLI-scheme revenue." };
  return { points: 0, max: 2, status: "fail", value: "Not on PLI list", note: `Not in the curated list of ${m.pli_companies.length} PLI beneficiaries.` };
}

function ruleRenewableEnergy(c) {
  const m = c?._macro;
  if (!m?.renewable_companies) return { ...NA, max: 2 };
  const inList = !!c?.in_renewable;
  if (inList) return { points: 2, max: 2, status: "pass", value: "On renewable list", note: "Direct participant in renewable capacity addition (solar / wind / green energy)." };
  return { points: 0, max: 2, status: "fail", value: "Not on renewable list", note: `Not in the curated list of ${m.renewable_companies.length} renewable players.` };
}

// ---- master ----

const ACTIVE_RULES = [
  { key: "infra",      label: "Infra Push",         category: "Sector Rotation",     criteria: "Capex-led sector",     fn: ruleInfraPush },
  { key: "ratecut",    label: "Rate Cuts",          category: "Sector Rotation",     criteria: "Rate-sensitive sector", fn: ruleRateCuts },
  { key: "chinaplus1", label: "China+1",            category: "Sector Rotation",     criteria: "Chemicals / EMS",      fn: ruleChinaPlusOne },
  { key: "rural",      label: "Rural Recovery",     category: "Sector Rotation",     criteria: "FMCG / Tractors",      fn: ruleRuralRecovery },
  { key: "gdp",        label: "GDP Growth",         category: "Economic Indicators", criteria: "≥ 6.5% and rising",    fn: ruleGDPGrowth },
  { key: "cpi",        label: "Inflation",          category: "Economic Indicators", criteria: "Within RBI band",      fn: ruleInflation },
  { key: "crude",      label: "Crude Oil",          category: "Economic Indicators", criteria: "Brent < $85 (sector overlay)", fn: ruleCrudeOil },
  { key: "inr",        label: "INR / USD",          category: "Economic Indicators", criteria: "Sector × INR trend",   fn: ruleINRUSD },
  { key: "bonds",      label: "Bond Yields",        category: "Economic Indicators", criteria: "10Y yield direction",  fn: ruleBondYields },
  { key: "pli",        label: "PLI Beneficiary",    category: "Government Policy",   criteria: "On PLI scheme",        fn: rulePLIBeneficiary },
  { key: "renewable",  label: "Renewable Energy",   category: "Government Policy",   criteria: "Renewable participant", fn: ruleRenewableEnergy },
];

const DEFERRED = [];   // no deferred rules in Macro for now

export function scoreCompany(c) {
  const breakdown = ACTIVE_RULES.map((r) => ({ ...r, ...r.fn(c) }));
  const totalPoints = breakdown.reduce((s, b) => s + b.points, 0);
  const totalMax = breakdown.reduce((s, b) => s + b.max, 0);
  const naCount = breakdown.filter((b) => b.status === "na").length;
  return {
    company: c, breakdown, deferred: DEFERRED,
    totalPoints, totalMax,
    scorePct: totalMax ? Math.round((totalPoints / totalMax) * 100) : 0,
    hardFails: [],
    naCount,
  };
}

export { ACTIVE_RULES, DEFERRED };
