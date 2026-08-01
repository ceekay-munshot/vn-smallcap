// excel-export.js - Multi-sheet, color-graded Excel workbook export.
//
// Uses ExcelJS (loaded as window.ExcelJS via CDN in index.html). Produces a
// 4-sheet workbook per tab:
//   1. "Ranked Companies"  - main grid with actual numeric values per rule,
//                            color-graded green/amber/red, total-score heat
//                            map, in-cell score % data bars, hyperlinked
//                            company names.
//   2. "Scoring Framework" - client's scoring spec per rule (from rule-meta).
//   3. "Distribution"      - per-rule pass / partial / fail / na counts
//                            across the full universe.
//   4. "Top 25"            - the headline list with key stats only.

// ARGB hex (ExcelJS expects alpha-first 8-char hex).
const COLOR = {
  headerFill:   "FF1E3A8A",  // dark navy
  headerText:   "FFFFFFFF",
  passFill:     "FFD1FAE5",
  passText:     "FF047857",
  partialFill:  "FFFEF3C7",
  partialText:  "FFB45309",
  failFill:     "FFFEE2E2",
  failText:     "FFB91C1C",
  hardFailFill: "FFFECACA",
  hardFailText: "FF991B1B",
  naFill:       "FFF1F5F9",
  naText:       "FF64748B",
  bandFill:     "FFF8FAFC",  // zebra stripe
  borderLight:  "FFE2E8F0",
  link:         "FF1D4ED8",
  // 3-color heat map endpoints for total score
  heatLow:      "FFFCA5A5",  // red
  heatMid:      "FFFDE68A",  // yellow
  heatHigh:     "FF6EE7B7",  // green
  // data-bar color (in-cell horizontal bar) for score percent
  dataBar:      "FF3B82F6",
};

// Series-based rules where we can render the value as "FY21: x -> FY26: y"
// instead of just the latest year. Keyed by tab -> rule key.
// `field` = column in the company JSON; `unit` = suffix to append per value.
const SERIES_TREND_MAP = {
  fundamentals: {
    ebitda: { field: "OPM Series",          unit: "%" },
    cfo:    { field: "CF Operations Series", unit: " Cr" },
    rev3y:  { field: "Revenue Series",       unit: " Cr" },
    pat3y:  { field: "PAT Series",           unit: " Cr" },
    div:    { field: "Dividend Series",      unit: "" },
    npm:    { field: "OPM Series",           unit: "%" }, // we only have OPM not NPM series
  },
};

// Final fiscal year that fully closed in India (April-March). May 2026 -> FY26.
function currentFY() {
  const d = new Date();
  const y = d.getFullYear();
  const m = d.getMonth() + 1; // 1-12
  // FY ends Mar 31. Before April we're still in the FY that started prev April.
  const fyEndYear = m >= 4 ? y : y - 1;
  return fyEndYear % 100; // last two digits, e.g. 26
}

function formatSeriesAsTrend(seriesStr, unit) {
  if (!seriesStr) return null;
  const parts = String(seriesStr).split("|").map((s) => s.trim()).filter(Boolean);
  if (parts.length < 2) return null;
  const endFY = currentFY();
  const startFY = endFY - parts.length + 1;
  const pad = (n) => String(n).padStart(2, "0");
  return `FY${pad(startFY)}: ${parts[0]}${unit} → FY${pad(endFY)}: ${parts[parts.length - 1]}${unit}`;
}

// Status -> {fill, text} colors. Used per cell.
function statusColors(status) {
  switch (status) {
    case "pass":      return { fill: COLOR.passFill,     text: COLOR.passText,     bold: false };
    case "partial":   return { fill: COLOR.partialFill,  text: COLOR.partialText,  bold: false };
    case "fail":      return { fill: COLOR.failFill,     text: COLOR.failText,     bold: false };
    case "hard_fail": return { fill: COLOR.hardFailFill, text: COLOR.hardFailText, bold: true };
    case "na":        return { fill: COLOR.naFill,       text: COLOR.naText,       bold: false };
    default:          return null;
  }
}

function applyCellStyle(cell, status, opts = {}) {
  const c = statusColors(status);
  if (!c) return;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.fill } };
  cell.font = { color: { argb: c.text }, bold: c.bold || opts.bold || false, italic: status === "na" };
  cell.alignment = { vertical: "middle", horizontal: opts.align || "center", wrapText: true };
  cell.border = {
    top:    { style: "thin", color: { argb: COLOR.borderLight } },
    left:   { style: "thin", color: { argb: COLOR.borderLight } },
    bottom: { style: "thin", color: { argb: COLOR.borderLight } },
    right:  { style: "thin", color: { argb: COLOR.borderLight } },
  };
}

function styleHeaderRow(row) {
  row.height = 28;
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR.headerFill } };
    cell.font = { color: { argb: COLOR.headerText }, bold: true, size: 11 };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      top:    { style: "thin", color: { argb: COLOR.headerFill } },
      left:   { style: "thin", color: { argb: COLOR.headerFill } },
      bottom: { style: "medium", color: { argb: COLOR.headerFill } },
      right:  { style: "thin", color: { argb: COLOR.headerFill } },
    };
  });
}

// Extract the displayed text from a cell value (handles hyperlink objects).
function cellText(value) {
  if (value == null) return "";
  if (typeof value === "object") return value.text || value.richText?.map((r) => r.text).join("") || "";
  return String(value);
}

// After data is in place, walk the sheet and:
//   1. Widen each column to fit its actual longest content (clamped by
//      minWidth/maxWidth), respecting any preset width that's already wider.
//   2. Grow each data-row height enough to show wrapped multi-line content.
// This is our best ExcelJS substitute for "auto-fit columns / rows" — the
// library has no native equivalent, and our previous fixed heights of 22pt
// were truncating wrapped trend cells like
// "FII 18.56 → 18.14 → 18.51 → 19.51 → DII 25.45 → 23.6" into adjacent rows.
function autoFitSheet(sheet, opts = {}) {
  const minWidth     = opts.minWidth     ?? 10;
  const maxWidth     = opts.maxWidth     ?? 40;
  const headerHeight = opts.headerHeight ?? 28;
  const minRowHeight = opts.minRowHeight ?? 22;
  const lineHeight   = opts.lineHeight   ?? 14;

  // Pass 1: determine max content length per column.
  const colMaxLen = {};
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const text = cellText(cell.value);
      const lineLens = text.split(/\n/).map((s) => s.length);
      const longest = Math.max(0, ...lineLens);
      if (longest > (colMaxLen[col] || 0)) colMaxLen[col] = longest;
    });
  });
  // Apply widths (don't shrink columns that were explicitly set wider).
  Object.entries(colMaxLen).forEach(([col, len]) => {
    const target = Math.max(minWidth, Math.min(maxWidth, len + 2));
    const c = sheet.getColumn(Number(col));
    c.width = Math.max(c.width || 0, target);
  });

  // Pass 2: row heights based on wrap-line estimate.
  sheet.eachRow({ includeEmpty: false }, (row, rowIdx) => {
    if (rowIdx === 1) { row.height = headerHeight; return; }
    let maxLines = 1;
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const text = cellText(cell.value);
      if (!text) return;
      const wrapText = cell.alignment?.wrapText !== false;
      const colWidth = sheet.getColumn(col).width || minWidth;
      // Lines come from explicit \n plus, if wrapping, ceil(line / colWidth).
      const explicitLines = text.split(/\n/);
      let total = 0;
      for (const line of explicitLines) {
        if (wrapText) total += Math.max(1, Math.ceil(line.length / Math.max(1, colWidth - 1)));
        else total += 1;
      }
      if (total > maxLines) maxLines = total;
    });
    row.height = Math.max(minRowHeight, maxLines * lineHeight + 6);
  });
}

// ---------------- Sheet 1: Ranked Companies ----------------
function buildRankedSheet(wb, tab, cfg, scored, ruleMetaForTab) {
  const sheet = wb.addWorksheet("Ranked Companies", {
    views: [{ state: "frozen", xSplit: 2, ySplit: 1 }],
  });

  // Build column definitions. Widths are tentative — the auto-fit pass at
  // the end of buildRankedSheet will widen narrow columns based on actual
  // content and grow row heights to accommodate wrapped text.
  const columns = [
    { header: "Rank",       key: "rank",      width: 7 },
    { header: "Company",    key: "company",   width: 34 },
    { header: "Score",      key: "score",     width: 9 },
    { header: "Score %",    key: "scorePct",  width: 11 },
    { header: "Hard Fails", key: "hardFails", width: 22 },
  ];
  cfg.rules.forEach((r) => {
    columns.push({ header: r.label, key: r.key, width: Math.max(18, r.label.length + 2) });
  });
  sheet.columns = columns;

  // Header row.
  styleHeaderRow(sheet.getRow(1));

  // Data rows.
  const trendMap = SERIES_TREND_MAP[tab] || {};
  scored.forEach((s, idx) => {
    const company = s.company;
    const rowIdx = idx + 2;
    const row = sheet.getRow(rowIdx);

    row.getCell("rank").value = idx + 1;
    row.getCell("rank").alignment = { vertical: "middle", horizontal: "center" };

    // Company name as hyperlink to Screener.
    const screenerUrl = cfg.screenerUrl(company);
    const companyCell = row.getCell("company");
    if (screenerUrl) {
      companyCell.value = { text: cfg.name(company), hyperlink: screenerUrl };
      companyCell.font = { color: { argb: COLOR.link }, underline: true, bold: true };
    } else {
      companyCell.value = cfg.name(company);
      companyCell.font = { bold: true };
    }
    companyCell.alignment = { vertical: "middle", horizontal: "left" };

    // Total score (numeric).
    row.getCell("score").value = s.totalPoints;
    row.getCell("score").numFmt = "0.0";
    row.getCell("score").alignment = { vertical: "middle", horizontal: "center" };
    row.getCell("score").font = { bold: true };

    // Score percent (0-100, formatted with data bar via conditional formatting).
    row.getCell("scorePct").value = s.scorePct;
    row.getCell("scorePct").numFmt = '0"%"';
    row.getCell("scorePct").alignment = { vertical: "middle", horizontal: "center" };

    // Hard fails.
    const hf = s.hardFails || [];
    const hfCell = row.getCell("hardFails");
    hfCell.value = hf.length ? hf.join("; ") : "—";
    if (hf.length) {
      hfCell.font = { color: { argb: COLOR.hardFailText }, bold: true };
      hfCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR.hardFailFill } };
    } else {
      hfCell.font = { color: { argb: COLOR.naText } };
      hfCell.alignment = { horizontal: "center" };
    }

    // Per-rule cells: show the actual value, color-graded by status.
    cfg.rules.forEach((r) => {
      const b = s.breakdown.find((x) => x.key === r.key);
      const cell = row.getCell(r.key);
      if (!b || b.status === "na") {
        cell.value = "—";
        applyCellStyle(cell, "na");
        return;
      }
      // Trend format wins for series-based rules.
      const trendCfg = trendMap[r.key];
      const trendValue = trendCfg ? formatSeriesAsTrend(company[trendCfg.field], trendCfg.unit) : null;
      cell.value = trendValue || b.value || `${b.points}/${b.max}`;
      applyCellStyle(cell, b.status, { align: trendValue ? "left" : "center" });
    });

    // Zebra stripe alternate rows (skip header).
    if (rowIdx % 2 === 1) {
      ["rank", "company", "scorePct"].forEach((k) => {
        const c = row.getCell(k);
        if (!c.fill || !c.fill.fgColor) {
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR.bandFill } };
        }
      });
    }
    // Row height is set by the autoFitSheet pass after all data is in place.
  });

  // Heat-map color scale on total Score column (col C = index 3).
  const lastRow = scored.length + 1;
  if (lastRow > 1) {
    // Score: 3-color heat map (red -> yellow -> green) based on points.
    const maxScore = scored.length ? Math.max(...scored.map((s) => s.totalMax || 0)) : 0;
    sheet.addConditionalFormatting({
      ref: `C2:C${lastRow}`,
      rules: [{
        type: "colorScale",
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: maxScore / 2 },
          { type: "num", value: maxScore },
        ],
        color: [
          { argb: COLOR.heatLow },
          { argb: COLOR.heatMid },
          { argb: COLOR.heatHigh },
        ],
      }],
    });
    // Score %: in-cell data bar.
    sheet.addConditionalFormatting({
      ref: `D2:D${lastRow}`,
      rules: [{
        type: "dataBar",
        cfvo: [{ type: "num", value: 0 }, { type: "num", value: 100 }],
        color: { argb: COLOR.dataBar },
      }],
    });
  }

  // Auto-filter on full data range.
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to:   { row: lastRow, column: columns.length },
  };

  autoFitSheet(sheet, { minWidth: 10, maxWidth: 38, lineHeight: 14, minRowHeight: 24 });
}

// ---------------- Sheet 2: Scoring Framework ----------------
function buildFrameworkSheet(wb, tab, cfg, ruleMetaForTab) {
  const sheet = wb.addWorksheet("Scoring Framework", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  // Show a Pillar column when rules are tagged (composite-tab export).
  const hasPillarTag = cfg.rules.some((r) => r._pillar);
  const PILLAR_LABEL = { fundamentals: "Fundamentals", technicals: "Technicals", macro: "Macro", sentiment: "Sentiment / Liquidity" };
  const baseCols = [
    { header: "Rule",          key: "rule",      width: 32 },
    ...(hasPillarTag ? [{ header: "Pillar", key: "pillar", width: 20 }] : []),
    { header: "Category",      key: "category",  width: 18 },
    { header: "Max Points",    key: "max",       width: 12 },
    { header: "Short Criteria", key: "criteria", width: 28 },
    { header: "Client Scoring Logic", key: "logic", width: 80 },
  ];
  sheet.columns = baseCols;
  styleHeaderRow(sheet.getRow(1));

  cfg.rules.forEach((r, i) => {
    const row = sheet.getRow(i + 2);
    row.getCell("rule").value = r.label;
    row.getCell("rule").font = { bold: true };
    if (hasPillarTag) {
      row.getCell("pillar").value = PILLAR_LABEL[r._pillar] || r._pillar || "—";
      row.getCell("pillar").alignment = { horizontal: "center" };
    }
    row.getCell("category").value = r.category || "—";
    row.getCell("max").value = r.max || "";
    row.getCell("max").alignment = { horizontal: "center" };
    row.getCell("criteria").value = r.criteria || "—";

    const logicCell = row.getCell("logic");
    // Composite export passes a flattened meta under the "composite" tab key,
    // but the existing lookup also works for per-pillar tabs unchanged.
    const meta = ruleMetaForTab?.[r.key];
    logicCell.value = meta?.clientLogic || "—";
    logicCell.alignment = { wrapText: true, vertical: "top" };

    // Zebra stripe.
    if (i % 2 === 0) {
      ["rule", ...(hasPillarTag ? ["pillar"] : []), "category", "max", "criteria", "logic"].forEach((k) => {
        row.getCell(k).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR.bandFill } };
      });
    }
  });

  // Deferred rules section (if any) - appended below.
  if (cfg.deferred && cfg.deferred.length) {
    const startRow = cfg.rules.length + 3;
    const titleRow = sheet.getRow(startRow);
    titleRow.getCell(1).value = "Deferred Rules (not currently scored)";
    titleRow.getCell(1).font = { bold: true, italic: true, color: { argb: COLOR.naText } };
    titleRow.getCell(1).alignment = { vertical: "middle" };

    cfg.deferred.forEach((d, i) => {
      const row = sheet.getRow(startRow + 1 + i);
      row.getCell("rule").value = d.label;
      row.getCell("rule").font = { italic: true };
      row.getCell("category").value = d.category || "—";
      row.getCell("max").value = d.max || "";
      row.getCell("criteria").value = "—";
      row.getCell("logic").value = d.reason || "";
      ["rule", "category", "max", "criteria", "logic"].forEach((k) => {
        row.getCell(k).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR.naFill } };
        const cur = row.getCell(k).font || {};
        row.getCell(k).font = { ...cur, color: { argb: COLOR.naText } };
      });
    });
  }

  autoFitSheet(sheet, { minWidth: 12, maxWidth: 80, lineHeight: 14, minRowHeight: 28 });
}

// ---------------- Sheet 3: Distribution ----------------
function buildDistributionSheet(wb, cfg, scored) {
  const sheet = wb.addWorksheet("Distribution", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  sheet.columns = [
    { header: "Rule",       key: "rule",     width: 32 },
    { header: "Pass",       key: "pass",     width: 10 },
    { header: "Partial",    key: "partial",  width: 10 },
    { header: "Fail",       key: "fail",     width: 10 },
    { header: "Hard Fail",  key: "hardFail", width: 12 },
    { header: "N/A",        key: "na",       width: 8 },
    { header: "Pass Rate",  key: "passRate", width: 12 },
  ];
  styleHeaderRow(sheet.getRow(1));

  cfg.rules.forEach((r, i) => {
    const counts = { pass: 0, partial: 0, fail: 0, hard_fail: 0, na: 0 };
    scored.forEach((s) => {
      const b = s.breakdown.find((x) => x.key === r.key);
      if (b) counts[b.status] = (counts[b.status] || 0) + 1;
    });
    const ratedTotal = counts.pass + counts.partial + counts.fail + counts.hard_fail;
    const passRate = ratedTotal ? (counts.pass / ratedTotal) * 100 : 0;
    const row = sheet.getRow(i + 2);
    row.getCell("rule").value = r.label;
    row.getCell("rule").font = { bold: true };

    const cells = [
      { key: "pass",     value: counts.pass,      color: COLOR.passFill,     text: COLOR.passText },
      { key: "partial",  value: counts.partial,   color: COLOR.partialFill,  text: COLOR.partialText },
      { key: "fail",     value: counts.fail,      color: COLOR.failFill,     text: COLOR.failText },
      { key: "hardFail", value: counts.hard_fail, color: COLOR.hardFailFill, text: COLOR.hardFailText },
      { key: "na",       value: counts.na,        color: COLOR.naFill,       text: COLOR.naText },
    ];
    cells.forEach((c) => {
      const cell = row.getCell(c.key);
      cell.value = c.value;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      if (c.value > 0) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.color } };
        cell.font = { color: { argb: c.text }, bold: true };
      } else {
        cell.font = { color: { argb: COLOR.naText } };
      }
    });

    const rateCell = row.getCell("passRate");
    rateCell.value = passRate / 100;
    rateCell.numFmt = "0.0%";
    rateCell.alignment = { horizontal: "center", vertical: "middle" };
    rateCell.font = { bold: true };

    if (i % 2 === 0) {
      row.getCell("rule").fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR.bandFill } };
    }
  });

  // Heat-map on the pass-rate column (G).
  const lastRow = cfg.rules.length + 1;
  sheet.addConditionalFormatting({
    ref: `G2:G${lastRow}`,
    rules: [{
      type: "colorScale",
      cfvo: [
        { type: "num", value: 0 },
        { type: "num", value: 0.5 },
        { type: "num", value: 1 },
      ],
      color: [
        { argb: COLOR.heatLow },
        { argb: COLOR.heatMid },
        { argb: COLOR.heatHigh },
      ],
    }],
  });

  autoFitSheet(sheet, { minWidth: 10, maxWidth: 30, lineHeight: 14, minRowHeight: 22 });
}

// ---------------- Sheet 4: Top 25 ----------------
function buildTop25Sheet(wb, cfg, scored) {
  const sheet = wb.addWorksheet("Top 25", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  sheet.columns = [
    { header: "Rank",         key: "rank",       width: 6 },
    { header: "Company",      key: "company",    width: 34 },
    { header: "Sector",       key: "sector",     width: 22 },
    { header: "Market Cap",   key: "marketCap",  width: 14 },
    { header: "Score",        key: "score",      width: 9 },
    { header: "Score %",      key: "scorePct",   width: 10 },
    { header: "Hard Fails",   key: "hardFails",  width: 18 },
  ];
  styleHeaderRow(sheet.getRow(1));

  const top = scored.slice(0, 25);
  top.forEach((s, idx) => {
    const row = sheet.getRow(idx + 2);
    const c = s.company;
    row.getCell("rank").value = idx + 1;
    row.getCell("rank").alignment = { horizontal: "center" };
    row.getCell("rank").font = { bold: true };

    const screenerUrl = cfg.screenerUrl(c);
    const companyCell = row.getCell("company");
    if (screenerUrl) {
      companyCell.value = { text: cfg.name(c), hyperlink: screenerUrl };
      companyCell.font = { color: { argb: COLOR.link }, underline: true, bold: true };
    } else {
      companyCell.value = cfg.name(c);
      companyCell.font = { bold: true };
    }

    row.getCell("sector").value = cfg.sector(c) || "—";
    row.getCell("marketCap").value = cfg.marketCap(c) || "—";
    row.getCell("marketCap").alignment = { horizontal: "right" };
    row.getCell("score").value = s.totalPoints;
    row.getCell("score").numFmt = "0.0";
    row.getCell("score").alignment = { horizontal: "center" };
    row.getCell("score").font = { bold: true };
    row.getCell("scorePct").value = s.scorePct;
    row.getCell("scorePct").numFmt = '0"%"';
    row.getCell("scorePct").alignment = { horizontal: "center" };
    const hf = s.hardFails || [];
    const hfCell = row.getCell("hardFails");
    hfCell.value = hf.length ? hf.join("; ") : "—";
    if (hf.length) {
      hfCell.font = { color: { argb: COLOR.hardFailText }, bold: true };
    } else {
      hfCell.alignment = { horizontal: "center" };
      hfCell.font = { color: { argb: COLOR.naText } };
    }
  });

  // Heat-map on Score column (E).
  const lastRow = top.length + 1;
  if (top.length) {
    const maxScore = Math.max(...top.map((s) => s.totalMax || 0));
    sheet.addConditionalFormatting({
      ref: `E2:E${lastRow}`,
      rules: [{
        type: "colorScale",
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: maxScore / 2 },
          { type: "num", value: maxScore },
        ],
        color: [
          { argb: COLOR.heatLow },
          { argb: COLOR.heatMid },
          { argb: COLOR.heatHigh },
        ],
      }],
    });
    sheet.addConditionalFormatting({
      ref: `F2:F${lastRow}`,
      rules: [{
        type: "dataBar",
        cfvo: [{ type: "num", value: 0 }, { type: "num", value: 100 }],
        color: { argb: COLOR.dataBar },
      }],
    });
  }

  autoFitSheet(sheet, { minWidth: 8, maxWidth: 36, lineHeight: 14, minRowHeight: 22 });
}

// ---------------- Public API ----------------
export async function exportToExcel({ tab, tabLabel, cfg, scored, ruleMeta }) {
  if (!window.ExcelJS) {
    alert("Excel export library failed to load. Refresh the page and try again.");
    return;
  }
  const wb = new window.ExcelJS.Workbook();
  wb.creator = "VN Smallcap Dash";
  wb.created = new Date();
  wb.title = `${tabLabel} - Ranked Companies`;

  const ruleMetaForTab = ruleMeta?.[tab] || {};
  buildRankedSheet(wb, tab, cfg, scored, ruleMetaForTab);
  buildFrameworkSheet(wb, tab, cfg, ruleMetaForTab);
  buildDistributionSheet(wb, cfg, scored);
  buildTop25Sheet(wb, cfg, scored);

  const today = new Date().toISOString().slice(0, 10);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `lkp-${tab}-${today}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
