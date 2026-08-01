#!/usr/bin/env node
// Screener.in screen scraper (headless-browser version).
// Screener injects custom ribbon ratios via JavaScript, so a plain fetch
// only sees the 9 default ratios. This uses Playwright (headless Chrome)
// to render each page fully before extracting.
import { writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import * as cheerio from "cheerio";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = resolve(__dirname, "output");
const UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";
const ORIGIN = "https://www.screener.in";

const EMAIL = process.env.SCREENER_EMAIL;
const PASSWORD = process.env.SCREENER_PASSWORD;
const screenUrl = process.env.SCREEN_URL || "https://www.screener.in/screens/3675531/fundareal-klp-final/";
const maxPages = Number(process.env.MAX_PAGES || "30");
const maxCompanies = Number(process.env.MAX_COMPANIES || "0"); // 0 = all
const startAt = Math.max(0, Number(process.env.START_AT || "0"));

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

run().catch((err) => {
  console.error("Failed:", err.message);
  process.exit(1);
});

async function run() {
  if (!EMAIL || !PASSWORD) throw new Error("Set SCREENER_EMAIL + SCREENER_PASSWORD in repo secrets.");
  mkdirSync(OUT_DIR, { recursive: true });

  const browser = await chromium.launch();
  const context = await browser.newContext({ userAgent: UA });
  const page = await context.newPage();

  try {
    console.log("Logging in to Screener...");
    await loginViaBrowser(page);
    console.log("Login OK.\n");

    const base = screenUrl.split("?")[0].replace(/\/+$/, "");
    const companies = await fetchScreenCompanies(page, base);
    console.log(`\nScreen returned ${companies.length} companies.`);
    if (!companies.length) throw new Error("Screen returned no companies.");

    const start = Math.min(startAt, companies.length);
    const end = maxCompanies > 0 ? Math.min(start + maxCompanies, companies.length) : companies.length;
    const total = end - start;
    if (startAt > 0) console.log(`Resuming from index ${start} (skipping first ${start}).`);
    console.log(`Scraping ${total} company pages (range ${start}..${end - 1})...\n`);

    const rows = [];
    const allKeys = new Set(["Company", "Screener URL"]);
    let failures = 0;
    const csvPath = resolve(OUT_DIR, "screener-companies.csv");
    const jsonPath = resolve(OUT_DIR, "screener-companies.json");
    const flush = () => {
      const headers = [...allKeys];
      const csv = [headers, ...rows.map((r) => headers.map((h) => r[h] ?? ""))]
        .map((line) => line.map(csvCell).join(","))
        .join("\n");
      writeFileSync(csvPath, csv + "\n");
      writeFileSync(jsonPath, JSON.stringify(rows, null, 2) + "\n");
    };

    for (let i = start; i < end; i++) {
      const c = companies[i];
      process.stdout.write(`[${i + 1}/${companies.length}] ${c.name} ... `);
      try {
        const data = await fetchCompanyData(page, c.path, i === start);
        const row = { Company: c.name, "Screener URL": `${ORIGIN}${c.path}`, ...data };
        Object.keys(data).forEach((k) => allKeys.add(k));
        rows.push(row);
        console.log(`${Object.keys(data).length} fields`);
      } catch (err) {
        failures++;
        rows.push({ Company: c.name, "Screener URL": `${ORIGIN}${c.path}`, Error: err.message });
        console.log(`FAILED: ${err.message}`);
      }
      flush();
      await sleep(300);
    }

    const headers = [...allKeys];

    // Write a small metadata file the dashboard reads for "last updated" etc.
    writeFileSync(resolve(OUT_DIR, "metadata.json"), JSON.stringify({
      generated_at: new Date().toISOString(),
      source: `Screener saved screen: ${screenUrl}`,
      company_count: rows.length,
      failures,
      columns: headers,
    }, null, 2) + "\n");

    console.log("\n=== Done ===");
    console.log(`Companies scraped: ${rows.length} (${failures} failed)`);
    console.log(`Columns found: ${headers.length}`);
    console.log(`Columns: ${headers.join(" | ")}`);
  } finally {
    await browser.close();
  }
}

async function loginViaBrowser(page) {
  await page.goto(`${ORIGIN}/login/`, { waitUntil: "domcontentloaded", timeout: 60000 });
  await page.fill('input[name="username"]', EMAIL);
  await page.fill('input[name="password"]', PASSWORD);
  await Promise.all([
    page.waitForLoadState("domcontentloaded").catch(() => {}),
    page.click('button[type="submit"]')
  ]);
  await page.waitForTimeout(1500);
  const html = await page.content();
  if (!/\/logout\//.test(html)) {
    throw new Error("Login failed — check SCREENER_EMAIL / SCREENER_PASSWORD.");
  }
}

async function fetchScreenCompanies(page, base) {
  const companies = [];
  const seen = new Set();
  let totalPages = 1;

  for (let p = 1; p <= maxPages; p++) {
    process.stdout.write(`Reading screen page ${p} ... `);
    await page.goto(`${base}/?page=${p}`, { waitUntil: "domcontentloaded", timeout: 60000 });
    const $ = cheerio.load(await page.content());
    const table = $("table.data-table").first();
    if (!table.length) {
      console.log("no table");
      if (p === 1) throw new Error("No data table on screen page 1.");
      break;
    }
    let added = 0;
    table.find("tbody tr").each((_, tr) => {
      const link = $(tr).find("td a[href^='/company/']").first();
      const href = link.attr("href");
      if (href && !seen.has(href)) {
        seen.add(href);
        companies.push({ name: link.text().trim().replace(/\s+/g, " "), path: href });
        added++;
      }
    });
    const m = $("body").text().match(/page\s+\d+\s+of\s+(\d+)/i);
    if (m) totalPages = Number(m[1]);
    console.log(`${added} companies (page ${p} of ${totalPages})`);
    if (p >= totalPages) break;
    await sleep(300);
  }
  return companies;
}

async function fetchCompanyData(page, path, debug = false) {
  const url = `${ORIGIN}${path}`;
  let html = null;
  let ribbonCount = 0;

  // Screener fetches custom ribbon ratios via XHR after page load. Poll the
  // DOM for them directly — networkidle is unreliable here (analytics never
  // settle). Retry the load once with extra patience if the first miss.
  for (let attempt = 1; attempt <= 2; attempt++) {
    const patient = attempt === 2;
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
    await page
      .waitForFunction(() => document.querySelectorAll("#top-ratios li").length > 9, {
        timeout: patient ? 25000 : 12000,
      })
      .catch(() => {});
    await page.waitForTimeout(patient ? 1200 : 400);

    // Best-effort: expand the Shareholding sub-rows (Pledged shares is the
    // one we care about — Screener collapses it under the Promoters row
    // unless promoters have actually pledged). Try a small set of selectors
    // that have all been seen on Screener over the last year. Failures here
    // are silent — most companies have no pledge sub-row to expand and we
    // don't want a missing button to abort the company scrape.
    try {
      const expandSelectors = [
        '#shareholding button[onclick*="shareholding_change"]',
        '#shareholding button.button-secondary',
        '#shareholding button:has-text("+")',
        '#shareholding [data-toggle="collapse"]',
      ];
      for (const sel of expandSelectors) {
        const btns = await page.locator(sel).all().catch(() => []);
        for (const btn of btns) {
          await btn.click({ timeout: 1500 }).catch(() => {});
        }
        if (btns.length) break; // first matching selector wins
      }
      await page.waitForTimeout(300);
    } catch { /* expander absent — fine */ }

    html = await page.content();
    ribbonCount = cheerio.load(html)("#top-ratios li").length;
    if (ribbonCount > 9) break;
    if (attempt < 2) {
      process.stdout.write(`(only ${ribbonCount} ratios; retrying) `);
      await sleep(1500);
    }
  }

  if (debug) {
    writeFileSync(resolve(OUT_DIR, "_debug-first-company.html"), html);
  }
  const $ = cheerio.load(html);

  const data = {};
  $("#top-ratios li").each((_, li) => {
    const name = $(li).find(".name").text().trim().replace(/\s+/g, " ");
    const value = $(li).find(".value").text().trim().replace(/\s+/g, " ");
    if (name) data[name] = cleanValue(value);
  });
  if (!Object.keys(data).length) throw new Error("no #top-ratios block found");
  if (ribbonCount <= 9) {
    throw new Error(`custom ribbon ratios did not load after retry (got ${ribbonCount}, expected >9)`);
  }

  // Sector classification (4 levels). Used by the dashboard to apply
  // sector-specific scoring exceptions (e.g. exclude financial sector from D/E).
  data["Broad Sector"] = $('p.sub a[title="Broad Sector"]').first().text().trim();
  data["Sector"] = $('p.sub a[title="Sector"]').first().text().trim();
  data["Broad Industry"] = $('p.sub a[title="Broad Industry"]').first().text().trim();
  data["Industry"] = $('p.sub a[title="Industry"]').first().text().trim();

  const npq = parseSectionRow($, "#quarters", /^net profit/i);
  if (npq && npq.values.length) {
    const last4 = npq.values.slice(-4);
    data["Net Profit Qtr (latest)"] = last4[3] ?? "";
    data["Net Profit Qtr (-1)"] = last4[2] ?? "";
    data["Net Profit Qtr (-2)"] = last4[1] ?? "";
    data["Net Profit Qtr (-3)"] = last4[0] ?? "";
  }

  // PR B: deeper history series — last 8 quarters EPS, last 10 yrs CFO,
  // last 5 yrs OPM. Used by EPS Growth, Operating Cash Flow, and EBITDA
  // Margin rules to match the client's multi-period framework.
  const epsQ = parseSectionRow($, "#quarters", /^eps/i);
  if (epsQ && epsQ.values.length) {
    const last8 = epsQ.values.slice(-8).map((v) => parseFloat(String(v).replace(/[^0-9.\-]/g, "")));
    data["EPS Qtr Series"] = last8.filter((n) => Number.isFinite(n)).join("|");
  }

  const cf = parseSectionRow($, "#cash-flow", /cash from operating/i);
  if (cf && cf.values.length) {
    const yearly = cf.headers
      .map((h, i) => ({ h, v: cf.values[i] }))
      .filter((p) => p.v !== undefined && !/ttm/i.test(p.h))
      .map((p) => p.v);
    data["CF Operations LY"] = yearly[yearly.length - 1] ?? "";
    data["CF Operations PY"] = yearly[yearly.length - 2] ?? "";
    // Full series (oldest → newest) — up to 10 years.
    const series = yearly
      .map((s) => parseFloat(String(s).replace(/,/g, "").replace(/[^0-9.\-]/g, "")))
      .filter((n) => Number.isFinite(n));
    data["CF Operations Series"] = series.join("|");
  }

  const opm = parseSectionRow($, "#profit-loss", /^opm\s*%/i);
  if (opm && opm.values.length) {
    const nums = opm.headers
      .map((h, i) => ({ h, v: opm.values[i] }))
      .filter((p) => p.v !== undefined && !/ttm/i.test(p.h))
      .map((p) => parseFloat(String(p.v).replace(/[^0-9.\-]/g, "")))
      .filter((n) => Number.isFinite(n));
    const last5 = nums.slice(-5);
    data["OPM 5Year %"] = last5.length
      ? (last5.reduce((a, b) => a + b, 0) / last5.length).toFixed(2)
      : "";
    // Full OPM series for the EBITDA Margin YoY rule (multi-year contraction check).
    data["OPM Series"] = nums.join("|");
  }

  // PR D: deeper history for the remaining 3 Fundamentals deviations.
  // All from the same Screener company page tables we already render —
  // just read more rows.

  // Promoter Pledge % across recent quarters (Shareholding section).
  // Note: Screener row labels have a trailing "+" because they're expandable —
  // regexes use ^prefix matching not full-match. Pledge row only exists for
  // companies where promoters have actually pledged; we tolerate its absence.
  data["Pledge Series"] = extractSeries($, "#shareholding", /pledge/i, 8);
  data["FII Holding Series"] = extractSeries($, "#shareholding", /^FIIs?\b/i, 8);
  data["DII Holding Series"] = extractSeries($, "#shareholding", /^DIIs?\b/i, 8);

  // Dividend Payout % across annual periods (>0 means dividend paid that year).
  // From P&L table, skip the TTM column.
  data["Dividend Series"] = extractAnnualSeries($, "#profit-loss", /^dividend payout/i);

  // Future-proof: full annual revenue + PAT series. Lets us re-verify
  // 3Y / 5Y CAGRs client-side later without re-scraping.
  data["Revenue Series"] = extractAnnualSeries($, "#profit-loss", /^sales/i);
  data["PAT Series"]     = extractAnnualSeries($, "#profit-loss", /^net profit/i);

  return data;
}

// Helper: pull a single labelled row out of a section table and return the
// last N parsed numeric values as a pipe-delimited string. Strips ₹ / % /
// commas, skips non-numeric cells.
function extractSeries($, sectionSel, labelRegex, lastN) {
  const row = parseSectionRow($, sectionSel, labelRegex);
  if (!row || !row.values.length) return "";
  const nums = row.values
    .map((v) => parseFloat(String(v).replace(/,/g, "").replace(/[^0-9.\-]/g, "")))
    .filter((n) => Number.isFinite(n));
  return nums.slice(-lastN).join("|");
}

// Helper for annual series — also skip the TTM column when present.
function extractAnnualSeries($, sectionSel, labelRegex) {
  const row = parseSectionRow($, sectionSel, labelRegex);
  if (!row || !row.values.length) return "";
  const nums = row.headers
    .map((h, i) => ({ h, v: row.values[i] }))
    .filter((p) => p.v !== undefined && !/ttm/i.test(p.h))
    .map((p) => parseFloat(String(p.v).replace(/,/g, "").replace(/[^0-9.\-]/g, "")))
    .filter((n) => Number.isFinite(n));
  return nums.join("|");
}

function parseSectionRow($, sectionSel, labelRegex) {
  const table = $(`${sectionSel} table`).first();
  if (!table.length) return null;
  const headers = [];
  table.find("thead th").each((_, th) => headers.push($(th).text().trim().replace(/\s+/g, " ")));
  let result = null;
  table.find("tbody tr").each((_, tr) => {
    if (result) return;
    const cells = $(tr).find("td");
    const label = $(cells[0]).text().trim().replace(/\s+/g, " ");
    if (labelRegex.test(label)) {
      const values = [];
      cells.each((i, td) => {
        if (i > 0) values.push($(td).text().trim().replace(/\s+/g, " "));
      });
      result = { headers: headers.slice(1), values };
    }
  });
  return result;
}

function cleanValue(v) {
  return String(v || "").replace(/\s+/g, " ").replace(/^₹\s*/, "").trim();
}

function csvCell(v) {
  const s = String(v ?? "");
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
