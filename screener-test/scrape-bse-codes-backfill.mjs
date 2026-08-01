#!/usr/bin/env node
// One-off BSE-code backfill — for companies whose Screener URL slug is
// an NSE ticker (so we don't have their BSE numeric scrip code in our
// data). Visits each Screener company page with Playwright, scrapes the
// BSE code from the page header, writes a static supplement file the
// to-do generator reads from.
//
// Resilience pattern (after first attempt got rate-limited around row 30):
//   - 3s throttle between successful requests + 200-800ms jitter
//   - 90s navigation timeout (up from 45s)
//   - On request failure: retry once after 30s
//   - After 3 consecutive failures: pause 5 min + re-login
//   - CHECKPOINT after every successful capture — supplement file
//     written to disk, so re-running picks up where we left off
//   - Re-running is idempotent (already-captured slugs are skipped)
//
// Cost: 0 (uses existing SCREENER_EMAIL/SCREENER_PASSWORD secrets,
// runs on GH Actions). Runtime: ~10-30 min for ~120 companies
// depending on Screener's mood.

import { chromium } from "playwright";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SCREENER_PATH = resolve(__dirname, "../public/data/screener-companies.json");
const OUT_PATH      = resolve(__dirname, "static/bse-codes-supplement.json");

const ORIGIN = "https://www.screener.in";
const EMAIL = process.env.SCREENER_EMAIL;
const PASSWORD = process.env.SCREENER_PASSWORD;

const NAV_TIMEOUT_MS  = 90_000;
const THROTTLE_MS     = 3_000;     // base interval between successful requests
const RETRY_WAIT_MS   = 30_000;    // wait after a single failure before retry
const COOLDOWN_MS     = 5 * 60_000; // 5 min cool-down after 3 consecutive fails
const MAX_CONSEC_FAIL = 3;

if (!EMAIL || !PASSWORD) {
  console.error("SCREENER_EMAIL or SCREENER_PASSWORD env not set — aborting.");
  process.exit(1);
}

const screener = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));

// Resume: load existing supplement so previous runs' captures aren't repeated.
let supplement = {};
if (existsSync(OUT_PATH)) {
  try { supplement = JSON.parse(readFileSync(OUT_PATH, "utf8")).codes || {}; }
  catch { supplement = {}; }
}

function saveSupplement() {
  writeFileSync(OUT_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    source: "Screener company-page header (scraped via Playwright)",
    total_codes: Object.keys(supplement).length,
    codes: supplement,
  }, null, 2) + "\n");
}

const toFetch = [];
for (const c of screener) {
  const m = String(c["Screener URL"] || "").match(/\/company\/([^/]+)/);
  if (!m) continue;
  const slug = m[1].toUpperCase();
  if (/^\d+$/.test(slug)) continue;
  if (supplement[slug]) continue;
  toFetch.push({ name: c.Company, slug, url: c["Screener URL"] });
}
console.log(`Need to backfill ${toFetch.length} of ${screener.length} companies.`);
console.log(`Supplement currently has ${Object.keys(supplement).length} codes.`);

if (toFetch.length === 0) {
  console.log("Nothing to do — supplement already covers everything.");
  process.exit(0);
}

const jitter = () => 200 + Math.floor(Math.random() * 600);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let browser, ctx, page;
async function startBrowser() {
  browser = await chromium.launch();
  ctx = await browser.newContext({ userAgent: "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0 Safari/537.36" });
  page = await ctx.newPage();
}

async function login() {
  console.log("Logging into Screener...");
  await page.goto(`${ORIGIN}/login/`, { waitUntil: "domcontentloaded", timeout: NAV_TIMEOUT_MS });
  await page.fill('input[name="username"]', EMAIL);
  await page.fill('input[name="password"]', PASSWORD);
  await Promise.all([
    page.waitForNavigation({ timeout: NAV_TIMEOUT_MS }),
    page.click('button[type="submit"]'),
  ]);
  console.log("Logged in.");
}

await startBrowser();
await login();

async function captureBseCode(url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: NAV_TIMEOUT_MS });
  const bseCode = await page.evaluate(() => {
    const text = document.body.innerText || "";
    const m = text.match(/\bBSE\s*:\s*(\d{5,8})\b/);
    return m ? m[1] : null;
  });
  return bseCode;
}

let captured = 0, missing = 0, failed = 0, consecFails = 0;
const MAX = Number(process.env.MAX_BACKFILL || 0);
const limit = MAX > 0 ? Math.min(MAX, toFetch.length) : toFetch.length;

for (let i = 0; i < limit; i++) {
  const c = toFetch[i];
  process.stdout.write(`[${i + 1}/${limit}] ${c.name.padEnd(28).slice(0, 28)} `);
  let success = false;
  let attempts = 0;
  while (attempts < 2 && !success) {
    attempts++;
    try {
      const bseCode = await captureBseCode(c.url);
      if (bseCode) {
        supplement[c.slug] = bseCode;
        captured++;
        saveSupplement();   // checkpoint after every capture
        console.log(`✓ ${bseCode}`);
      } else {
        missing++;
        console.log("• no BSE code on page (NSE-only or layout changed)");
      }
      success = true;
      consecFails = 0;
    } catch (err) {
      const msg = (err.message || "").slice(0, 80);
      if (attempts < 2) {
        console.log(`ERR (will retry in ${RETRY_WAIT_MS/1000}s): ${msg}`);
        await sleep(RETRY_WAIT_MS);
      } else {
        console.log(`ERR (giving up): ${msg}`);
        failed++;
        consecFails++;
      }
    }
  }

  // If we've hit MAX_CONSEC_FAIL in a row, Screener has probably
  // throttled our IP. Take a long cool-down + re-login before
  // continuing.
  if (consecFails >= MAX_CONSEC_FAIL) {
    console.log(`\n⚠ ${consecFails} consecutive failures — cooling down ${COOLDOWN_MS/60000} min and re-logging in...`);
    await sleep(COOLDOWN_MS);
    try { await browser.close(); } catch {}
    await startBrowser();
    await login();
    consecFails = 0;
  }

  // Normal between-request throttle (only on success, since retries already waited).
  if (success) await sleep(THROTTLE_MS + jitter());
}

await browser.close();

console.log(`\nDone. Captured ${captured} BSE codes (${missing} pages had no code, ${failed} failed).`);
console.log(`Supplement now covers ${Object.keys(supplement).length} slugs total.`);
console.log(`Wrote ${OUT_PATH}`);

