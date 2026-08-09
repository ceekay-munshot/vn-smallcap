#!/usr/bin/env node
// Maintains public/data/pinned-companies.json — the set of companies we keep
// tracking after they have left the Rs 2,000-12,500 Cr screen.
//
// Run this immediately AFTER the fundamentals scrape has written the new
// screener-companies.json into public/data, but BEFORE that file is committed.
// At that moment the working tree holds the new universe and git HEAD still
// holds the previous one, which is the only place a just-departed company's
// fundamentals row still exists. Miss that window and the row is gone for good.
//
// Who gets pinned: any member of a still-active monthly cohort that is absent
// from the new universe. The cohort is derived from the snapshot trail using
// the same rule the dashboard uses -- top COHORT_SIZE by composite among
// stocks that are rated, data-complete and not hard-failed, taken from the
// first snapshot of each month within TRACK_MONTHS.
//
// Pins expire on their own: once a cohort ages past TRACK_MONTHS, or the
// company comes back inside the band, the pin is dropped on the next run.
//
// No snapshots yet means no cohorts, which means no pins. That is correct,
// not a failure -- the script is a safe no-op until tracking has begun.

import { readFileSync, writeFileSync, readdirSync, existsSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { PINNED_FILENAME, tickerFromUrl, readPinned } from "./lib/pinned.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR      = resolve(__dirname, "../public/data");
const UNIVERSE_PATH = resolve(DATA_DIR, "screener-companies.json");
const SNAP_DIR      = resolve(DATA_DIR, "snapshots");
const OUT_PATH      = resolve(DATA_DIR, PINNED_FILENAME);
const REPO_REL      = "public/data/screener-companies.json";

const COHORT_SIZE  = Number(process.env.PIN_COHORT_SIZE || 7);
const TRACK_MONTHS = Number(process.env.PIN_TRACK_MONTHS || 5);

run();

function run() {
  const universe = readJson(UNIVERSE_PATH) || [];
  const live = new Set(universe.map((c) => tickerFromUrl(c["Screener URL"])).filter(Boolean));
  console.log(`Live universe: ${live.size} tickers`);

  const cohortTickers = activeCohortTickers();
  if (!cohortTickers.size) {
    console.log("No cohorts in the snapshot trail yet — nothing to pin.");
    writeOut([], { reason: "no cohorts yet" });
    return;
  }
  console.log(`Active cohort members across last ${TRACK_MONTHS} month(s): ${cohortTickers.size}`);

  const departed = [...cohortTickers].filter((t) => !live.has(t));
  if (!departed.length) {
    console.log("Every active cohort member is still inside the band — nothing to pin.");
    writeOut([], { reason: "no departures" });
    return;
  }
  console.log(`Cohort members no longer in the universe: ${departed.join(", ")}`);

  // Last-known fundamentals row: prefer the previous committed universe (the
  // company was still in it as of the last refresh), else an existing pin.
  const prevRows = previousUniverse();
  const prevByTicker = new Map(
    prevRows.map((c) => [tickerFromUrl(c["Screener URL"]), c]).filter(([t]) => t),
  );
  const existing = new Map(
    readPinned(DATA_DIR).map((c) => [tickerFromUrl(c["Screener URL"]), c]).filter(([t]) => t),
  );

  const pins = [];
  const lost = [];
  for (const t of departed) {
    const row = prevByTicker.get(t) || existing.get(t);
    if (!row) { lost.push(t); continue; }
    pins.push({
      ...row,
      _pinned: true,
      _pinned_at: existing.get(t)?._pinned_at || new Date().toISOString().slice(0, 10),
      _pinned_reason: "cohort member outside the Rs 2,000-12,500 Cr band",
    });
  }

  if (lost.length) {
    // Only reachable if a company left before pinning was in place, or two
    // refreshes ran back to back. Loud, because the position is now untracked.
    console.log(`WARNING: no last-known row for ${lost.join(", ")} — cannot pin, their cohort returns will be incomplete.`);
  }
  console.log(`Pinning ${pins.length}: ${pins.map((p) => tickerFromUrl(p["Screener URL"])).join(", ") || "none"}`);
  writeOut(pins, { lost });
}

// The cohort members to keep tracking, across the newest TRACK_MONTHS months.
// For each month, prefer the captured cohort-entries basket (the locked
// framework holding, matching the app + the live-price feed); fall back to the
// first-snapshot-of-month top-N by composite for a month with no record.
function activeCohortTickers() {
  if (!existsSync(SNAP_DIR)) return new Set();
  const dates = readdirSync(SNAP_DIR)
    .filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .map((f) => f.slice(0, 10))
    .sort();
  if (!dates.length) return new Set();

  // First snapshot of each month, newest TRACK_MONTHS months.
  const firstOfMonth = new Map();
  for (const d of dates) if (!firstOfMonth.has(d.slice(0, 7))) firstOfMonth.set(d.slice(0, 7), d);
  const months = [...firstOfMonth.keys()].sort().slice(-TRACK_MONTHS);

  const cohortEntries = readJson(resolve(DATA_DIR, "cohort-entries.json"));
  const out = new Set();
  for (const m of months) {
    const rec = cohortEntries?.months?.[m];
    if (Array.isArray(rec?.tickers) && rec.tickers.length) {
      rec.tickers.forEach((t) => t && out.add(String(t).toUpperCase()));
      continue;
    }
    const snap = readJson(resolve(SNAP_DIR, `${firstOfMonth.get(m)}.json`));
    const stocks = snap?.stocks || [];
    stocks
      .filter((s) => s.composite != null && s.dataComplete && !s.hardFailed)
      .sort((a, b) => b.composite - a.composite)
      .slice(0, COHORT_SIZE)
      .forEach((s) => s.ticker && out.add(String(s.ticker).toUpperCase()));
  }
  return out;
}

// The universe as of the last commit. In CI this runs after the new file is
// copied into place but before it is committed, so HEAD still has the old one.
function previousUniverse() {
  try {
    const buf = execFileSync("git", ["show", `HEAD:${REPO_REL}`], {
      cwd: resolve(__dirname, ".."), maxBuffer: 64 * 1024 * 1024,
    });
    return JSON.parse(buf.toString("utf8"));
  } catch (err) {
    console.log(`Could not read previous universe from git (${err.message.split("\n")[0]}) — falling back to existing pins only.`);
    return [];
  }
}

function readJson(p) {
  try { return JSON.parse(readFileSync(p, "utf8")); } catch { return null; }
}

function writeOut(companies, extra = {}) {
  writeFileSync(OUT_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    note: "Cohort members being tracked outside the Rs 2,000-12,500 Cr screen. Injected into the universe by scrape-technicals.mjs and write-snapshot.mjs so their prices and scores keep being recorded until the hold window closes.",
    cohort_size: COHORT_SIZE,
    track_months: TRACK_MONTHS,
    count: companies.length,
    ...extra,
    companies,
  }, null, 2) + "\n");
  console.log(`Wrote ${OUT_PATH} (${companies.length} pinned)`);
}
