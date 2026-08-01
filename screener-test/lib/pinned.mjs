// Pinned holdings — companies we keep tracking after they leave the screen.
//
// The universe is a market-cap band (Rs 2,000-12,500 Cr). A cohort pick that
// performs well can cross the ceiling mid-hold; one that falls apart can drop
// through the floor. Either way the next weekly refresh drops it from
// screener-companies.json, the technicals scrape stops fetching it, and it
// vanishes from every subsequent snapshot -- so the cohort tracker silently
// loses the position part-way through the month it is supposed to be measuring.
//
// That is not a rounding error. Leakage is directional at both ends: winners
// exit through the ceiling and losers through the floor, so measured cohort
// performance is truncated on both tails by exactly the moves that matter most.
//
// A pin freezes the company's last known fundamentals row and re-injects it
// into the universe for the scrapers and the snapshot writer, so price and
// score keep being recorded until the hold window closes.

import { readFileSync } from "node:fs";
import { resolve } from "node:path";

export const PINNED_FILENAME = "pinned-companies.json";

export const tickerFromUrl = (url) => {
  const m = String(url || "").match(/\/company\/([^/]+)/);
  return m ? m[1].toUpperCase() : null;
};

export function readPinned(dataDir) {
  try {
    const j = JSON.parse(readFileSync(resolve(dataDir, PINNED_FILENAME), "utf8"));
    return Array.isArray(j?.companies) ? j.companies : [];
  } catch {
    return [];   // absent on a fresh repo, and before any cohort exists
  }
}

// Append pinned rows that are not already in the live universe. A pin whose
// company has come BACK inside the band is skipped, not duplicated -- the
// live row is always preferred because its fundamentals are current.
export function mergePinned(companies, dataDir) {
  const pinned = readPinned(dataDir);
  if (!pinned.length) return { companies, added: 0, addedTickers: [] };

  const present = new Set(
    (companies || []).map((c) => tickerFromUrl(c["Screener URL"])).filter(Boolean),
  );
  const add = pinned.filter((p) => {
    const t = tickerFromUrl(p["Screener URL"]);
    return t && !present.has(t);
  });
  return {
    companies: [...(companies || []), ...add],
    added: add.length,
    addedTickers: add.map((p) => tickerFromUrl(p["Screener URL"])),
  };
}
