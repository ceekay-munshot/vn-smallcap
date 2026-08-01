#!/usr/bin/env node
// Offline grid search: replays every parameter combination through the
// SAME backtest engine the Custom Lab uses, ranks by net-of-charges
// return, and prints the top strategies. Used to seed the Custom Lab's
// preset strategies with the historically best performers.
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAP_DIR = resolve(__dirname, "../public/data/snapshots");
const idx = JSON.parse(readFileSync(resolve(SNAP_DIR, "index.json"), "utf8"));
const snapshots = idx.dates.map((d) => JSON.parse(readFileSync(resolve(SNAP_DIR, `${d}.json`), "utf8")));
const lkp = JSON.parse(readFileSync(resolve(__dirname, "../public/data/lkp-manual-picks.json"), "utf8"));
const anchorDate = String(lkp.generated_at).slice(0, 10);   // 2026-06-05

// --- charges (default sim prefs) ---
const SIM = { capital: 500000, bufferPct: 0, brokeragePct: 0.03, sttPct: 0.10, exchangePct: 0.003, gstPct: 18 };
function perSideChargeRate(p) {
  const brok = p.brokeragePct / 100, stt = p.sttPct / 100, exch = p.exchangePct / 100, gst = p.gstPct / 100;
  return brok + stt + exch + (brok + exch) * gst;
}
function daysBetween(a, b) {
  return Math.max(0, Math.round((new Date(b + "T00:00:00Z") - new Date(a + "T00:00:00Z")) / 86400000));
}

// --- faithful port of simulateCustomStrategy (composite-driven params) ---
function simulate(strat) {
  let sorted = snapshots.filter((s) => s.date >= anchorDate).sort((a, b) => a.date.localeCompare(b.date));
  if (!sorted.length) return null;
  const N = strat.basketSize, thr = strat.threshold;
  const targetPct = strat.target / 100, slPct = strat.sl / 100;
  const maxHold = strat.maxHoldDays, rebalDays = strat.rebalanceDays;
  const capital = SIM.capital, bufferAmt = capital * SIM.bufferPct / 100, sideRate = perSideChargeRate(SIM);
  let cash = capital, totalCharges = 0, lastRebal = sorted[0].date, trades = 0;
  const holdings = new Map(); const equity = [];

  for (const snap of sorted) {
    const date = snap.date;
    const closeBy = new Map();
    for (const s of snap.stocks) if (typeof s.close === "number") closeBy.set(s.ticker, s.close);
    const isRebal = date === sorted[0].date || daysBetween(lastRebal, date) >= rebalDays;
    if (isRebal) lastRebal = date;
    const qualifying = snap.stocks
      .filter((s) => s.composite != null && s.composite >= thr && s.dataComplete && !s.hardFailed && typeof s.close === "number" && s.ticker)
      .sort((a, b) => b.composite - a.composite);
    const topN = qualifying.slice(0, N);
    const topNset = new Set(topN.map((s) => s.ticker));
    const exitedToday = new Set();
    for (const [ticker, pos] of [...holdings.entries()]) {
      const px = closeBy.get(ticker); if (typeof px !== "number") continue;
      const gain = px / pos.entryPrice - 1;
      let reason = null;
      if (gain >= targetPct) reason = "T"; else if (gain <= -slPct) reason = "S";
      else if (daysBetween(pos.entryDate, date) >= maxHold) reason = "H";
      else if (isRebal && !topNset.has(ticker)) reason = "R";
      if (!reason) continue;
      const gross = pos.units * px, fee = gross * sideRate;
      cash += gross - fee; totalCharges += fee; trades++; holdings.delete(ticker); exitedToday.add(ticker);
    }
    if (holdings.size < N) {
      for (const s of topN) {
        if (holdings.size >= N) break;
        if (holdings.has(s.ticker) || exitedToday.has(s.ticker)) continue;
        let nav = cash; for (const [t, p] of holdings) nav += p.units * (closeBy.get(t) ?? p.entryPrice);
        const slot = Math.max(0, nav - bufferAmt) / N;
        const buyValue = Math.min(slot, Math.max(0, cash) / (1 + sideRate));
        if (buyValue < 0.01) break;
        const fee = buyValue * sideRate;
        holdings.set(s.ticker, { units: buyValue / s.close, entryDate: date, entryPrice: s.close });
        cash -= buyValue + fee; totalCharges += fee; trades++;
      }
    }
    let value = cash; for (const [t, p] of holdings) value += p.units * (closeBy.get(t) ?? p.entryPrice);
    equity.push(value);
  }
  const finalReturn = (equity[equity.length - 1] / capital - 1) * 100;
  let peak = -Infinity, maxDD = 0;
  for (const v of equity) { if (v > peak) peak = v; const dd = (v / peak - 1) * 100; if (dd < maxDD) maxDD = dd; }
  return { finalReturn, maxDD, trades, charges: totalCharges, holds: holdings.size };
}

const GRID = {
  threshold: [70, 75, 80], basketSize: [5, 7, 10],
  target: [5, 8, 10, 15, 20], sl: [3, 5, 8],
  maxHoldDays: [15, 20, 30], rebalanceDays: [3, 5, 7, 10],
};
const results = [];
for (const threshold of GRID.threshold)
  for (const basketSize of GRID.basketSize)
    for (const target of GRID.target)
      for (const sl of GRID.sl)
        for (const maxHoldDays of GRID.maxHoldDays)
          for (const rebalanceDays of GRID.rebalanceDays) {
            const strat = { threshold, basketSize, target, sl, maxHoldDays, rebalanceDays };
            const r = simulate(strat);
            if (r && r.holds >= 3) results.push({ ...strat, ...r });
          }

// Score: reward return, lightly penalise drawdown (risk-adjusted).
for (const r of results) r.score = r.finalReturn + 0.4 * r.maxDD;   // maxDD is negative
results.sort((a, b) => b.score - a.score);

// Best-in-class across 5 distinct playbook styles — each is the highest
// scorer that fits that style, so the presets are both strong AND varied.
const best = (name, filter) => {
  const pool = results.filter(filter).sort((a, b) => b.score - a.score);
  return pool.length ? { name, ...pool[0] } : null;
};
const STYLES = [
  best("Concentrated — best overall", (r) => true),
  best("Weekly rotation", (r) => r.rebalanceDays === 7),
  best("Quick rotation — fast rebalance", (r) => r.rebalanceDays <= 3),
  best("Broad & steady — wide basket", (r) => r.basketSize >= 10),
  best("Balanced — 5-day rotation", (r) => r.rebalanceDays === 5 && r.basketSize === 7),
].filter(Boolean);
// Drop any style that duplicates another's exact params (keep the first).
const seen = new Set(), picks = [];
for (const s of STYLES) {
  const key = `${s.threshold}-${s.basketSize}-${s.target}-${s.sl}-${s.maxHoldDays}-${s.rebalanceDays}`;
  if (seen.has(key)) continue; seen.add(key); picks.push(s);
}

console.log(`Tested ${results.length} viable combos.\n`);
console.log("== TOP 5 BEST-IN-CLASS STRATEGIES ==");
console.log("  ret%    DD%   trades | thr basket tgt sl hold rebal  | style");
picks.forEach((r, i) => {
  console.log(
    `${i + 1}) ${r.finalReturn.toFixed(2).padStart(6)} ${r.maxDD.toFixed(2).padStart(6)} ${String(r.trades).padStart(6)}  | ${String(r.threshold).padStart(3)} ${String(r.basketSize).padStart(6)} ${String(r.target).padStart(3)} ${String(r.sl).padStart(2)} ${String(r.maxHoldDays).padStart(4)} ${String(r.rebalanceDays).padStart(5)}  | ${r.name}`
  );
});
console.log("\nJSON (for seeding):");
console.log(JSON.stringify(picks.map((r) => ({ name: r.name, threshold: r.threshold, basketSize: r.basketSize, target: r.target, sl: r.sl, maxHoldDays: r.maxHoldDays, rebalanceDays: r.rebalanceDays, _ret: +r.finalReturn.toFixed(2), _dd: +r.maxDD.toFixed(2) })), null, 0));
