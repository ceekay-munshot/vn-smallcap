#!/usr/bin/env node
// Offline runner for the five strategies. Replays each one over the full
// technical history and prints the comparison table -- the same numbers
// the Compare tab shows, so a regression here is visible without opening
// a browser. Also writes public/data/strategy-backtest.json, which the
// dashboard loads instead of recomputing 500 days x 620 tickers in the
// page (that took several seconds and blocked the first paint).
//
// Usage: node screener-test/backtest-strategies.mjs [--capital 50000]

import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { setStrategies, buildBacktestContext, runAll, buildUniverseCurve } from "../public/js/strategies.js";
import { CHARGE_DEFAULTS, chargeSummary } from "../public/js/charges.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA = resolve(__dirname, "../public/data");

const argCapital = Number((process.argv.find((a) => a.startsWith("--capital=")) || "").split("=")[1]);
const CAPITAL = Number.isFinite(argCapital) && argCapital > 0 ? argCapital : 50000;

const defs = JSON.parse(readFileSync(resolve(DATA, "strategies.json"), "utf8"));
setStrategies(defs);

const hist = JSON.parse(readFileSync(resolve(DATA, "history-technical.json"), "utf8"));
let benchmark = null;
try { benchmark = JSON.parse(readFileSync(resolve(DATA, "benchmark-history.json"), "utf8")); } catch {}

console.log(`Building context — ${hist.dates.length} days x ${Object.keys(hist.tickers).length} tickers…`);
const t0 = Date.now();
const ctx = buildBacktestContext(hist, benchmark);
console.log(`Context built in ${((Date.now() - t0) / 1000).toFixed(1)}s`);

const runs = runAll(ctx, CHARGE_DEFAULTS, { capital: CAPITAL, rebalanceDays: 21 });
const universe = buildUniverseCurve(ctx);

const cs = chargeSummary(CAPITAL, 7, CHARGE_DEFAULTS);
console.log(`\nCapital ₹${CAPITAL.toLocaleString("en-IN")} / 7 stocks = ₹${Math.round(cs.perPosition).toLocaleString("en-IN")} a slot`);
console.log(`Round trip ${cs.roundTripPct.toFixed(3)}% (DP fee is ${cs.dpSharePct.toFixed(0)}% of it) · monthly rotation = ${cs.monthlyRotationAnnualPct.toFixed(1)}%/yr\n`);

const pad = (s, n) => String(s).padStart(n);
console.log("  strategy      net%   CAGR%    DD%   vol%  win%  trades  chg%  expo%");
for (const r of runs) {
  const m = r.metrics;
  if (!m) { console.log(`  ${r.strategy.name} — no result`); continue; }
  console.log(
    `  ${r.strategy.name.padEnd(11)} ${pad(m.netReturn.toFixed(1), 6)} ${pad(m.cagr.toFixed(1), 6)} ` +
    `${pad(m.maxDD.toFixed(1), 6)} ${pad(m.vol.toFixed(1), 6)} ${pad((m.winRate ?? 0).toFixed(0), 5)} ` +
    `${pad(m.trades, 7)} ${pad(m.chargesPct.toFixed(1), 5)} ${pad((m.exposurePct ?? 0).toFixed(0), 6)}`
  );
}
const uni = universe[universe.length - 1].retPct;
console.log(`  ${"Universe".padEnd(11)} ${pad(uni.toFixed(1), 6)}   (equal-weight, buy and hold)\n`);

// Compact payload for the dashboard — curves are thinned to one point a
// week so the file stays small; full metrics and trades are kept.
const thin = (curve) => curve.filter((_, i) => i % 5 === 0 || i === curve.length - 1)
  .map((p) => ({ d: p.date, r: +p.retPct.toFixed(3) }));

const payload = {
  generated_at: new Date().toISOString(),
  capital: CAPITAL,
  rebalanceDays: 21,
  charges: CHARGE_DEFAULTS,
  window: { start: ctx.dates[0], end: ctx.dates[ctx.dates.length - 1], days: ctx.dates.length },
  caveats: [
    "Universe is today's 620 covered names replayed backwards — delisted and dropped-out companies are absent (survivorship bias).",
    "History carries no volume, so the liquidity gate is live-only and not backtested.",
    "ATR is approximated from close-to-close moves; a true ATR would be slightly wider.",
  ],
  universeCurve: thin(universe),
  strategies: runs.map((r) => ({
    id: r.strategy.id,
    name: r.strategy.name,
    curve: thin(r.result.curve),
    grossReturn: r.grossReturn,
    metrics: r.metrics,
    trades: r.result.trades
      .slice()
      .sort((a, b) => b.retPct - a.retPct)
      .filter((_, i, arr) => i < 5 || i >= arr.length - 5)      // best 5 + worst 5
      .map((t) => ({ ...t, retPct: +t.retPct.toFixed(2) })),
  })),
};
writeFileSync(resolve(DATA, "strategy-backtest.json"), JSON.stringify(payload) + "\n");
console.log(`Wrote ${resolve(DATA, "strategy-backtest.json")}`);
