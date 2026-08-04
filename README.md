# vn-smallcap

Fundamentals + technicals scoring dashboard for the **₹2,000–12,500 Cr NSE universe**
(724 companies as of Aug 2026).

Static site — no build step. `public/` is served directly by Cloudflare Workers.
GitHub Actions run the scrapers and commit refreshed data back to `main`.

## Scoring

Two pillars, weighted composite:

| Pillar | Weight | Rules |
|---|---|---|
| Fundamentals | 30% | 16 |
| Technicals | 70% | 16 |

Macro, Sentiment and Liquidity pillars are intentionally excluded.
ADTV is retained as a standalone tradability gate (hard fail below ₹1 Cr).

## Universe

Screener.in saved screen — market cap ≥ ₹2,000 Cr and ≤ ₹12,500 Cr,
plus a broad quality filter (any one of 24 fundamental conditions).

## Strategies

Five plans over one ranking, defined in `public/data/strategies.json`. Each
adds a single layer to the one before it, and the Compare tab is there to test
whether that helped — not to assert it.

All rebalance **monthly** — buy on the first trading day, hold to the first of
next month. 2-year backtest, net of charges, technicals only:

| | Adds | 2-yr net | CAGR | Worst fall |
|---|---|---|---|---|
| 1 Core | top 7, equal money (control) | +56.4% | 25.0% | −15.9% |
| 2 Equal risk **← what runs today** | inverse-ATR sizing | +66.6% | 29.1% | −13.6% |
| 3 Spread | 10 names instead of 7 | +56.7% | 25.2% | −12.4% |
| 4 Balanced | max 2 per sector | +43.1% | 19.6% | −14.1% |
| 5 Consensus | only names all 4 agree on, full weight | **+79.5%** | **34.0%** | **−12.2%** |

These are five *approaches*, not a ladder where each beats the last — with ~24
rebalances on 7 concentrated names, the numbers are very sensitive to timing and
should be read as a rehearsal, not a record.

Every layer controls risk rather than picking better stocks, because that is
what the score does not already do. Filters that repeat the score (MACD,
relative strength, above-50-DMA) were tested and dropped — they reject
high-scoring names and push the basket deeper down the ranking for no gain.

**#1 is Equal risk** — the plan already being run, and the backtest leader too.
A backtest is still a rehearsal; the live track is the real test.

Backtest: `node screener-test/backtest-strategies.mjs`, replayed over
`history-technical.json` (504 trading days, 620 tickers) and written to
`strategy-backtest.json`. Technicals only — the history carries no
point-in-time fundamentals — and the universe is today's names replayed
backwards, so survivorship flatters every figure.

## Charges

`public/js/charges.js`, verified against the Zerodha calculator. STT and the
exchange/SEBI fees hit both sides, stamp duty only the buy, and the DP fee is
a flat ₹15.34 per scrip sold. That flat fee is the one that matters: it is
0.43% of a ₹3,500 position and 0.11% of a ₹14,000 one, so cost per trade
depends on ticket size, and turnover is the largest controllable drag at a
small account.

## Tracking

Top 7 (or 10) by composite on the first trading day of each month, held for the
month, measured against Nifty Smallcap 250 / Nifty Midcap 150. Three bench
names are sized alongside: if a pick gaps past its skip price on entry morning,
the Entry Monitor says so and names the replacement.

## Status

🚧 Under construction — data pipeline not yet live.
