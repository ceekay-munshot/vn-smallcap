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

| | Adds | 2-yr net | CAGR | Worst fall |
|---|---|---|---|---|
| 1 Core | top 7, equal money | +22.5% | 10.7% | −26.0% |
| 2 Equal risk | inverse-ATR sizing | +23.5% | 11.1% | −25.5% |
| 3 Spread | 10 names instead of 7 | +28.3% | 13.3% | −20.2% |
| 4 Balanced | max 2 per sector | **+34.1%** | **15.8%** | **−19.6%** |
| 5 Consensus | conviction weighting | +24.9% | 11.8% | −19.2% |

Every layer controls risk rather than picking better stocks, because that is
what the score does not already do. Filters that repeat the score (MACD,
relative strength, above-50-DMA) were tested and dropped — they reject
high-scoring names and push the basket deeper down the ranking for no gain.

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
