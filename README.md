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

## Tracking

Top 7 by composite on the first trading day of each month, held for the month,
measured against Nifty Smallcap 250 / Nifty Midcap 150.

## Status

🚧 Under construction — data pipeline not yet live.
