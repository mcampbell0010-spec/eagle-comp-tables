# Comp Tables — P&C / Specialty Insurance Comparables

A small Python tool that pulls live market data from Yahoo Finance and primary-source fundamentals from SEC EDGAR XBRL filings, then writes a formatted Excel comparables table for a focus name (KNSL) versus a peer group of P&C and specialty insurers.

Built as a learning project by a financial analyst getting hands-on with Python — code prioritizes readability and analyst-relevant methodology over cleverness.

## What it does

For each ticker in the peer set, the tool pulls:

- **Market data** (Yahoo Finance): market cap, price, P/E TTM and Forward, revenue growth, net margin, dividend yield
- **Valuation** (SEC EDGAR): P/B and P/TBV computed locally from `StockholdersEquity` minus goodwill and intangibles
- **Returns** (SEC EDGAR): GAAP ROE *and* Operating ROE (ex-AOCI), TTM net income over a 5-point average equity denominator

It writes `comp_table.xlsx` with two sheets:

1. **Comp Table** — the main table, with mean and median rows, and the focus name's row highlighted by quartile vs. peers (green = top third, yellow = middle, red = bottom third)
2. **Data Notes** — every source, freshness tier, known caveat, and known trap (the standalone-Q4 XBRL quirk that bit me once and will bite future-me again)

## Sample output

Generated on the date in the file header — re-run any time to refresh.

| Ticker | Mkt Cap ($B) | P/E (TTM) | P/E (Fwd) | P/B | P/TBV | Rev Gr | GAAP ROE | Op ROE | Net Mgn | Div Yld |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| **KNSL** | 7.2 | 13.7 | 14.3 | 3.64 | 3.64 | +10.2% | 29.0% | 28.2% | 27.5% | 0.32% |
| PLMR | 3.1 | 16.1 | 10.3 | 3.18 | 3.18 | +59.7% | 22.3% | 22.0% | 20.1% | — |
| RLI | 5.0 | 12.6 | 19.5 | 2.77 | 2.84 | +4.0% | 22.5% | 21.1% | 20.8% | 1.33% |
| WRB | 25.1 | 14.3 | 14.0 | 2.58 | 2.63 | +4.0% | 19.8% | 18.6% | 12.6% | 0.59% |
| MKL | 22.9 | 13.2 | 15.0 | 1.26 | 1.66 | -16.9% | 9.9% | 9.8% | 11.1% | — |
| AXS | 7.3 | 7.4 | 6.9 | 1.15 | 1.19 | +8.0% | 17.2% | 17.0% | 16.0% | 1.77% |
| ACGL | 31.8 | 7.0 | 9.2 | 1.32 | 1.32 | -3.3% | 20.9% | 20.7% | 24.6% | — |

A sample workbook (`comp_table.xlsx`) is committed for quick inspection without having to run the tool first.

## Key findings (current peer-group screen)

**KNSL is the unambiguous quality leader.** Best ROE in the group on both reported (29.0%) and AOCI-adjusted (28.2%) measures — six points clear of the runner-up. Best net margin (27.5%) and second-best revenue growth (10.2%, behind only PLMR's outlier 60% which reflects a different stage of the lifecycle). The market prices this premium quality: KNSL is the most expensive name on every valuation metric (P/E, P/B, P/TBV).

**The AOCI / Operating ROE adjustment changes one ranking.** Stripping accumulated other comprehensive income (mostly unrealized bond-portfolio losses) from the equity denominator:

- Hurts **RLI** the most: −1.4 pp, enough to swap with **PLMR** for the #2 ROE rank.
- Modest impact on others (-0.1 to -1.1 pp).
- **KNSL stays #1 by a wide margin** — the quality story is robust to the lens.

**P/TBV reveals MKL's acquisition history.** Markel Group's P/B of 1.26 jumps to a P/TBV of 1.66 — a 31% gap that reflects roughly a quarter of its reported book value sitting in goodwill and intangibles from years of acquisitions. On a tangible-book basis MKL is no longer the second-cheapest name; ACGL takes that slot. KNSL, PLMR, and ACGL all show P/B ≈ P/TBV, indicating organically-built equity.

## Limitations

A short list — the workbook's **Data Notes** sheet has the full version.

- Yahoo's free feed is delayed ~15 minutes during US market hours.
- Forward P/E reliability varies with sell-side coverage depth. KNSL and especially PLMR have thinner coverage; the Fwd P/E is more sensitive to individual analyst revisions.
- "Revenue Growth YoY" compares one quarter to the same quarter prior year. Specialty insurance lines are lumpy; treat this as a noisy point estimate, not a trend.
- Net margin and ROE for diversified holding companies (MKL especially) include non-insurance earnings — not directly comparable to pure-play carriers' underwriting metrics.
- The XBRL standalone-Q4 trap: SEC filers don't tag Q4 as a standalone 3-month period; it lives inside the 10-K's full-year line only. The "sum the last 4 quarterly entries" approach silently miscounts. This tool uses the correct `FY + latest Q − prior-year Q` construction (see `edgar_metrics.py:ttm()` and Caveat 4 in the Data Notes sheet).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install yfinance pandas requests openpyxl
```

**Important — SEC EDGAR contact info.** SEC requires identifiable contact info on every EDGAR API request. Set this once via a `.env` file:

```bash
cp .env.example .env
# Edit .env and set SEC_USER_AGENT to your name + email, e.g.:
#   SEC_USER_AGENT=Jane Smith jane@example.com
```

`.env` is gitignored. The script will raise a clear error if `SEC_USER_AGENT` is missing.

## Usage

```bash
source .venv/bin/activate
python build_comp_table.py        # generates comp_table.xlsx
python verify_against_edgar.py    # standalone EDGAR vs Yahoo tie-out report
```

Edit `TICKERS` at the top of either script to change the peer set. Edit `METRICS` in `build_comp_table.py` to change the columns.

## File structure

```
comp-tables/
├── README.md                    # this file
├── .env.example                 # template — copy to .env and fill in your SEC contact info
├── .gitignore
├── build_comp_table.py          # orchestration, Excel writer, Data Notes sheet content
├── edgar_metrics.py             # reusable XBRL data layer (TTM, 5-point average, ROE, tangible book)
├── verify_against_edgar.py      # standalone EDGAR audit / tie-out report
└── comp_table.xlsx              # sample output, regenerated on each run
```

## Notes on methodology

- **TTM construction:** `most recent FY + latest quarter past FY-end − prior-year same quarter`. Avoids the standalone-Q4 XBRL trap.
- **Equity averaging for ROE:** 5-point average of `[period_end, period_end − 3mo, − 6mo, − 9mo, − 12mo]` rather than the simpler 2-point `(now + 1y ago) / 2`. The 5-point captures intra-year equity build more accurately and is the convention insurance analysts use.
- **Operating ROE:** `NI TTM / avg(StockholdersEquity − AOCI)`. Strips out unrealized portfolio gains/losses so the denominator reflects "true" operating capital.
- **Tangible book:** `StockholdersEquity − Goodwill − IntangibleAssetsNetExcludingGoodwill` (with fallback intangible concepts). DAC/VOBA are not stripped — standard insurance-analyst convention treats them as operational assets.
