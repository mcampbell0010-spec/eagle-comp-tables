"""Comparable companies table for KNSL and its P&C / specialty insurance peers.

Pulls most metrics from Yahoo Finance via yfinance, and ROE (both GAAP and
operating / ex-AOCI) from SEC EDGAR via the local `edgar_metrics` module.
Writes a formatted Excel file (comp_table.xlsx) with:
  - The main comp table, mean/median rows, KNSL conditional highlighting
  - A "Data Notes" sheet documenting sources, freshness, and known caveats.

Re-run any time:
    cd ~/eagle
    source .venv/bin/activate
    python comp-tables/build_comp_table.py
"""

from __future__ import annotations

from collections import namedtuple
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from edgar_metrics import compute_roe

# --- Configuration -----------------------------------------------------------

TICKERS = ["KNSL", "PLMR", "RLI", "WRB", "MKL", "AXS", "ACGL"]
FOCUS = "KNSL"
OUTPUT = Path(__file__).parent / "comp_table.xlsx"

Metric = namedtuple("Metric", "display source key kind higher_is_better")

# source: "yf"    → pulled from yfinance info dict (key is the info field name)
#         "edgar" → pulled from edgar_metrics.RoeResult (key is the attribute)
# higher_is_better=None → neutral, no KNSL coloring
METRICS = [
    Metric("Market Cap ($B)",     "yf",    "marketCap",      "billions", None),
    Metric("Price ($)",           "yf",    "currentPrice",   "price",    None),
    Metric("P/E (TTM)",           "yf",    "trailingPE",     "ratio",    False),
    Metric("P/E (Fwd)",           "yf",    "forwardPE",      "ratio",    False),
    Metric("P/B",                 "edgar", "p_b",            "ratio",    False),
    Metric("P/TBV",               "edgar", "p_tbv",          "ratio",    False),
    Metric("Rev Growth YoY",      "yf",    "revenueGrowth",  "percent",  True),
    Metric("GAAP ROE (TTM)",      "edgar", "gaap_roe",       "percent",  True),
    Metric("Operating ROE (TTM)", "edgar", "operating_roe",  "percent",  True),
    Metric("Net Margin",          "yf",    "profitMargins",  "percent",  True),
    Metric("Div Yield",           "yf",    "dividendYield",  "yield",    True),
]

NUMBER_FORMATS = {
    "billions": '"$"#,##0.0',
    "price":    '"$"#,##0.00',
    "ratio":    "0.00",
    "percent":  "0.0%",
    "yield":    "0.00%",
}

# --- Fetching ----------------------------------------------------------------

def fetch_yf(ticker: str) -> tuple[dict, Optional[float]]:
    """Pull yfinance info dict and resolved current price for one ticker."""
    t = yf.Ticker(ticker)
    info = t.info or {}
    price = info.get("currentPrice") or info.get("regularMarketPrice")
    if price is None:
        hist = t.history(period="1d")
        if not hist.empty:
            price = float(hist["Close"].iloc[-1])
    return info, price


def build_row(ticker: str, info: dict, price: Optional[float],
              edgar_data: dict) -> dict:
    """Return one row {display_label: value} combining yfinance + EDGAR."""
    company = info.get("shortName") or info.get("longName") or ticker
    row = {"Ticker": ticker, "Company": company}
    for m in METRICS:
        if m.source == "yf":
            raw = price if m.key == "currentPrice" else info.get(m.key)
            row[m.display] = _convert(raw, m.kind)
        elif m.source == "edgar":
            row[m.display] = edgar_data.get(m.key)
    return row


def _convert(value, kind):
    if value is None:
        return None
    try:
        v = float(value)
    except (TypeError, ValueError):
        return None
    if kind == "billions":
        return v / 1_000_000_000
    return v


def normalize_yield(series: pd.Series) -> pd.Series:
    """yfinance has flip-flopped on whether dividendYield is a fraction
    (0.012 = 1.2%) or a percentage (1.2 = 1.2%). Normalize to fractions."""
    sample = series.dropna()
    if sample.empty:
        return series
    return series / 100 if sample.median() > 1 else series


# --- Excel styling -----------------------------------------------------------

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TITLE_FONT    = Font(name="Calibri", size=14, bold=True)
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FILL   = PatternFill("solid", fgColor="1F3864")
HEADER_FONT   = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FOCUS_FILL    = PatternFill("solid", fgColor="FFF2CC")
SUMMARY_FILL  = PatternFill("solid", fgColor="E7E6E6")
SUMMARY_FONT  = Font(name="Calibri", size=11, bold=True)

GREEN_FILL  = PatternFill("solid", fgColor="C6EFCE")
GREEN_FONT  = Font(name="Calibri", size=11, bold=True, color="006100")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
YELLOW_FONT = Font(name="Calibri", size=11, bold=True, color="9C5700")
RED_FILL    = PatternFill("solid", fgColor="FFC7CE")
RED_FONT    = Font(name="Calibri", size=11, bold=True, color="9C0006")

NOTES_HEADER_FONT = Font(name="Calibri", size=12, bold=True, color="1F3864")
NOTES_BODY_FONT = Font(name="Calibri", size=11)
NOTES_BODY_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def knsl_cell_style(rank, n, higher_is_better):
    if higher_is_better is None or rank is None:
        return None, None
    top_cut = max(1, n // 3)
    bot_cut = n - max(1, n // 3) + 1
    if rank <= top_cut:
        return GREEN_FILL, GREEN_FONT
    if rank >= bot_cut:
        return RED_FILL, RED_FONT
    return YELLOW_FILL, YELLOW_FONT


# --- Writing the main sheet --------------------------------------------------

def write_comp_sheet(ws, df, mean_row, median_row, knsl_ranks, edgar_period_ends):
    metric_names = [m.display for m in METRICS]
    formats = {m.display: NUMBER_FORMATS[m.kind] for m in METRICS}
    n_cols = 2 + len(metric_names)

    ws.cell(row=1, column=1,
            value="P&C / Specialty Insurance Comparables — KNSL Peer Group").font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    edgar_pe = ", ".join(sorted(set(p for p in edgar_period_ends.values() if p)))
    ws.cell(row=2, column=1,
            value=(f"yfinance data pulled {ts}.  "
                   f"EDGAR ROE columns are TTM through {edgar_pe} (latest 10-Q).")
            ).font = SUBTITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)

    HEADER_ROW = 4
    headers = ["Ticker", "Company"] + metric_names
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=HEADER_ROW, column=col_idx, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    n_peers = len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        r = HEADER_ROW + 1 + i
        is_focus = row["Ticker"] == FOCUS

        ticker_cell = ws.cell(row=r, column=1, value=row["Ticker"])
        company_cell = ws.cell(row=r, column=2, value=row["Company"])
        ticker_cell.border = BORDER
        company_cell.border = BORDER
        if is_focus:
            ticker_cell.font = Font(bold=True)
            company_cell.font = Font(bold=True)
            ticker_cell.fill = FOCUS_FILL
            company_cell.fill = FOCUS_FILL

        for j, m in enumerate(METRICS):
            cell = ws.cell(row=r, column=3 + j, value=row[m.display])
            cell.number_format = formats[m.display]
            cell.alignment = Alignment(horizontal="right")
            cell.border = BORDER

            if is_focus:
                rank = knsl_ranks.get(m.display)
                fill, font = knsl_cell_style(rank, n_peers, m.higher_is_better)
                if fill is not None:
                    cell.fill = fill
                    cell.font = font
                else:
                    cell.fill = FOCUS_FILL
                    cell.font = Font(bold=True)

    summary_start = HEADER_ROW + 1 + n_peers + 1
    for offset, (label, source) in enumerate([("Mean", mean_row), ("Median", median_row)]):
        r = summary_start + offset
        ws.cell(row=r, column=1, value=label).font = SUMMARY_FONT
        ws.cell(row=r, column=2, value=f"Peer {label}").font = SUMMARY_FONT
        for col in (1, 2):
            ws.cell(row=r, column=col).fill = SUMMARY_FILL
            ws.cell(row=r, column=col).border = BORDER
        for j, m in enumerate(METRICS):
            cell = ws.cell(row=r, column=3 + j, value=source[m.display])
            cell.number_format = formats[m.display]
            cell.alignment = Alignment(horizontal="right")
            cell.fill = SUMMARY_FILL
            cell.font = SUMMARY_FONT
            cell.border = BORDER

    key_row = summary_start + 3
    ws.cell(row=key_row, column=1,
            value=("Highlighting on KNSL row: green = top third of peers, "
                   "yellow = middle, red = bottom third. Direction is metric-aware "
                   "(lower P/E better; higher ROE better). "
                   "See Data Notes tab for sources & caveats.")
            ).font = SUBTITLE_FONT
    ws.merge_cells(start_row=key_row, start_column=1, end_row=key_row, end_column=n_cols)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 32
    for j in range(len(metric_names)):
        ws.column_dimensions[get_column_letter(3 + j)].width = 16
    ws.row_dimensions[HEADER_ROW].height = 32
    ws.freeze_panes = "C5"


# --- Data Notes sheet --------------------------------------------------------

NOTES_SECTIONS = [
    ("Data sources by column",
     "Each column in the comp table comes from one of two sources. "
     "Yahoo Finance (via the yfinance Python library) is convenient but "
     "opaque — Yahoo computes the ratio for you and you can't fully audit "
     "the inputs. SEC EDGAR XBRL company-facts are the primary-source "
     "filings; the script computes the ratio locally and you can re-derive "
     "every step.\n\n"
     "  Market Cap, Price                 → Yahoo Finance\n"
     "  P/E (TTM), P/E (Fwd)              → Yahoo Finance\n"
     "  P/B                               → SEC EDGAR (computed locally)\n"
     "  P/TBV                             → SEC EDGAR (computed locally)\n"
     "  Revenue Growth YoY, Net Margin    → Yahoo Finance\n"
     "  Dividend Yield                    → Yahoo Finance\n"
     "  GAAP ROE (TTM)                    → SEC EDGAR (computed locally)\n"
     "  Operating ROE (TTM, ex-AOCI)      → SEC EDGAR (computed locally)\n"
     "\n"
     "Note on P/B: this is computed from EDGAR's StockholdersEquity (TOTAL "
     "equity at the latest 10-Q period end), not Yahoo's bookValue field. "
     "The two agree closely for most names but differ by a few percent "
     "for issuers with preferred stock — Yahoo appears to exclude preferred, "
     "EDGAR includes it. ACGL is the relevant case in this peer set. "
     "Sourcing both P/B and P/TBV from EDGAR keeps their denominators "
     "consistent so the relationship is interpretable."
     ),

    ("Data freshness",
     "  • Yahoo intraday fields (price, market cap, dividend yield "
     "denominator): ~15-minute delay during US market hours.\n"
     "  • Yahoo-derived fundamentals (TTM EPS, book value, ROE, margins): "
     "update only when a new 10-Q/10-K is filed, typically 4–8 weeks after "
     "quarter-end. The cell is current, but the underlying value reflects "
     "a balance sheet that can be up to ~75 days old.\n"
     "  • Forward P/E: updates continuously as sell-side analysts revise "
     "estimates, but coverage depth varies (see caveat below).\n"
     "  • EDGAR ROE columns: also update only on new 10-Q/10-K. The "
     "timestamp at the top of the Comp Table sheet says when this file "
     "was generated, NOT when the underlying data was last reported."
     ),

    ("P/TBV methodology",
     "P/TBV (Price to Tangible Book Value) strips goodwill and other "
     "intangibles out of GAAP equity. For insurance companies it's the "
     "cleaner valuation lens because tangible equity is what actually "
     "backs underwriting risk — goodwill and intangibles cannot absorb "
     "losses.\n\n"
     "    Tangible book = Stockholders' Equity − Goodwill − Intangibles\n"
     "    P/TBV         = Market Cap / Tangible book\n\n"
     "EDGAR concepts used (all from us-gaap taxonomy, balance-sheet date = "
     "most recent 10-Q period end):\n"
     "  • Equity:    StockholdersEquity\n"
     "  • Goodwill:  Goodwill (treated as $0 if tag absent)\n"
     "  • Intangibles: IntangibleAssetsNetExcludingGoodwill,\n"
     "                 with fallback to FiniteLivedIntangibleAssetsNet\n"
     "                 or OtherIntangibleAssetsNet ($0 if all absent)\n\n"
     "Note: insurance-specific items like DAC (Deferred Acquisition "
     "Costs) and VOBA (Value of Business Acquired) are NOT treated as "
     "intangibles here — they're operational assets, not goodwill-style "
     "accounting overhang. Standard insurance-analyst convention.\n\n"
     "Practical reading: when a name's P/TBV is materially HIGHER than "
     "its P/B, that name has been built partly via acquisition (goodwill "
     "on the balance sheet). MKL and ACGL show this clearly. KNSL, PLMR, "
     "and RLI have negligible goodwill so P/B ≈ P/TBV."
     ),

    ("Caveat 2 — Forward P/E coverage thinness",
     "Forward P/E uses the consensus analyst EPS estimate for the next 12 "
     "months. The number is only as reliable as the consensus behind it.\n\n"
     "  • WRB, MKL, ACGL: deep sell-side coverage. Consensus is robust.\n"
     "  • KNSL, AXS: moderate coverage. One large revision can shift the "
     "consensus.\n"
     "  • PLMR: thin coverage. The Fwd P/E here is more analyst-opinion "
     "than analyst-consensus.\n\n"
     "Practical implication: don't use Fwd P/E in isolation for ranking. "
     "Always check it against TTM P/E and the growth trajectory."
     ),

    ("Caveat 3 — One-quarter revenue growth is noisy",
     "The Revenue Growth YoY column compares the SINGLE most recent "
     "quarter to the same quarter one year ago. That can swing widely "
     "with:\n"
     "  • Timing of premium written (lumpy specialty lines)\n"
     "  • Large CAT losses (affect net earned premium via reinstatement)\n"
     "  • One-off accounting items (reserve releases, deferred acquisition cost)\n\n"
     "For trend reading, prefer TTM revenue growth (4-quarter average) or "
     "look at the 3-year CAGR. PLMR's 60%+ shown here reflects rapid "
     "scaling of a small specialty book and is not directly comparable "
     "to mature carriers' single-digit growth."
     ),

    ("Caveat 4 — TTM XBRL quirk (the standalone-Q4 trap)",
     "*** READ THIS BEFORE EXTENDING THE EDGAR PIPELINE. ***\n\n"
     "SEC filers do NOT tag standalone Q4 as a 3-month period in XBRL. "
     "Q4 only appears inside the 10-K's full-year (12-month) line. So the "
     "naive approach of 'sum the last 4 quarterly entries from EDGAR' "
     "SILENTLY SKIPS Q4 and gives you the wrong TTM by one quarter.\n\n"
     "The correct construction is:\n\n"
     "    TTM = most recent FY value\n"
     "          + latest quarter past FY-end\n"
     "          − prior-year same quarter\n\n"
     "Example for KNSL today (period through 2026-03-31):\n"
     "    TTM NI = FY2025 ($503.6M) + Q1'26 ($112.6M) − Q1'25 ($89.2M)\n"
     "           = $526.9M\n\n"
     "This is what edgar_metrics.ttm() implements. If you ever see EDGAR-"
     "derived TTM figures ~10% off Yahoo's, the standalone-Q4 trap is the "
     "first thing to check."
     ),

    ("Caveat 5 — GAAP ROE vs Operating ROE (AOCI adjustment)",
     "AOCI = Accumulated Other Comprehensive Income (net of tax). It sits "
     "in shareholders' equity but reflects unrealized gains/losses on the "
     "investment portfolio that have not flowed through net income.\n\n"
     "When AOCI is negative (as for all P&C insurers post-2022 rate hikes), "
     "the GAAP equity denominator is SMALLER than the 'true' capital "
     "deployed in the business. That mechanically INFLATES GAAP ROE.\n\n"
     "Operating ROE strips AOCI from the denominator:\n"
     "    Operating ROE = NI TTM / avg(stockholders' equity − AOCI)\n\n"
     "For carriers running smaller investment books relative to equity "
     "(KNSL, PLMR, AXS), the adjustment is small (-0.1 to -0.3 pp).\n"
     "For carriers with larger fixed-income portfolios (RLI, WRB), the "
     "adjustment is meaningfully larger (-1.1 to -1.4 pp) and can change "
     "the relative ranking — see the rank-change in the file generation log."
     ),

    ("Methodology — equity averaging",
     "ROE denominators use a 5-POINT AVERAGE: the most recent quarter-end "
     "equity plus the previous 4 quarter-end values. A simple "
     "(now + one-year-ago) / 2 average can miss capital raises, buybacks, "
     "or dividend timing within the year. Yahoo appears to use a 2-point "
     "average, which is why our EDGAR GAAP ROE is typically 0.5–2 pp lower "
     "than yfinance's reported ROE — the 5-point captures more of the "
     "equity build."
     ),

    ("When to NOT trust these numbers",
     "  • Right after a major event (CAT loss, M&A, secondary offering): "
     "the TTM-based ratios reflect a 'before' world. Wait for the next "
     "10-Q or model the event manually.\n"
     "  • When MKL or other diversified holdcos show outlier metrics: "
     "Markel's reported results are heavily influenced by Markel Ventures "
     "(non-insurance) and mark-to-market equity investments. Net Margin "
     "and ROE here are NOT clean insurance-underwriting metrics for MKL.\n"
     "  • When the FY is brand new (script run in January–February): "
     "the 'latest FY' for TTM construction may be last year's because "
     "the new 10-K hasn't filed yet. Check the EDGAR period-end at the "
     "top of the Comp Table sheet.\n"
     "  • For any number that drives a recommendation: re-pull from the "
     "actual 10-Q PDF or your firm's data feed. Yahoo is fine for "
     "screening; not for published notes."
     ),

    ("How to regenerate this file",
     "    cd ~/eagle\n"
     "    source .venv/bin/activate\n"
     "    python comp-tables/build_comp_table.py\n\n"
     "Edit comp-tables/build_comp_table.py to change the peer group "
     "(TICKERS) or metrics (METRICS list at the top of the file). "
     "edgar_metrics.py is a reusable module — verify_against_edgar.py "
     "uses the same helpers for stand-alone tie-out reports."
     ),
]


def write_notes_sheet(ws):
    ws.column_dimensions["A"].width = 110
    ws.cell(row=1, column=1, value="Data Notes & Limitations").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value=("Use this tab to remember what's authoritative, what's "
                   "convenient-but-flawed, and the traps you've already "
                   "stepped on once.")
            ).font = SUBTITLE_FONT

    row = 4
    for header, body in NOTES_SECTIONS:
        c = ws.cell(row=row, column=1, value=header)
        c.font = NOTES_HEADER_FONT
        row += 1
        body_cell = ws.cell(row=row, column=1, value=body)
        body_cell.font = NOTES_BODY_FONT
        body_cell.alignment = NOTES_BODY_WRAP
        # Crude row-height heuristic: ~15px per visual line, ~95 chars/line.
        line_count = sum(max(1, len(line) // 95 + 1) for line in body.split("\n"))
        ws.row_dimensions[row].height = max(30, 15 * line_count + 8)
        row += 2  # spacer


# --- Orchestration -----------------------------------------------------------

def main():
    print(f"Fetching {len(TICKERS)} tickers from Yahoo Finance and SEC EDGAR…")
    edgar_data = {}
    edgar_period_ends = {}
    rows = []
    for t in TICKERS:
        info, price = fetch_yf(t)
        market_cap = info.get("marketCap")
        try:
            r = compute_roe(t)
            edgar_period_ends[t] = r.period_end
            if r.notes:
                print(f"  {t}: notes — {'; '.join(r.notes)}")
            p_b = (market_cap / r.book_value
                   if (market_cap and r.book_value and r.book_value > 0)
                   else None)
            p_tbv = (market_cap / r.tangible_book
                     if (market_cap and r.tangible_book and r.tangible_book > 0)
                     else None)
            edgar_data[t] = {
                "gaap_roe": r.gaap_roe,
                "operating_roe": r.operating_roe,
                "p_b": p_b,
                "p_tbv": p_tbv,
                "_book_value": r.book_value,
                "_tangible_book": r.tangible_book,
                "_goodwill": r.goodwill,
                "_intangibles": r.intangibles,
            }
        except Exception as e:
            print(f"  {t}: EDGAR fetch failed — {type(e).__name__}: {e}")
            edgar_data[t] = {}
            edgar_period_ends[t] = None
        rows.append(build_row(t, info, price, edgar_data[t]))

    df = pd.DataFrame(rows)
    df["Div Yield"] = normalize_yield(df["Div Yield"])

    metric_names = [m.display for m in METRICS]
    mean_row   = {name: df[name].mean()   for name in metric_names}
    median_row = {name: df[name].median() for name in metric_names}

    knsl_idx = df.index[df["Ticker"] == FOCUS][0]
    knsl_ranks = {}
    for m in METRICS:
        if m.higher_is_better is None:
            continue
        ranks = df[m.display].rank(method="min", ascending=not m.higher_is_better)
        v = ranks.iloc[knsl_idx]
        knsl_ranks[m.display] = None if pd.isna(v) else int(v)

    wb = Workbook()
    comp_ws = wb.active
    comp_ws.title = "Comp Table"
    write_comp_sheet(comp_ws, df, mean_row, median_row, knsl_ranks, edgar_period_ends)
    notes_ws = wb.create_sheet("Data Notes")
    write_notes_sheet(notes_ws)
    wb.save(OUTPUT)
    print(f"Wrote {OUTPUT}")

    # Console preview
    print("\nPreview:\n")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 220)
    def fmt(v):
        if v is None or pd.isna(v):
            return "—"
        return f"{v:,.3f}" if abs(v) < 100 else f"{v:,.1f}"
    print(df[["Ticker", "Company"] + metric_names].to_string(
        index=False, na_rep="—", formatters={c: fmt for c in metric_names}))

    print("\nKNSL rank vs peers (1 = best):")
    for name, rank in knsl_ranks.items():
        n = df[name].notna().sum()
        print(f"  {name:<22} {rank}/{n}")


if __name__ == "__main__":
    main()
