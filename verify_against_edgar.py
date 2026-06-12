"""Compute EDGAR-derived GAAP and Operating ROE for the full peer set, and
compare against what yfinance reports.

Uses the corrected TTM construction (FY + latest Q − prior-yr Q) and a 5-point
average equity (this Q-end + previous 4 Q-ends). Operating ROE strips AOCI
from the denominator.
"""

from __future__ import annotations

import yfinance as yf

from edgar_metrics import USER_AGENT, compute_roe

TICKERS = ["KNSL", "PLMR", "RLI", "WRB", "MKL", "AXS", "ACGL"]
FOCUS = "KNSL"


def yf_roe(ticker: str):
    info = yf.Ticker(ticker).info or {}
    return info.get("returnOnEquity")


def pct(a, b):
    if a is None or b is None or b == 0:
        return None
    return (a - b) / b


def flag(diff, tol=0.02):
    if diff is None:
        return "n/a"
    return f"[{'WARN' if abs(diff) > tol else 'OK  '}] {diff*100:+.2f}%"


def print_detail(r, yfr):
    print(f"\n{'='*78}\n{r.ticker} — {r.company}\n{'='*78}")
    print(f"  Period end: {r.period_end}")
    if r.ni_method.get("method") == "FY + latest Q − prior-yr Q":
        m = r.ni_method
        print(f"  NI TTM = FY ({m['fy_end']}) ${m['fy_val']/1e6:,.1f}M"
              f"  +  Q ({m['lq_end']}) ${m['lq_val']/1e6:,.1f}M"
              f"  −  Q ({m['pq_end']}) ${m['pq_val']/1e6:,.1f}M"
              f"  =  ${r.ni_ttm/1e6:,.1f}M")
    elif r.ni_ttm:
        print(f"  NI TTM: ${r.ni_ttm/1e6:,.1f}M  ({r.ni_method.get('method')})")

    if r.avg_equity is not None:
        print(f"  Avg equity ({r.avg_equity_n_points}-point):       "
              f"${r.avg_equity/1e9:.3f}B")
    if r.avg_aoci is not None and r.aoci_concept:
        print(f"  Avg AOCI ({r.avg_aoci_n_points}-point):           "
              f"${r.avg_aoci/1e6:+,.1f}M")
    if r.avg_equity_ex_aoci is not None:
        print(f"  Avg equity ex-AOCI:        ${r.avg_equity_ex_aoci/1e9:.3f}B")

    if r.gaap_roe is not None:
        print(f"  GAAP ROE:        {r.gaap_roe*100:.2f}%", end="")
        if yfr is not None:
            print(f"   vs yfinance {yfr*100:.2f}%   {flag(pct(yfr, r.gaap_roe))}")
        else:
            print()
    if r.operating_roe is not None and r.gaap_roe is not None:
        gap_pp = (r.operating_roe - r.gaap_roe) * 100
        print(f"  Operating ROE:   {r.operating_roe*100:.2f}%   ({gap_pp:+.2f} pp vs GAAP)")

    if r.notes:
        print("  Notes:")
        for n in r.notes:
            print(f"    - {n}")


def print_rankings(results):
    print(f"\n{'='*78}\nROE RANKINGS — GAAP vs Operating (ex-AOCI)\n{'='*78}")

    by_gaap = sorted([r for r in results if r.gaap_roe is not None],
                     key=lambda r: -r.gaap_roe)
    by_op = sorted([r for r in results if r.operating_roe is not None],
                   key=lambda r: -r.operating_roe)
    gaap_rank = {r.ticker: i + 1 for i, r in enumerate(by_gaap)}
    op_rank   = {r.ticker: i + 1 for i, r in enumerate(by_op)}

    header = (f"{'Ticker':<7}{'GAAP ROE':>10}{'Rank':>6}"
              f"{'Op ROE':>10}{'Rank':>6}{'Δ Rank':>8}{'Δ ROE (pp)':>13}")
    print(header)
    print("-" * len(header))
    for r in sorted(results, key=lambda x: -x.gaap_roe if x.gaap_roe else 0):
        if r.gaap_roe is None:
            continue
        gr, orr = gaap_rank[r.ticker], op_rank[r.ticker]
        delta_rank = gr - orr  # +ve = stripping AOCI moves rank UP
        delta_pp = (r.operating_roe - r.gaap_roe) * 100
        focus_marker = " ←" if r.ticker == FOCUS else ""
        print(f"{r.ticker:<7}{r.gaap_roe*100:>9.2f}%{gr:>6}"
              f"{r.operating_roe*100:>9.2f}%{orr:>6}"
              f"{delta_rank:>+8d}{delta_pp:>+12.2f}{focus_marker}")

    print("\nΔ Rank: positive = stripping AOCI moves this name UP the rankings")
    print("Δ ROE (pp): negative = AOCI was inflating GAAP ROE (negative AOCI shrinks denominator)")


def main():
    print(f"SEC EDGAR User-Agent: {USER_AGENT}")
    print(f"Computing ROE for {len(TICKERS)} tickers...\n")
    results = []
    for t in TICKERS:
        try:
            r = compute_roe(t)
            yfr = yf_roe(t)
            print_detail(r, yfr)
            results.append(r)
        except Exception as e:
            print(f"\n!! {t} failed: {type(e).__name__}: {e}")
    print_rankings(results)


if __name__ == "__main__":
    main()
