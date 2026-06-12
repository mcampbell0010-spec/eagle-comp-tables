"""Reusable EDGAR XBRL data layer.

Pulls SEC company-facts JSON and exposes helpers for computing fundamentals:
  - TTM via FY + latest Q − prior-year same Q (avoids the standalone-Q4 trap)
  - 5-point average of any balance-sheet concept (this Q + previous 4 Q-ends)
  - High-level `compute_roe(ticker)` returning GAAP ROE and Operating ROE

USER_AGENT must contain a contact email per SEC's compliance rules.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import requests


def _load_dotenv(path: Path) -> None:
    """Populate os.environ from KEY=VALUE lines in a .env file. Existing
    environment variables take precedence over the file."""
    if not path.exists():
        return
    for raw in path.read_text().splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


_load_dotenv(Path(__file__).parent / ".env")

USER_AGENT = os.environ.get("SEC_USER_AGENT")
if not USER_AGENT:
    raise RuntimeError(
        "SEC_USER_AGENT is not set. SEC EDGAR requires identifiable contact "
        "info on every request. Copy .env.example to .env and set your name "
        "and email, e.g.:\n\n"
        "    SEC_USER_AGENT=Jane Smith jane@example.com\n"
    )

HEADERS = {"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"}

# Module-level cache so we don't re-pull facts within one run.
_facts_cache: dict[str, dict] = {}
_ticker_map_cache: Optional[dict] = None

# Fallback chains for concept names that vary by filer.
NET_INCOME_CONCEPTS = [
    "NetIncomeLoss",
    "NetIncomeLossAvailableToCommonStockholdersBasic",
    "ProfitLoss",
]
EQUITY_CONCEPTS = [
    "StockholdersEquity",
    "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
]
AOCI_CONCEPTS = [
    "AccumulatedOtherComprehensiveIncomeLossNetOfTax",
    "AccumulatedOtherComprehensiveIncomeLossAttributableToParentNetOfTax",
]
GOODWILL_CONCEPTS = ["Goodwill"]
INTANGIBLES_CONCEPTS = [
    "IntangibleAssetsNetExcludingGoodwill",
    "FiniteLivedIntangibleAssetsNet",
    "OtherIntangibleAssetsNet",
]


# --- HTTP / lookup -----------------------------------------------------------

def _ticker_map() -> dict:
    global _ticker_map_cache
    if _ticker_map_cache is None:
        r = requests.get("https://www.sec.gov/files/company_tickers.json",
                         headers=HEADERS, timeout=30)
        r.raise_for_status()
        _ticker_map_cache = r.json()
    return _ticker_map_cache


def get_cik(ticker: str) -> str:
    for entry in _ticker_map().values():
        if entry["ticker"].upper() == ticker.upper():
            return f"{int(entry['cik_str']):010d}"
    raise ValueError(f"No CIK found for {ticker}")


def company_facts(ticker: str) -> dict:
    if ticker not in _facts_cache:
        cik = get_cik(ticker)
        url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        _facts_cache[ticker] = r.json()
    return _facts_cache[ticker]


def entries(facts: dict, concept: str,
            taxonomy: str = "us-gaap", unit: str = "USD") -> list[dict]:
    try:
        return facts["facts"][taxonomy][concept]["units"][unit]
    except KeyError:
        return []


def first_available(facts: dict, candidates: list[str],
                    taxonomy: str = "us-gaap", unit: str = "USD"
                    ) -> tuple[str, list[dict]]:
    """Try each concept in order, return (concept_used, entries) for the first
    that yields data. Returns (None, []) if none work."""
    for c in candidates:
        es = entries(facts, c, taxonomy=taxonomy, unit=unit)
        if es:
            return c, es
    return None, []


# --- Period helpers ----------------------------------------------------------

def quarterly(es: list[dict]) -> list[dict]:
    out = []
    for e in es:
        if "start" not in e or "end" not in e:
            continue
        days = (date.fromisoformat(e["end"]) - date.fromisoformat(e["start"])).days
        if 80 <= days <= 100:
            out.append(e)
    return out


def annuals(es: list[dict]) -> list[dict]:
    out = []
    for e in es:
        if "start" not in e or "end" not in e:
            continue
        days = (date.fromisoformat(e["end"]) - date.fromisoformat(e["start"])).days
        if 350 <= days <= 380:
            out.append(e)
    return out


def instants(es: list[dict]) -> list[dict]:
    return [e for e in es if "end" in e and ("start" not in e or e["start"] == e["end"])]


def latest_per_end(es: list[dict]) -> list[dict]:
    by_end: dict[str, dict] = {}
    for e in es:
        end = e["end"]
        prev = by_end.get(end)
        if prev is None or e.get("filed", "") > prev.get("filed", ""):
            by_end[end] = e
    return sorted(by_end.values(), key=lambda x: x["end"])


def _last_day_of_month(year: int, month: int) -> int:
    nxt = date(year + (1 if month == 12 else 0), 1 if month == 12 else month + 1, 1)
    return (nxt - timedelta(days=1)).day


def _months_before(iso: str, months: int) -> str:
    d = date.fromisoformat(iso)
    year, month = d.year, d.month - months
    while month <= 0:
        month += 12
        year -= 1
    day = min(d.day, _last_day_of_month(year, month))
    return date(year, month, day).isoformat()


def _one_year_before(iso: str) -> str:
    return _months_before(iso, 12)


def instant_near(es: list[dict], target_iso: str,
                 window_days: int = 35) -> Optional[dict]:
    target = date.fromisoformat(target_iso)
    best, best_diff = None, 10**6
    for e in latest_per_end(instants(es)):
        diff = abs((date.fromisoformat(e["end"]) - target).days)
        if diff < best_diff and diff <= window_days:
            best, best_diff = e, diff
    return best


# --- TTM and 5-point average -------------------------------------------------

def ttm(es: list[dict]) -> tuple[Optional[float], dict]:
    """TTM = most recent FY + latest quarter past FY-end − prior-year same Q.

    Avoids the standalone-Q4-not-tagged trap (Q4 only appears inside the 10-K's
    full-year line). Returns (value, diagnostics).
    """
    yrs = latest_per_end(annuals(es))
    if not yrs:
        return None, {"method": "no annual entries"}
    fy = yrs[-1]
    fy_end = date.fromisoformat(fy["end"])

    qs = latest_per_end(quarterly(es))
    newer = [q for q in qs if date.fromisoformat(q["end"]) > fy_end]
    if not newer:
        # No quarter past FY-end → TTM is just the latest FY.
        return fy["val"], {"method": "FY only", "period_end": fy["end"],
                           "fy_end": fy["end"], "fy_val": fy["val"]}

    lq = newer[-1]
    target = _one_year_before(lq["end"])
    pq = None
    for q in qs:
        if abs((date.fromisoformat(q["end"]) - date.fromisoformat(target)).days) <= 7:
            pq = q
            break
    if pq is None:
        return None, {"method": "no prior-year matching Q"}

    val = fy["val"] + lq["val"] - pq["val"]
    return val, {
        "method": "FY + latest Q − prior-yr Q",
        "period_end": lq["end"],
        "fy_end": fy["end"], "fy_val": fy["val"],
        "lq_end": lq["end"], "lq_val": lq["val"],
        "pq_end": pq["end"], "pq_val": pq["val"],
    }


def five_point_average(es: list[dict], period_end: str,
                       window_days: int = 35
                       ) -> tuple[Optional[float], list[dict]]:
    """Average of a balance-sheet concept at 5 quarter-end snapshots ending
    at `period_end`. Returns (avg, list_of_snapshots_used). Falls back to
    whatever points it can find within `window_days` of each target."""
    targets = [period_end] + [_months_before(period_end, k) for k in (3, 6, 9, 12)]
    snaps = []
    for t in targets:
        s = instant_near(es, t, window_days=window_days)
        if s is not None:
            snaps.append(s)
    if not snaps:
        return None, []
    return sum(s["val"] for s in snaps) / len(snaps), snaps


# --- High-level: ROE for one ticker ------------------------------------------

@dataclass
class RoeResult:
    ticker: str
    company: str
    period_end: str
    ni_ttm: Optional[float]
    ni_method: dict
    ni_concept: Optional[str]
    eq_concept: Optional[str]
    aoci_concept: Optional[str]
    avg_equity: Optional[float]
    avg_equity_n_points: int
    avg_aoci: Optional[float]
    avg_aoci_n_points: int
    avg_equity_ex_aoci: Optional[float]
    gaap_roe: Optional[float]
    operating_roe: Optional[float]
    # Balance-sheet items at period_end (for P/TBV)
    book_value: Optional[float] = None
    goodwill: Optional[float] = None
    intangibles: Optional[float] = None
    tangible_book: Optional[float] = None
    goodwill_concept: Optional[str] = None
    intangibles_concept: Optional[str] = None
    notes: list[str] = field(default_factory=list)


def compute_roe(ticker: str) -> RoeResult:
    facts = company_facts(ticker)
    name = facts.get("entityName", ticker)
    notes: list[str] = []

    # Net income TTM (try concept fallbacks in order).
    ni_concept, ni_es = first_available(facts, NET_INCOME_CONCEPTS)
    if not ni_es:
        notes.append("no net-income concept found")
    if ni_concept and ni_concept != "NetIncomeLoss":
        notes.append(f"net income from us-gaap:{ni_concept}")
    ni_ttm_val, ni_method = (None, {"method": "no data"}) if not ni_es else ttm(ni_es)
    period_end = ni_method.get("period_end")

    # Equity (concept fallbacks).
    eq_concept, eq_es = first_available(facts, EQUITY_CONCEPTS)
    if eq_concept and eq_concept != "StockholdersEquity":
        notes.append(f"equity from us-gaap:{eq_concept}")

    avg_eq, eq_snaps = (None, [])
    if period_end and eq_es:
        avg_eq, eq_snaps = five_point_average(eq_es, period_end)
        if len(eq_snaps) < 5:
            notes.append(f"only {len(eq_snaps)}/5 equity snapshots within window")

    # AOCI (concept fallbacks).
    aoci_concept, aoci_es = first_available(facts, AOCI_CONCEPTS)
    avg_aoci, aoci_snaps = (None, [])
    if aoci_es and period_end:
        avg_aoci, aoci_snaps = five_point_average(aoci_es, period_end)
        if aoci_concept != "AccumulatedOtherComprehensiveIncomeLossNetOfTax":
            notes.append(f"AOCI from us-gaap:{aoci_concept}")
        if len(aoci_snaps) < 5:
            notes.append(f"only {len(aoci_snaps)}/5 AOCI snapshots within window")
    elif not aoci_es:
        notes.append("no AOCI concept found — operating ROE = GAAP ROE")
        avg_aoci = 0  # treat as zero AOCI so op ROE falls back to GAAP

    avg_eq_ex = (avg_eq - avg_aoci) if (avg_eq is not None and avg_aoci is not None) else None
    gaap = (ni_ttm_val / avg_eq) if (ni_ttm_val is not None and avg_eq) else None
    op = (ni_ttm_val / avg_eq_ex) if (ni_ttm_val is not None and avg_eq_ex) else None

    # Latest book value (= equity at period_end) — used for P/TBV.
    book_value = None
    if period_end and eq_es:
        snap = instant_near(eq_es, period_end, window_days=10)
        if snap is not None:
            book_value = snap["val"]

    # Goodwill at period_end (treat absent tag as $0).
    goodwill_concept, gw_es = first_available(facts, GOODWILL_CONCEPTS)
    goodwill = 0.0
    if gw_es and period_end:
        snap = instant_near(gw_es, period_end, window_days=10)
        if snap is not None:
            goodwill = snap["val"]
    elif not gw_es:
        notes.append("no Goodwill tag found — treating as $0")

    # Intangibles at period_end (concept fallbacks; absent tag = $0).
    intangibles_concept, int_es = first_available(facts, INTANGIBLES_CONCEPTS)
    intangibles = 0.0
    if int_es and period_end:
        snap = instant_near(int_es, period_end, window_days=10)
        if snap is not None:
            intangibles = snap["val"]
        if intangibles_concept != "IntangibleAssetsNetExcludingGoodwill":
            notes.append(f"intangibles from us-gaap:{intangibles_concept}")
    elif not int_es:
        notes.append("no intangibles tag found — treating as $0")

    tangible_book = (book_value - goodwill - intangibles) if book_value is not None else None

    return RoeResult(
        ticker=ticker, company=name, period_end=period_end or "?",
        ni_ttm=ni_ttm_val, ni_method=ni_method,
        ni_concept=ni_concept, eq_concept=eq_concept, aoci_concept=aoci_concept,
        avg_equity=avg_eq, avg_equity_n_points=len(eq_snaps),
        avg_aoci=avg_aoci, avg_aoci_n_points=len(aoci_snaps),
        avg_equity_ex_aoci=avg_eq_ex,
        gaap_roe=gaap, operating_roe=op,
        book_value=book_value, goodwill=goodwill, intangibles=intangibles,
        tangible_book=tangible_book,
        goodwill_concept=goodwill_concept, intangibles_concept=intangibles_concept,
        notes=notes,
    )
