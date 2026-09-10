#!/usr/bin/env python3
"""
Catalaxia / KO - SEC financial extractor
Primary layer: ECONOMIC FACTS

The extractor uses SEC CompanyFacts as the only source of financial facts.

CRITICAL FIX:
For KO the fiscal quarter identity is anchored by the known SEC period-end
date. We do NOT trust a comparative fact's fy/fp alone and we NEVER recover
period_end by matching a numeric value.

Q1/Q2/Q3:
    select facts by EXACT KO period_end + duration window.
Q1/Q2/Q3 standalone:
    Q1 = Q1 YTD
    Q2 = Q2 YTD - Q1 YTD
    Q3 = Q3 YTD - Q2 YTD
Q4:
    FY - Q3 standalone
FY:
    annual 10-K fact

FY2020 is exported to Calculation_Buffer for 5Y calculations.
"""

from __future__ import annotations

import json
import os
import re
from io import StringIO
from pathlib import Path
from typing import Optional

import pandas as pd
import requests


CIK = "0000021344"
TICKER = "KO"

START_YEAR = 2020
END_YEAR = 2025

SEC_URL = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"

RAW_DIR = Path("data/raw/companyfacts")
RAW_FILE = RAW_DIR / f"CIK{CIK}.json"

OUT_CSV = Path("KO_FINANCIAL_INPUTS_2021_2025.csv")
OUT_XLSX = Path("KO_FINANCIAL_INPUTS_2021_2025.xlsx")

USER_AGENT = os.getenv(
    "SEC_USER_AGENT",
    "Catalaxia Research contact@example.com",
)


# ---------------------------------------------------------------------------
# KO fiscal calendar used as PERIOD IDENTITY CONTRACT.
# These are period ends, not fabricated financial values.
# ---------------------------------------------------------------------------

KO_PERIOD_ENDS = {
    2020: {
        "Q1": "2020-03-27",
        "Q2": "2020-06-26",
        "Q3": "2020-09-25",
        "FY": "2020-12-31",
    },
    2021: {
        "Q1": "2021-04-02",
        "Q2": "2021-07-02",
        "Q3": "2021-10-01",
        "FY": "2021-12-31",
    },
    2022: {
        "Q1": "2022-04-01",
        "Q2": "2022-07-01",
        "Q3": "2022-09-30",
        "FY": "2022-12-31",
    },
    2023: {
        "Q1": "2023-03-31",
        "Q2": "2023-06-30",
        "Q3": "2023-09-29",
        "FY": "2023-12-31",
    },
    2024: {
        "Q1": "2024-03-29",
        "Q2": "2024-06-28",
        "Q3": "2024-09-27",
        "FY": "2024-12-31",
    },
    2025: {
        "Q1": "2025-03-28",
        "Q2": "2025-06-27",
        "Q3": "2025-09-26",
        "FY": "2025-12-31",
    },
}


FLOW_CONCEPTS = {
    "revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
    ],
    "gross_profit": [
        "GrossProfit",
    ],
    "operating_income": [
        "OperatingIncomeLoss",
    ],
    "net_income": [
        "NetIncomeLoss",
        "ProfitLoss",
    ],
    "da": [
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
        "Depreciation",
    ],
    "sga": [
        "SellingGeneralAndAdministrativeExpense",
        "SellingGeneralAndAdministrativeExpenses",
        "SellingAndMarketingExpense",
    ],
    "interest_expense": [
        "InterestExpenseNonOperating",
        "InterestExpenseDebt",
        "InterestExpenseNonOperatingAndOther",
        "InterestExpense",
    ],
    "interest_income": [
        "InterestIncomeExpenseNonOperatingNet",
        "InterestIncomeNonOperating",
        "InterestIncomeExpenseNonOperating",
        "InterestIncome",
    ],
    "income_before_tax": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxes",
        "IncomeBeforeTaxExpenseBenefit",
    ],
    "income_tax_expense": [
        "IncomeTaxExpenseBenefit",
        "IncomeTaxExpenseBenefitContinuingOperations",
    ],
    "cfo": [
        "NetCashProvidedByUsedInOperatingActivities",
    ],
    "capex": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
    ],
    "eps_diluted": [
        "EarningsPerShareDiluted",
    ],
    "diluted_shares": [
        "WeightedAverageNumberOfDilutedSharesOutstanding",
    ],
}


STOCK_CONCEPTS = {
    "assets": [
        "Assets",
    ],
    "cash": [
        "CashAndCashEquivalentsAtCarryingValue",
    ],
    "short_term_investments": [
        "ShortTermInvestments",
        "MarketableSecuritiesCurrent",
        "AvailableForSaleSecuritiesCurrent",
        "MarketableSecurities",
    ],
    "accounts_receivable": [
        "AccountsReceivableNetCurrent",
        "AccountsReceivableNet",
        "ReceivablesNetCurrent",
    ],
    "inventory": [
        "InventoryNet",
        "InventoryFinishedGoods",
        "InventoryRawMaterialsAndSupplies",
    ],
    "current_assets": [
        "AssetsCurrent",
    ],
    "ppe_net": [
        "PropertyPlantAndEquipmentNet",
    ],
    "current_liabilities": [
        "LiabilitiesCurrent",
    ],
    "accounts_payable": [
        "AccountsPayableAndAccruedLiabilitiesCurrent",
        "AccountsPayableAndAccruedExpensesCurrent",
        "AccountsPayableAndOtherAccruedLiabilitiesCurrent",
        "AccountsPayableAndOtherAccruedLiabilities",
        "AccountsPayableCurrent",
    ],
    "debt_current": [
        "LongTermDebtCurrent",
        "LongTermDebtAndFinanceLeaseObligationsCurrent",
        "LongTermDebtAndFinanceLeaseObligationsCurrentMaturities",
        "ShortTermBorrowings",
        "ShortTermDebt",
        "LoansAndNotesPayableCurrent",
        "CurrentMaturitiesOfLongTermDebt",
    ],
    "debt_long_term": [
        "LongTermDebtNoncurrent",
        "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
        "LongTermDebtAndFinanceLeaseObligations",
        "LongTermDebt",
    ],
    "liabilities": [
        "Liabilities",
    ],
    "equity": [
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ],
}


# ---------------------------------------------------------------------------
# SEC
# ---------------------------------------------------------------------------

def load_companyfacts() -> dict:
    RAW_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    if RAW_FILE.exists():
        return json.loads(
            RAW_FILE.read_text(
                encoding="utf-8"
            )
        )

    headers = {
        "User-Agent": USER_AGENT,
        "Accept-Encoding": "gzip, deflate",
        "Host": "data.sec.gov",
    }

    response = requests.get(
        SEC_URL,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    data = response.json()

    RAW_FILE.write_text(
        json.dumps(
            data,
            indent=2,
        ),
        encoding="utf-8",
    )

    return data


def concept_rows(
    data: dict,
    concept: str,
) -> list[dict]:
    node = (
        data
        .get("facts", {})
        .get("us-gaap", {})
        .get(concept)
    )

    if not node:
        return []

    result = []

    for unit, rows in node.get(
        "units",
        {},
    ).items():
        for row in rows:
            item = dict(row)
            item["_concept"] = concept
            item["_unit"] = unit
            result.append(item)

    return result


def value(row: Optional[dict]):
    if row is None:
        return None

    try:
        return float(row["val"])
    except (
        KeyError,
        TypeError,
        ValueError,
    ):
        return None


def duration_days(row: dict) -> Optional[int]:
    start = row.get("start")
    end = row.get("end")

    if not start or not end:
        return None

    return (
        pd.Timestamp(end)
        - pd.Timestamp(start)
    ).days + 1


def deduplicate(
    rows: list[dict],
) -> list[dict]:
    """
    Deduplicate only identical economic facts.
    Different start/end periods remain separate.
    """
    groups = {}

    for row in rows:
        key = (
            row.get("_concept"),
            row.get("_unit"),
            row.get("start"),
            row.get("end"),
            row.get("form"),
            row.get("accn"),
        )

        if key not in groups:
            groups[key] = row

    return list(groups.values())


# ---------------------------------------------------------------------------
# PERIOD SELECTORS
# ---------------------------------------------------------------------------

def select_ytd_exact_end(
    rows: list[dict],
    year: int,
    period: str,
) -> Optional[dict]:
    """
    Exact-end selector.

    This is the important correction:
    period identity comes from the exact KO period-end contract.
    """
    expected_end = (
        KO_PERIOD_ENDS
        .get(year, {})
        .get(period)
    )

    if expected_end is None:
        return None

    ranges = {
        "Q1": (60, 130),
        "Q2": (130, 230),
        "Q3": (220, 330),
    }

    lo, hi = ranges[period]

    candidates = []

    for row in deduplicate(rows):
        if row.get("end") != expected_end:
            continue

        if row.get("form") not in {
            "10-Q",
            "10-Q/A",
        }:
            continue

        d = duration_days(row)

        if d is None:
            continue

        if not (
            lo <= d <= hi
        ):
            continue

        # Prefer facts filed in target FY, then original 10-Q.
        fy_rank = (
            1
            if row.get("fy") == year
            else 0
        )

        form_rank = (
            1
            if row.get("form") == "10-Q"
            else 0
        )

        candidates.append(
            (
                fy_rank,
                form_rank,
                str(row.get("filed", "")),
                row,
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (
            x[0],
            x[1],
            x[2],
        )
    )

    return candidates[-1][3]


def select_fy_exact_end(
    rows: list[dict],
    year: int,
) -> Optional[dict]:
    expected_end = (
        KO_PERIOD_ENDS
        .get(year, {})
        .get("FY")
    )

    if expected_end is None:
        return None

    candidates = []

    for row in deduplicate(rows):
        if row.get("end") != expected_end:
            continue

        if row.get("form") not in {
            "10-K",
            "10-K/A",
        }:
            continue

        d = duration_days(row)

        if d is None or d < 300:
            continue

        fy_rank = (
            1
            if row.get("fy") == year
            else 0
        )

        form_rank = (
            1
            if row.get("form") == "10-K"
            else 0
        )

        candidates.append(
            (
                fy_rank,
                form_rank,
                str(row.get("filed", "")),
                row,
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (
            x[0],
            x[1],
            x[2],
        )
    )

    return candidates[-1][3]


def select_stock_exact_end(
    rows: list[dict],
    period_end: str,
) -> Optional[dict]:
    candidates = []

    for row in deduplicate(rows):
        if row.get("end") != period_end:
            continue

        if row.get("form") not in {
            "10-Q",
            "10-Q/A",
            "10-K",
            "10-K/A",
        }:
            continue

        form_rank = (
            1
            if row.get("form")
            in {"10-Q", "10-K"}
            else 0
        )

        candidates.append(
            (
                form_rank,
                str(row.get("filed", "")),
                row,
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (
            x[0],
            x[1],
        )
    )

    return candidates[-1][2]


# ---------------------------------------------------------------------------
# FACT HELPERS
# ---------------------------------------------------------------------------

def select_flow_fact(
    data: dict,
    concepts: list[str],
    year: int,
    period: str,
) -> Optional[dict]:
    for concept in concepts:
        rows = concept_rows(
            data,
            concept,
        )

        if period == "FY":
            row = select_fy_exact_end(
                rows,
                year,
            )
        else:
            row = select_ytd_exact_end(
                rows,
                year,
                period,
            )

        if row is not None:
            return row

    return None


def select_stock_fact(
    data: dict,
    concepts: list[str],
    period_end: str,
) -> Optional[dict]:
    for concept in concepts:
        row = select_stock_exact_end(
            concept_rows(
                data,
                concept,
            ),
            period_end,
        )

        if row is not None:
            return row

    return None


def fact_meta(
    row: Optional[dict],
) -> dict:
    if row is None:
        return {}

    return {
        "concept": row.get("_concept"),
        "unit": row.get("_unit"),
        "start": row.get("start"),
        "end": row.get("end"),
        "fy": row.get("fy"),
        "fp": row.get("fp"),
        "form": row.get("form"),
        "filed": row.get("filed"),
        "accn": row.get("accn"),
        "frame": row.get("frame"),
    }


def usd_billions(
    x: Optional[float],
):
    if x is None:
        return None

    return x / 1_000_000_000


# ---------------------------------------------------------------------------
# SEC FILING FALLBACK FOR DEBT
# ---------------------------------------------------------------------------

# CompanyFacts is already the authoritative source used by this extractor.
# For the filing fallback we therefore reuse the accession number present in
# the selected XBRL facts instead of depending on data.sec.gov/submissions.
_filing_balance_cache: dict[str, dict] = {}
_filing_doc_cache: dict[str, str | None] = {}
SEC_DATA: dict = {}


def sec_headers() -> dict:
    return {
        "User-Agent": USER_AGENT,
        "Accept-Encoding": "gzip, deflate",
        "Host": "www.sec.gov",
    }


def filing_document_from_accession(accn: str) -> str | None:
    """Resolve the primary filing HTML from an SEC accession directory.

    This avoids the submissions endpoint, which can intermittently return 404
    even when CompanyFacts and the filing archive are available.
    """
    if not accn:
        return None
    if accn in _filing_doc_cache:
        return _filing_doc_cache[accn]

    accn_nodash = accn.replace("-", "")
    index_url = (
        f"https://www.sec.gov/Archives/edgar/data/{int(CIK)}/"
        f"{accn_nodash}/index.json"
    )
    try:
        r = requests.get(index_url, headers=sec_headers(), timeout=60)
        r.raise_for_status()
        data = r.json()
    except requests.RequestException:
        _filing_doc_cache[accn] = None
        return None

    items = data.get("directory", {}).get("item", [])
    names = [str(x.get("name", "")) for x in items if isinstance(x, dict)]
    htm = [n for n in names if n.lower().endswith((".htm", ".html"))]
    if not htm:
        _filing_doc_cache[accn] = None
        return None

    # Prefer KO's filing document naming convention; otherwise choose the
    # largest HTML file reported by the archive index.
    preferred = [n for n in htm if n.lower().startswith(("ko-", "a202", "form10"))]
    candidates = preferred or htm
    if len(candidates) > 1:
        sizes = {
            str(x.get("name", "")): int(x.get("size", 0) or 0)
            for x in items if isinstance(x, dict)
        }
        candidates.sort(key=lambda n: sizes.get(n, 0), reverse=True)
    doc = candidates[0]
    _filing_doc_cache[accn] = doc
    return doc


def filing_for_period(period_end: str) -> Optional[dict]:
    """Find an SEC accession for the target period from CompanyFacts facts."""
    candidates = []
    # Use facts already loaded by the extractor. Prefer a stock fact because
    # its exact end date identifies the filing balance-sheet period.
    for concept in (
        "Assets", "AssetsCurrent", "Liabilities", "StockholdersEquity",
        "LongTermDebtNoncurrent", "LongTermDebtCurrent",
        "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
    ):
        for row in deduplicate(concept_rows(SEC_DATA, concept)):
            if row.get("end") != period_end:
                continue
            if row.get("form") not in {"10-Q", "10-Q/A", "10-K", "10-K/A"}:
                continue
            accn = row.get("accn")
            if not accn:
                continue
            candidates.append({
                "form": row.get("form"),
                "accn": accn,
                "filed": row.get("filed", ""),
            })
        if candidates:
            break

    if not candidates:
        return None

    candidates.sort(key=lambda x: (
        not str(x["form"]).endswith("/A"),
        str(x["filed"]),
    ))
    chosen = candidates[-1]
    doc = filing_document_from_accession(chosen["accn"])
    if doc is None:
        return None
    chosen["doc"] = doc
    return chosen

def _clean_number(x) -> Optional[float]:
    if pd.isna(x):
        return None
    s = str(x).replace("$", "").replace(",", "").replace(" ", "").strip()
    if not s or s in {"-", "—", "nan", "NaN"}:
        return None
    # Parentheses denote negatives in SEC tables.
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else None


def _row_value_current(row) -> Optional[float]:
    # The first numeric value after the row label is the current-period
    # amount in KO's consolidated balance-sheet tables.
    vals = []
    for cell in row.tolist():
        v = _clean_number(cell)
        if v is not None:
            vals.append(v)
    return vals[0] if vals else None


def fetch_debt_from_filing(period_end: str) -> dict:
    """
    Fallback for KO debt periods where CompanyFacts does not expose the
    required standard fact cleanly. The filing itself is authoritative.

    Returns USD billions and provenance. No values are hard-coded.
    """
    if period_end in _filing_balance_cache:
        return _filing_balance_cache[period_end]

    filing = filing_for_period(period_end)
    if filing is None:
        result = {"debt_current": None, "debt_long_term": None,
                  "debt_current_meta": {}, "debt_long_meta": {},
                  "null_reason": "FILING_FALLBACK_UNAVAILABLE"}
        _filing_balance_cache[period_end] = result
        return result

    accn_nodash = filing["accn"].replace("-", "")
    url = (
        f"https://www.sec.gov/Archives/edgar/data/{int(CIK)}/"
        f"{accn_nodash}/{filing['doc']}"
    )

    try:
        r = requests.get(url, headers=sec_headers(), timeout=60)
        r.raise_for_status()
        tables = pd.read_html(StringIO(r.text))
    except (requests.RequestException, ValueError):
        result = {"debt_current": None, "debt_long_term": None,
                  "debt_current_meta": {}, "debt_long_meta": {},
                  "null_reason": "FILING_FALLBACK_UNAVAILABLE"}
        _filing_balance_cache[period_end] = result
        return result

    loans = None
    maturities = None
    long_term = None

    for table in tables:
        text = " ".join(table.astype(str).fillna("").to_numpy().ravel()).lower()
        if "current liabilities" not in text or "long-term debt" not in text:
            continue

        for _, row in table.iterrows():
            row_text = " ".join(str(x) for x in row.tolist()).lower()
            if "loans and notes payable" in row_text:
                loans = _row_value_current(row)
            elif "current maturities of long-term debt" in row_text:
                maturities = _row_value_current(row)
            elif re.search(r"\blong-term debt\b", row_text) and "current maturities" not in row_text:
                long_term = _row_value_current(row)

        if loans is not None and maturities is not None and long_term is not None:
            break

    current = None if loans is None and maturities is None else (loans or 0) + (maturities or 0)
    result = {
        "debt_current": current / 1000 if current is not None else None,
        "debt_long_term": long_term / 1000 if long_term is not None else None,
        "debt_current_meta": {
            "concept": "SEC_FILING_TABLE:LoansAndNotesPayable+CurrentMaturitiesOfLongTermDebt",
            "unit": "USD millions",
            "end": period_end,
            "form": filing["form"],
            "filed": filing["filed"],
            "accn": filing["accn"],
            "source_url": url,
        },
        "debt_long_meta": {
            "concept": "SEC_FILING_TABLE:LongTermDebt",
            "unit": "USD millions",
            "end": period_end,
            "form": filing["form"],
            "filed": filing["filed"],
            "accn": filing["accn"],
            "source_url": url,
        },
    }
    _filing_balance_cache[period_end] = result
    return result


# ---------------------------------------------------------------------------
# BUILD
# ---------------------------------------------------------------------------

def add_source_columns(
    record: dict,
    name: str,
    row: Optional[dict],
):
    meta = fact_meta(row)

    record[
        f"{name}_source_concept"
    ] = meta.get("concept")

    record[
        f"{name}_source_unit"
    ] = meta.get("unit")

    record[
        f"{name}_source_start"
    ] = meta.get("start")

    record[
        f"{name}_source_end"
    ] = meta.get("end")

    record[
        f"{name}_source_fy"
    ] = meta.get("fy")

    record[
        f"{name}_source_fp"
    ] = meta.get("fp")

    record[
        f"{name}_source_form"
    ] = meta.get("form")

    record[
        f"{name}_source_filed"
    ] = meta.get("filed")

    record[
        f"{name}_source_accn"
    ] = meta.get("accn")

    record[
        f"{name}_source_frame"
    ] = meta.get("frame")


def build_record(
    data: dict,
    year: int,
    period: str,
    period_end: str,
    flow_facts: dict,
    flow_values: dict,
) -> dict:
    record = {
        "ticker": TICKER,
        "cik": CIK,
        "fiscal_year": year,
        "period": period,
        "period_end": period_end,
        "source": "SEC_EDGAR_COMPANYFACTS",
    }

    for name, v in flow_values.items():
        record[name] = v
        add_source_columns(
            record,
            name,
            flow_facts.get(name),
        )

    for name, concepts in STOCK_CONCEPTS.items():
        row = select_stock_fact(
            data,
            concepts,
            period_end,
        )

        v = value(row)

        record[name] = usd_billions(v)

        add_source_columns(
            record,
            name,
            row,
        )

    debt_current = record.get("debt_current")
    debt_long = record.get("debt_long_term")

    # CompanyFacts is primary. If KO's post-2024 balance-sheet debt facts
    # are not exposed cleanly, use the actual SEC filing table as a controlled
    # fallback. This is still SEC source data, not a hard-coded value.
    if debt_current is None or debt_long is None:
        fb = fetch_debt_from_filing(period_end)
        if debt_current is None and fb.get("debt_current") is not None:
            record["debt_current"] = fb["debt_current"]
            for k, v in fb["debt_current_meta"].items():
                if k == "concept":
                    record["debt_current_source_concept"] = v
                elif k == "unit":
                    record["debt_current_source_unit"] = v
                elif k == "end":
                    record["debt_current_source_end"] = v
                elif k == "form":
                    record["debt_current_source_form"] = v
                elif k == "filed":
                    record["debt_current_source_filed"] = v
                elif k == "accn":
                    record["debt_current_source_accn"] = v
        if debt_long is None and fb.get("debt_long_term") is not None:
            record["debt_long_term"] = fb["debt_long_term"]
            for k, v in fb["debt_long_meta"].items():
                if k == "concept":
                    record["debt_long_term_source_concept"] = v
                elif k == "unit":
                    record["debt_long_term_source_unit"] = v
                elif k == "end":
                    record["debt_long_term_source_end"] = v
                elif k == "form":
                    record["debt_long_term_source_form"] = v
                elif k == "filed":
                    record["debt_long_term_source_filed"] = v
                elif k == "accn":
                    record["debt_long_term_source_accn"] = v

    debt_current = record.get("debt_current")
    debt_long = record.get("debt_long_term")
    record["total_debt"] = (
        (debt_current or 0) + (debt_long or 0)
        if debt_current is not None or debt_long is not None
        else None
    )

    # If the direct Liabilities fact is absent, derive it from
    # Assets - Equity. This remains an ECONOMIC FACT derivation and
    # is explicitly provenance-labelled below.
    if record.get("liabilities") is None:
        assets_v = record.get("assets")
        equity_v = record.get("equity")
        if assets_v is not None and equity_v is not None:
            record["liabilities"] = assets_v - equity_v
            record["liabilities_source_concept"] = "DERIVED:Assets-Equity"
            record["liabilities_source_unit"] = "USD"
            record["liabilities_source_start"] = None
            record["liabilities_source_end"] = period_end
            record["liabilities_source_form"] = "DERIVED"
            record["liabilities_source_accn"] = None
            record["liabilities_source_frame"] = None

    op = record.get(
        "operating_income"
    )

    da = record.get("da")

    if op is not None and da is not None:
        record["ebitda"] = op + da
    else:
        record["ebitda"] = None

    cfo = record.get("cfo")
    capex = record.get("capex")

    if cfo is not None and capex is not None:
        record["fcf"] = cfo - capex
    else:
        record["fcf"] = None

    return record


def extract_year(
    data: dict,
    year: int,
) -> list[dict]:
    q1_facts = {}
    q2_facts = {}
    q3_facts = {}
    fy_facts = {}

    for name, concepts in FLOW_CONCEPTS.items():
        q1_facts[name] = select_flow_fact(
            data,
            concepts,
            year,
            "Q1",
        )

        q2_facts[name] = select_flow_fact(
            data,
            concepts,
            year,
            "Q2",
        )

        q3_facts[name] = select_flow_fact(
            data,
            concepts,
            year,
            "Q3",
        )

        fy_facts[name] = select_flow_fact(
            data,
            concepts,
            year,
            "FY",
        )

    anchors = {
        "Q1": q1_facts["revenue"],
        "Q2": q2_facts["revenue"],
        "Q3": q3_facts["revenue"],
        "FY": fy_facts["revenue"],
    }

    missing = [
        p
        for p, fact in anchors.items()
        if fact is None
    ]

    if missing:
        raise RuntimeError(
            f"{year}: missing SEC revenue anchor(s): "
            f"{missing}"
        )

    # YTD values in billions except EPS/shares.
    def fact_values(facts):
        out = {}

        for name, row in facts.items():
            v = value(row)

            if name in {
                "eps_diluted",
                "diluted_shares",
            }:
                out[name] = v
            else:
                out[name] = usd_billions(v)

        return out

    q1_ytd = fact_values(q1_facts)
    q2_ytd = fact_values(q2_facts)
    q3_ytd = fact_values(q3_facts)
    fy_values = fact_values(fy_facts)

    records = []

    # ---------------------------------------------------------------
    # Q1
    # ---------------------------------------------------------------
    q1_values = dict(q1_ytd)

    # EPS is not additive. Calculate from quarterly NI / YTD diluted shares.
    if (
        q1_values.get("net_income") is not None
        and q1_values.get("diluted_shares")
    ):
        q1_values["eps_diluted"] = (
            q1_values["net_income"]
            * 1_000_000_000
            / q1_values["diluted_shares"]
        )

    q1 = build_record(
        data,
        year,
        "Q1",
        anchors["Q1"]["end"],
        q1_facts,
        q1_values,
    )

    records.append(q1)

    # ---------------------------------------------------------------
    # Q2 standalone = Q2 YTD - Q1 YTD
    # ---------------------------------------------------------------
    q2_values = {}

    for name in FLOW_CONCEPTS:
        if name in {
            "eps_diluted",
            "diluted_shares",
        }:
            q2_values[name] = None
            continue

        current = q2_ytd.get(name)
        previous = q1_ytd.get(name)

        q2_values[name] = (
            current - previous
            if current is not None
            and previous is not None
            else None
        )

    if (
        q2_values.get("net_income") is not None
        and q2_ytd.get("diluted_shares")
    ):
        q2_values["eps_diluted"] = (
            q2_values["net_income"]
            * 1_000_000_000
            / q2_ytd["diluted_shares"]
        )

    q2 = build_record(
        data,
        year,
        "Q2",
        anchors["Q2"]["end"],
        q2_facts,
        q2_values,
    )

    records.append(q2)

    # ---------------------------------------------------------------
    # Q3 standalone = Q3 YTD - Q2 YTD
    # ---------------------------------------------------------------
    q3_values = {}

    for name in FLOW_CONCEPTS:
        if name in {
            "eps_diluted",
            "diluted_shares",
        }:
            q3_values[name] = None
            continue

        current = q3_ytd.get(name)
        previous = q2_ytd.get(name)

        q3_values[name] = (
            current - previous
            if current is not None
            and previous is not None
            else None
        )

    if (
        q3_values.get("net_income") is not None
        and q3_ytd.get("diluted_shares")
    ):
        q3_values["eps_diluted"] = (
            q3_values["net_income"]
            * 1_000_000_000
            / q3_ytd["diluted_shares"]
        )

    q3 = build_record(
        data,
        year,
        "Q3",
        anchors["Q3"]["end"],
        q3_facts,
        q3_values,
    )

    records.append(q3)

    # ---------------------------------------------------------------
    # Q4 = FY - Q3 standalone
    # ---------------------------------------------------------------
    q4_values = {}

    for name in FLOW_CONCEPTS:
        if name in {
            "eps_diluted",
            "diluted_shares",
        }:
            q4_values[name] = None
            continue

        fy = fy_values.get(name)
        q3 = q3_values.get(name)

        q4_values[name] = (
            fy - q3
            if fy is not None
            and q3 is not None
            else None
        )

    if (
        q4_values.get("net_income") is not None
        and fy_values.get("diluted_shares")
    ):
        q4_values["eps_diluted"] = (
            q4_values["net_income"]
            * 1_000_000_000
            / fy_values["diluted_shares"]
        )

    q4 = build_record(
        data,
        year,
        "Q4",
        anchors["FY"]["end"],
        fy_facts,
        q4_values,
    )

    records.append(q4)

    # ---------------------------------------------------------------
    # FY
    # ---------------------------------------------------------------
    fy = build_record(
        data,
        year,
        "FY",
        anchors["FY"]["end"],
        fy_facts,
        fy_values,
    )

    records.append(fy)

    return records


# ---------------------------------------------------------------------------
# VALIDATION
# ---------------------------------------------------------------------------

def validate(
    records: list[dict],
):
    expected = {
        (year, period)
        for year in range(2021, 2026)
        for period in (
            "Q1",
            "Q2",
            "Q3",
            "Q4",
            "FY",
        )
    }

    actual = {
        (
            int(r["fiscal_year"]),
            r["period"],
        )
        for r in records
    }

    if actual != expected:
        raise AssertionError(
            "Period set incorrect. "
            f"Missing/extra: {expected ^ actual}"
        )

    for year in range(
        2021,
        2026,
    ):
        for period in (
            "Q1",
            "Q2",
            "Q3",
        ):
            row = next(
                r
                for r in records
                if int(r["fiscal_year"]) == year
                and r["period"] == period
            )

            expected_end = (
                KO_PERIOD_ENDS[year][period]
            )

            if row["period_end"] != expected_end:
                raise AssertionError(
                    f"{year} {period} wrong period_end: "
                    f"{row['period_end']} != {expected_end}"
                )

            source_end = row.get(
                "revenue_source_end"
            )

            if source_end != expected_end:
                raise AssertionError(
                    f"{year} {period}: revenue source end "
                    f"{source_end} != {expected_end}"
                )

    for row in records:
        if row["period"] in {
            "Q1",
            "Q2",
            "Q3",
            "Q4",
            "FY",
        }:
            if row.get("period_end") is None:
                raise AssertionError(
                    f"Missing period_end: {row}"
                )

    if len(records) != 25:
        raise AssertionError(
            f"Expected 25 rows, got {len(records)}"
        )


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    data = load_companyfacts()
    global SEC_DATA
    SEC_DATA = data

    buffer_2020 = extract_year(
        data,
        2020,
    )

    records = []

    for year in range(
        2021,
        END_YEAR + 1,
    ):
        records.extend(
            extract_year(
                data,
                year,
            )
        )

    # FY2020 baseline is provenance-preserving metadata used only by the
    # downstream 5Y growth calculation. It is not a new economic period in
    # the 25-row output.
    fy2020 = next(
        r for r in buffer_2020
        if r["period"] == "FY"
    )
    for r in records:
        r["eps_fy2020_baseline"] = fy2020.get("eps_diluted")
        r["eps_fy2020_baseline_source"] = "Calculation_Buffer/FY2020"
        r["eps_fy2020_baseline_period_end"] = fy2020.get("period_end")

    # FY2020 balance baselines support average-balance ratios for FY2021.
    for r in records:
        r["assets_fy2020_baseline"] = fy2020.get("assets")
        r["equity_fy2020_baseline"] = fy2020.get("equity")
        r["ppe_net_fy2020_baseline"] = fy2020.get("ppe_net")
        r["accounts_receivable_fy2020_baseline"] = fy2020.get("accounts_receivable")
        r["inventory_fy2020_baseline"] = fy2020.get("inventory")
        r["accounts_payable_fy2020_baseline"] = fy2020.get("accounts_payable")

    validate(records)

    debt_missing = [
        (r["fiscal_year"], r["period"], r["period_end"])
        for r in records
        if r.get("debt_current") is None or r.get("debt_long_term") is None
    ]
    if debt_missing:
        raise AssertionError(
            "Debt extraction incomplete after SEC CompanyFacts + filing fallback: "
            + str(debt_missing)
        )

    order = {
        "Q1": 1,
        "Q2": 2,
        "Q3": 3,
        "Q4": 4,
        "FY": 5,
    }

    df = pd.DataFrame(records)

    df["_order"] = df[
        "period"
    ].map(order)

    df = (
        df
        .sort_values(
            [
                "fiscal_year",
                "_order",
            ]
        )
        .drop(
            columns="_order"
        )
    )

    # Explicitly overwrite the output generated by previous runs.
    df.to_csv(
        OUT_CSV,
        index=False,
        encoding="utf-8-sig",
    )

    with pd.ExcelWriter(
        OUT_XLSX,
        engine="openpyxl",
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="KO_Financials",
            index=False,
        )

        pd.DataFrame(
            buffer_2020
        ).to_excel(
            writer,
            sheet_name="Calculation_Buffer",
            index=False,
        )

    print()
    print("==============================================")
    print("KO SEC EXTRACTION")
    print("==============================================")
    print(f"Rows: {len(df)}")
    print(f"CSV:  {OUT_CSV.resolve()}")
    print(f"XLSX: {OUT_XLSX.resolve()}")
    print()
    print("PERIODS:")

    for row in df.itertuples():
        print(
            f"{int(row.fiscal_year)} "
            f"{row.period:<2} "
            f"{row.period_end} "
            f"Revenue={row.revenue}"
        )

    print()
    print("PERIOD VALIDATION: PASS")
    print("PROVENANCE VALIDATION: PASS")
    print("==============================================")


if __name__ == "__main__":
    main()
