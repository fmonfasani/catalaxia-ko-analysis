#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent
EDGAR = BASE / "KO_FINANCIAL_INPUTS_2021_2025.csv"
YAHOO = BASE / "KO_YAHOO_MARKET_2021_2025.csv"
OUT_CSV = BASE / "KO_RATIOS_2021_2025.csv"
OUT_XLSX = BASE / "KO_RATIOS_2021_2025.xlsx"

RATIOS = [
    "ROA_%", "ROIC_%", "ROE_%", "Common_Equity_ROE_%",
    "Gross_Margin_%", "SGA_%", "EBITDA_Margin_%", "EBITA_Margin_%",
    "EBIT_Margin_%", "Continuing_Operations_Income_Margin_%", "Net_Margin_%",
    "Normalized_Net_Margin_%", "FCF_Levered_Margin_%", "FCF_Unlevered_Margin_%",
    "CFO_Current_Liabilities_%", "FCFonCE_%", "Asset_Turnover", "Fixed_Asset_Turnover",
    "AR_Turnover", "Inventory_Turnover", "Working_Capital_Turnover", "CCC", "DSO", "DIO", "DPO",
    "Current_Ratio", "Quick_Ratio", "Debt_Equity_%", "Debt_Total_Capital_%", "Liabilities_Assets_%",
    "Debt_EBITDA", "Net_Debt_EBITDA", "Net_Debt_EBITDA_Capex", "EBIT_Interest", "EBITDA_Interest",
    "EBITDA_Capex_Interest", "FFO_Interest", "FFO_Debt", "P_E", "Annual_EPS", "EPS_Growth_5Y_%",
    "P_E_Growth_5Y_%", "Price", "52W_High", "Diff_vs_52W_High_%", "52W_Low", "Diff_vs_52W_Low_%",
    "Dividend_Yield_%", "Payout_Ratio_%",
]

FLOW_COLS = [
    "revenue", "gross_profit", "operating_income", "net_income", "ebitda", "cfo", "capex",
    "cost_of_revenue", "sga", "interest_expense", "income_before_tax", "income_tax_expense",
]


def num(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(np.nan, index=df.index, dtype="float64")
    return pd.to_numeric(df[col], errors="coerce")


def div(a, b):
    a = pd.to_numeric(a, errors="coerce")
    b = pd.to_numeric(b, errors="coerce")
    return a.div(b).where(b.notna() & b.ne(0))


def pct(a, b):
    return div(a, b) * 100.0


def require_inputs(e, y):
    if len(e) != 25 or len(y) != 25:
        raise AssertionError(f"Se requieren 25 filas: EDGAR={len(e)}, YAHOO={len(y)}")
    keys = ["fiscal_year", "period", "period_end"]
    for k in keys:
        if k not in e.columns or k not in y.columns:
            raise AssertionError(f"Falta clave {k}")
    e["period_end"] = pd.to_datetime(e["period_end"], errors="raise").dt.strftime("%Y-%m-%d")
    y["period_end"] = pd.to_datetime(y["period_end"], errors="raise").dt.strftime("%Y-%m-%d")
    if e.duplicated(keys).any() or y.duplicated(keys).any():
        raise AssertionError("Hay claves duplicadas en las tablas de entrada")


def build_quarter_ttm(df: pd.DataFrame, col: str) -> pd.Series:
    """
    TTM strictly by fiscal quarter identity, never by physical row adjacency.
    FY rows are excluded from the quarter window. For a target quarter, use
    that quarter plus the previous three fiscal quarters.
    """
    out = pd.Series(np.nan, index=df.index, dtype="float64")
    q = df[df["period"].isin(["Q1", "Q2", "Q3", "Q4"])].copy()
    q["fy"] = pd.to_numeric(q["fiscal_year"], errors="coerce").astype(int)
    order = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
    q["qnum"] = q["period"].map(order).astype(int)
    q["seq"] = q["fy"] * 4 + q["qnum"]
    q = q.sort_values("seq")

    vals = pd.to_numeric(q[col], errors="coerce") if col in q.columns else pd.Series(np.nan, index=q.index)
    q = q.assign(_v=vals)
    by_seq = q.set_index("seq")

    for idx, row in q.iterrows():
        s = int(row["seq"])
        needed = [s - 3, s - 2, s - 1, s]
        if not all(x in by_seq.index for x in needed):
            continue
        window = by_seq.loc[needed, "_v"]
        if window.notna().all():
            out.loc[idx] = float(window.sum())
    return out


def current_or_ttm(df, col):
    out = pd.Series(np.nan, index=df.index, dtype="float64")
    ttm = build_quarter_ttm(df, col)
    fy = df["period"].eq("FY")
    raw = num(df, col)
    out.loc[df["period"].isin(["Q1", "Q2", "Q3", "Q4"])] = ttm.loc[df["period"].isin(["Q1", "Q2", "Q3", "Q4"])]
    out.loc[fy] = raw.loc[fy]
    return out


def previous_balance(df, col):
    """Previous balance by fiscal identity, with FY2020 baseline for FY2021."""
    result = pd.Series(np.nan, index=df.index, dtype="float64")
    cur = num(df, col)
    baseline_col = f"{col}_fy2020_baseline"
    baseline = num(df, baseline_col) if baseline_col in df.columns else pd.Series(np.nan, index=df.index)

    # Build lookup independent of dataframe row order.
    lookup = {(int(r["fiscal_year"]), r["period"]): cur.loc[i] for i, r in df.iterrows()}
    for i, r in df.iterrows():
        fy = int(r["fiscal_year"])
        p = r["period"]
        if p == "Q1":
            key = (fy - 1, "FY")
        elif p == "Q2":
            key = (fy, "Q1")
        elif p == "Q3":
            key = (fy, "Q2")
        elif p == "Q4":
            key = (fy, "Q3")
        else:
            key = (fy - 1, "FY")
        if key in lookup and pd.notna(lookup[key]):
            result.loc[i] = lookup[key]
        elif fy == 2021 and pd.notna(baseline.loc[i]):
            result.loc[i] = baseline.loc[i]
    return result


def average_balance(df, col):
    cur = num(df, col)
    prev = previous_balance(df, col)
    return (cur + prev) / 2.0

def main():
    if not EDGAR.exists():
        raise FileNotFoundError(f"Falta {EDGAR}")
    if not YAHOO.exists():
        raise FileNotFoundError(f"Falta {YAHOO}")

    e = pd.read_csv(EDGAR)
    y = pd.read_csv(YAHOO)
    require_inputs(e, y)

    keys = ["fiscal_year", "period", "period_end"]
    df = e.merge(y, on=keys, how="left", validate="one_to_one", suffixes=("_edgar", "_yahoo"))
    if len(df) != 25:
        raise AssertionError(f"Merge produjo {len(df)} filas, esperado 25")
    df = df.sort_values(["period_end", "period"], key=lambda s: s).reset_index(drop=True)

    # Inputs.
    revenue = num(df, "revenue")
    gross_profit = num(df, "gross_profit")
    operating_income = num(df, "operating_income")
    net_income = num(df, "net_income")
    ebitda = num(df, "ebitda")
    cfo = num(df, "cfo")
    capex = num(df, "capex")
    assets = num(df, "assets")
    equity = num(df, "equity")
    liabilities = num(df, "liabilities")
    current_assets = num(df, "current_assets")
    current_liabilities = num(df, "current_liabilities")
    cash = num(df, "cash")
    sti = num(df, "short_term_investments")
    ar = num(df, "accounts_receivable")
    inventory = num(df, "inventory")
    ppe = num(df, "ppe_net")
    ap = num(df, "accounts_payable")
    debt = num(df, "total_debt")
    eps = num(df, "eps_diluted")
    cogs = num(df, "cost_of_revenue")
    # KO reports gross profit; when an explicit COGS fact is absent,
    # COGS is the deterministic identity Revenue - Gross Profit.
    cogs = cogs.where(cogs.notna(), revenue - gross_profit)
    sga = num(df, "sga")
    interest = num(df, "interest_expense")
    pretax = num(df, "income_before_tax")
    tax = num(df, "income_tax_expense")
    ffo = num(df, "ffo")

    out = df[keys].copy()

    if "cost_of_revenue" not in df.columns or pd.to_numeric(df["cost_of_revenue"], errors="coerce").isna().all():
        df["cost_of_revenue"] = revenue - gross_profit
    else:
        df["cost_of_revenue"] = pd.to_numeric(df["cost_of_revenue"], errors="coerce").where(
            pd.to_numeric(df["cost_of_revenue"], errors="coerce").notna(), revenue - gross_profit
        )

    # TTM flow series. FY uses the FY fact itself; quarters use 4 standalone quarters.
    rev_ttm = current_or_ttm(df, "revenue")
    gp_ttm = current_or_ttm(df, "gross_profit")
    op_ttm = current_or_ttm(df, "operating_income")
    ni_ttm = current_or_ttm(df, "net_income")
    ebitda_ttm = current_or_ttm(df, "ebitda")
    cfo_ttm = current_or_ttm(df, "cfo")
    capex_ttm = current_or_ttm(df, "capex")
    cogs_ttm = current_or_ttm(df, "cost_of_revenue")
    sga_ttm = current_or_ttm(df, "sga")
    interest_ttm = current_or_ttm(df, "interest_expense")
    pretax_ttm = current_or_ttm(df, "income_before_tax")
    tax_ttm = current_or_ttm(df, "income_tax_expense")

    # Profitability/margins.
    out["Gross_Margin_%"] = pct(gp_ttm, rev_ttm)
    out["SGA_%"] = pct(sga_ttm, rev_ttm)
    out["EBITDA_Margin_%"] = pct(ebitda_ttm, rev_ttm)
    out["EBIT_Margin_%"] = pct(op_ttm, rev_ttm)
    out["Net_Margin_%"] = pct(ni_ttm, rev_ttm)
    out["FCF_Levered_Margin_%"] = pct(cfo_ttm - capex_ttm, rev_ttm)

    # Balance-sheet averages use fiscal sequence, not physical row adjacency.
    avg_assets = average_balance(df, "assets")
    avg_equity = average_balance(df, "equity")
    avg_ppe = average_balance(df, "ppe_net")
    avg_ar = average_balance(df, "accounts_receivable")
    avg_inventory = average_balance(df, "inventory")
    avg_ap = average_balance(df, "accounts_payable")

    out["ROA_%"] = pct(ni_ttm, avg_assets)
    out["ROE_%"] = pct(ni_ttm, avg_equity)
    out["Common_Equity_ROE_%"] = pct(ni_ttm, avg_equity)
    out["Asset_Turnover"] = div(rev_ttm, avg_assets)
    out["Fixed_Asset_Turnover"] = div(rev_ttm, avg_ppe)
    out["AR_Turnover"] = div(rev_ttm, avg_ar)
    out["Inventory_Turnover"] = div(cogs_ttm, avg_inventory)
    out["Working_Capital_Turnover"] = div(rev_ttm, current_assets - current_liabilities)

    out["DSO"] = div(avg_ar, rev_ttm) * 365
    out["DIO"] = div(avg_inventory, cogs_ttm) * 365
    out["DPO"] = div(avg_ap, cogs_ttm) * 365
    out["CCC"] = out["DSO"] + out["DIO"] - out["DPO"]

    # Liquidity.
    out["Current_Ratio"] = div(current_assets, current_liabilities)
    out["Quick_Ratio"] = div(current_assets - inventory, current_liabilities)
    out["CFO_Current_Liabilities_%"] = pct(cfo_ttm, current_liabilities)
    out["FCFonCE_%"] = pct(cfo_ttm - capex_ttm, equity)

    # Leverage.
    out["Debt_Equity_%"] = pct(debt, equity)
    out["Debt_Total_Capital_%"] = pct(debt, debt + equity)
    out["Liabilities_Assets_%"] = pct(liabilities, assets)
    net_debt = debt - cash - sti
    out["Debt_EBITDA"] = div(debt, ebitda_ttm)
    out["Net_Debt_EBITDA"] = div(net_debt, ebitda_ttm)
    out["Net_Debt_EBITDA_Capex"] = div(net_debt, ebitda_ttm - capex_ttm)

    # Coverage.
    out["EBIT_Interest"] = div(op_ttm, interest_ttm)
    out["EBITDA_Interest"] = div(ebitda_ttm, interest_ttm)
    out["EBITDA_Capex_Interest"] = div(ebitda_ttm - capex_ttm, interest_ttm)
    out["FFO_Interest"] = div(ffo, interest_ttm)
    out["FFO_Debt"] = div(ffo, debt)

    # Contract B: ROIC = NOPAT / Average Invested Capital.
    # NOPAT = EBIT * (1 - tax rate). Invested Capital = Equity + Debt - Cash - ST investments.
    tax_rate = div(tax_ttm, pretax_ttm)
    nopat = op_ttm * (1.0 - tax_rate)
    avg_equity_ic = average_balance(df, "equity")
    avg_debt_ic = average_balance(df, "total_debt")
    avg_cash_ic = average_balance(df, "cash")
    avg_sti_ic = average_balance(df, "short_term_investments")
    avg_invested_capital = avg_equity_ic + avg_debt_ic - avg_cash_ic - avg_sti_ic
    out["ROIC_%"] = pct(nopat, avg_invested_capital)

    # Contract B ratios for which the required source inputs/methodology are not yet available.
    out["EBITA_Margin_%"] = np.nan
    out["Continuing_Operations_Income_Margin_%"] = np.nan
    out["Normalized_Net_Margin_%"] = np.nan
    out["FCF_Unlevered_Margin_%"] = np.nan

    # Market fields.
    out["Price"] = num(df, "price_usd")
    out["52W_High"] = num(df, "52w_high_usd")
    out["52W_Low"] = num(df, "52w_low_usd")
    out["Diff_vs_52W_High_%"] = pct(out["Price"], out["52W_High"]) - 100
    out["Diff_vs_52W_Low_%"] = pct(out["Price"], out["52W_Low"]) - 100
    out["Dividend_Yield_%"] = num(df, "dividend_yield_ttm_pct")

    fy = out["period"].eq("FY")
    out["Annual_EPS"] = np.where(fy, eps, np.nan)
    out["P_E"] = div(out["Price"], out["Annual_EPS"])

    # 5Y EPS growth: FY2020 baseline + current FY EPS.
    baseline = num(df, "eps_fy2020_baseline")
    out["EPS_Growth_5Y_%"] = np.nan
    valid_eps5 = fy & baseline.notna() & out["Annual_EPS"].notna() & baseline.gt(0) & pd.Series(out["Annual_EPS"], index=out.index).gt(0)
    out.loc[valid_eps5, "EPS_Growth_5Y_%"] = (
        (out.loc[valid_eps5, "Annual_EPS"] / baseline.loc[valid_eps5]) ** (1 / 5) - 1
    ) * 100

    # Contract B: P/E Growth 5Y = CAGR of P/E from FY2020 to current FY.
    # P/E FY = fiscal-year-end trading close / annual diluted EPS.
    fy2020_price = num(df, "fy2020_price_usd_baseline")
    pe2020 = div(fy2020_price, baseline)
    out["P_E_Growth_5Y_%"] = np.nan
    valid_pe5 = fy & pe2020.notna() & out["P_E"].notna() & pe2020.gt(0) & out["P_E"].gt(0)
    out.loc[valid_pe5, "P_E_Growth_5Y_%"] = (
        (out.loc[valid_pe5, "P_E"] / pe2020.loc[valid_pe5]) ** (1 / 5) - 1
    ) * 100

    divps = num(df, "dividend_per_share_ttm_usd")
    out["Payout_Ratio_%"] = np.where(fy, pct(divps, out["Annual_EPS"]), np.nan)

    for r in RATIOS:
        if r not in out.columns:
            out[r] = np.nan

    # Explicit NULL reasons. A reason is assigned at the row level whenever
    # at least one contractual field is NULL; this is not a substitute for
    # field-level provenance but prevents silent missing values.
    def reasons(row):
        rr = []
        p = row["period"]
        if p != "FY":
            rr.append("ANNUAL_ONLY_FIELDS_NOT_APPLICABLE")
        if pd.isna(row["ROIC_%"]):
            rr.append("ROIC_REQUIRED_INPUT_MISSING")
        if pd.isna(row["EBITA_Margin_%"]):
            rr.append("CONTRACT_B_EBITA")
        if pd.isna(row["Continuing_Operations_Income_Margin_%"]):
            rr.append("CONTRACT_B_CONTINUING_OPERATIONS")
        if pd.isna(row["Normalized_Net_Margin_%"]):
            rr.append("CONTRACT_B_NORMALIZED_NET_MARGIN")
        if pd.isna(row["FCF_Unlevered_Margin_%"]):
            rr.append("CONTRACT_B_FCF_UNLEVERED")
        if pd.isna(row["FFO_Interest"]) or pd.isna(row["FFO_Debt"]):
            rr.append("FFO_INPUT_NOT_AVAILABLE")
        if pd.isna(row["P_E_Growth_5Y_%"]):
            rr.append("PE_5Y_REQUIRES_FY2020_PRICE_AND_EPS")
        return ";".join(rr)

    out["null_reason"] = out.apply(reasons, axis=1)
    out = out[keys + RATIOS + ["null_reason"]]

    # Structural verification.
    if len(out) != 25:
        raise AssertionError(f"Salida: {len(out)} filas; esperado 25")
    missing_cols = [r for r in RATIOS if r not in out.columns]
    if missing_cols:
        raise AssertionError(f"Faltan ratios contractuales: {missing_cols}")

    out.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    out.to_excel(OUT_XLSX, index=False)

    print("=" * 60)
    print("KO RATIOS — BLOQUE B")
    print("=" * 60)
    print("Rows:", len(out))
    print("Ratios:", len(RATIOS))
    print("NaN por columna:")
    for c in RATIOS:
        n = int(out[c].isna().sum())
        if n:
            print(f"  {c}: {n}/25")
    print("OUTPUT:", OUT_CSV)
    print("STRUCTURAL VERIFICATION: PASS")
    print("=" * 60)


if __name__ == "__main__":
    main()
