
from __future__ import annotations

import hashlib
import importlib.metadata
import math
from pathlib import Path
from datetime import timedelta

import pandas as pd
import yfinance as yf

BASE = Path(__file__).resolve().parent
INPUT = BASE / "KO_FINANCIAL_INPUTS_2021_2025.csv"
OUTPUT_CSV = BASE / "KO_YAHOO_MARKET_2021_2025.csv"
OUTPUT_XLSX = BASE / "KO_YAHOO_MARKET_2021_2025.xlsx"

TICKER = "KO"

def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def get_price_history():
    tk = yf.Ticker(TICKER)
    # Broad range: enough history to calculate each period's trailing 52-week high/low.
    hist = tk.history(start="2020-01-01", end="2026-01-02",
                      auto_adjust=False, actions=True)
    if hist is None or hist.empty:
        raise RuntimeError("Yahoo no devolvió histórico para KO.")
    hist = hist.reset_index()
    hist["Date"] = pd.to_datetime(hist["Date"], errors="coerce").dt.tz_localize(None)
    return hist

def trailing_dividend(hist, end_date):
    start = end_date - timedelta(days=365)
    x = hist[(hist["Date"] > start) & (hist["Date"] <= end_date)]
    if "Dividends" not in x:
        return math.nan
    v = pd.to_numeric(x["Dividends"], errors="coerce").fillna(0).sum()
    return float(v) if pd.notna(v) else math.nan

def main():
    if not INPUT.exists():
        raise FileNotFoundError(f"No existe {INPUT}")

    fin = pd.read_csv(INPUT)
    if len(fin) != 25:
        raise AssertionError(f"Se esperaban 25 períodos; encontrados {len(fin)}")

    fin["period_end"] = pd.to_datetime(fin["period_end"], errors="raise")
    hist = get_price_history()

    # Support point for contractual P/E Growth 5Y: last trading close on/before FY2020 end.
    fy2020_end = pd.Timestamp("2020-12-31")
    base = hist[hist["Date"] <= fy2020_end]
    if base.empty:
        raise RuntimeError("No hay precio Yahoo para FY2020.")
    base_day = base.iloc[-1]
    fy2020_price = float(base_day["Close"]) if pd.notna(base_day["Close"]) else math.nan
    fy2020_trade_date = pd.Timestamp(base_day["Date"]).strftime("%Y-%m-%d")

    rows = []
    for _, r in fin.iterrows():
        end = r["period_end"]
        # Last trading day on or before the SEC period end.
        eligible = hist[hist["Date"] <= end]
        if eligible.empty:
            raise RuntimeError(f"No hay precio Yahoo para {end.date()}")

        day = eligible.iloc[-1]
        price = float(day["Close"]) if pd.notna(day["Close"]) else math.nan

        window = hist[(hist["Date"] <= end) &
                      (hist["Date"] > end - timedelta(days=365))]
        hi = pd.to_numeric(window["High"], errors="coerce").max()
        lo = pd.to_numeric(window["Low"], errors="coerce").min()

        div = trailing_dividend(hist, end)
        dy = (div / price * 100) if pd.notna(div) and pd.notna(price) and price != 0 else math.nan

        rows.append({
            "fiscal_year": int(r["fiscal_year"]),
            "period": str(r["period"]),
            "period_end": end.strftime("%Y-%m-%d"),
            "price_usd": price,
            "52w_high_usd": float(hi) if pd.notna(hi) else math.nan,
            "52w_low_usd": float(lo) if pd.notna(lo) else math.nan,
            "dividend_per_share_ttm_usd": div,
            "dividend_yield_ttm_pct": dy,
            "trading_date_used": pd.Timestamp(day["Date"]).strftime("%Y-%m-%d"),
            "source": "Yahoo Finance",
            "ticker": TICKER,
            "fy2020_price_usd_baseline": fy2020_price,
            "fy2020_price_trading_date_baseline": fy2020_trade_date,
            "fy2020_price_source_baseline": "Yahoo Finance",
            "yahoo_history_sha256": hashlib.sha256(
                pd.util.hash_pandas_object(hist, index=True).values.tobytes()
            ).hexdigest(),
            "yfinance_version": importlib.metadata.version("yfinance"),
        })

    out = pd.DataFrame(rows)

    checks = {
        c: int(out[c].notna().sum())
        for c in ["price_usd", "52w_high_usd", "52w_low_usd",
                  "dividend_per_share_ttm_usd", "dividend_yield_ttm_pct"]
    }

    if checks["price_usd"] != 25:
        raise AssertionError(f"YAHOO FAIL: price_usd {checks['price_usd']}/25")
    if checks["52w_high_usd"] != 25 or checks["52w_low_usd"] != 25:
        raise AssertionError(f"YAHOO FAIL: 52W high/low = {checks['52w_high_usd']}/25, {checks['52w_low_usd']}/25")
    if checks["dividend_per_share_ttm_usd"] != 25:
        raise AssertionError(f"YAHOO FAIL: dividend TTM {checks['dividend_per_share_ttm_usd']}/25")

    out.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    out.to_excel(OUTPUT_XLSX, index=False)

    print("=" * 46)
    print("KO YAHOO EXTRACTION")
    print("=" * 46)
    print("Rows:", len(out))
    for k, v in checks.items():
        print(f"{k}: {v}/25")
    print("YAHOO EXTRACTION: PASS")
    print("=" * 46)

if __name__ == "__main__":
    main()
