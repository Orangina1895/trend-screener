# trendscore2025.py
# Wöchentlicher TrendScore-Scanner (Top 30) für vorgegebene Ticker
#
# pip install pandas numpy yfinance openpyxl

import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta

# ============================================================
# KONFIGURATION
# ============================================================

HISTORY_START = "2016-01-01"      # für stabile 12M/6M Scores
SCAN_START    = "2024-12-15"
SCAN_END      = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")

TOP_K = 30
OUTPUT_XLSX = "trendscore2025_weekly_top30.xlsx"

RAW_TICKERS = """
AAPL
ABNB
ADBE
ADI
ADP
ADSK
AEP
AMAT
AMD
AMGN
AMZN
APP
ARM
ASML
AVGO
AXON
AZN
BIIB
BKNG
BKR
CCEP
CDNS
CDW
CEG
CHTR
CMCSA
COST
CPRT
CRWD
CSCO
CSGP
CSX
CTAS
CTSH
DASH
DDOG
DXCM
EA
EBAY
EXC
FAST
FTNT
GEHC
GFS
GILD
GOOG
GOOGL
HON
IDXX
ILMN
INTC
INTU
ISRG
KDP
KHC
KLAC
LIN
LRCX
LULU
MAR
MCHP
MDB
MDLZ
MELI
META
MNST
MRVL
MSFT
MU
NFLX
NVDA
NXPI
ODFL
ON
ORLY
PANW
PAYX
PCAR
PDD
PEP
PLTR
PYPL
QCOM
REGN
ROP
ROST
SBUX
SHOP
SIRI
SNPS
TEAM
TMUS
TRI
TSLA
TTD
TTWO
TXN
VRSK
VRSN
VRTX
WBD
WDAY
XEL
ZS
""".strip()

# ============================================================
# TICKER BEREINIGEN
# ============================================================

TICKERS = sorted(set(t.strip().upper() for t in RAW_TICKERS.splitlines() if t.strip()))
print(f"Ticker (unique): {len(TICKERS)}")

# ============================================================
# DATENLADEN
# ============================================================

def download_ohlc(tickers, start, end):
    data = yf.download(
        tickers,
        start=start,
        end=end,
        auto_adjust=True,
        group_by="ticker",
        progress=False,
        threads=True,
    )

    # Fallback (falls yfinance keinen MultiIndex liefert)
    if not isinstance(data.columns, pd.MultiIndex):
        closes = pd.DataFrame({tickers[0]: data["Close"]})
        highs  = pd.DataFrame({tickers[0]: data["High"]})
        lows   = pd.DataFrame({tickers[0]: data["Low"]})
        return closes, highs, lows

    closes = pd.DataFrame({t: data[t]["Close"] for t in tickers if t in data.columns.get_level_values(0)})
    highs  = pd.DataFrame({t: data[t]["High"]  for t in tickers if t in data.columns.get_level_values(0)})
    lows   = pd.DataFrame({t: data[t]["Low"]   for t in tickers if t in data.columns.get_level_values(0)})
    return closes, highs, lows

# ============================================================
# INDICATORS
# ============================================================

def compute_atr(highs, lows, closes, window=20):
    prev_close = closes.shift(1)
    tr1 = highs - lows
    tr2 = (highs - prev_close).abs()
    tr3 = (lows - prev_close).abs()
    tr = pd.concat([tr1, tr2, tr3], axis=0).groupby(level=0).max()
    return tr.rolling(window).mean()

def compute_trend_score_daily(closes, highs, lows, as_of_date, universe):
    """
    Daily TrendScore:
    - Momentum 6M / 12M
    - SMA-Struktur (50/100/200)
    - Volatility-adjusted Momentum
    """
    d = pd.to_datetime(as_of_date)
    universe = [t for t in universe if t in closes.columns]

    closes_hist = closes.loc[:d, universe]
    highs_hist  = highs.loc[:d, universe]
    lows_hist   = lows.loc[:d, universe]

    if len(closes_hist) < 252:
        return pd.DataFrame(columns=["score"])

    last = closes_hist.iloc[-1]

    enough = closes_hist.notna().tail(252).count() == 252
    valid = enough[enough].index.tolist()
    if not valid:
        return pd.DataFrame(columns=["score"])

    closes_hist = closes_hist[valid]
    highs_hist  = highs_hist[valid]
    lows_hist   = lows_hist[valid]
    last        = last[valid]

    m6  = last / closes_hist.iloc[-126] - 1
    m12 = last / closes_hist.iloc[-252] - 1

    sma50  = closes_hist.rolling(50).mean().iloc[-1]
    sma100 = closes_hist.rolling(100).mean().iloc[-1]
    sma200 = closes_hist.rolling(200).mean().iloc[-1]

    atr20 = compute_atr(highs_hist, lows_hist, closes_hist, 20).iloc[-1]

    sma_points = (
        (last > sma50).astype(int)
        + (last > sma100).astype(int)
        + (last > sma200).astype(int)
        + (sma50 > sma100).astype(int)
        + (sma100 > sma200).astype(int)
    )
    sma_norm = sma_points / 5.0

    vol = (atr20 / last).replace(0, np.nan)
    va_raw = m12 / vol

    df = pd.DataFrame({
        "M6": m6,
        "M12": m12,
        "sma_norm": sma_norm,
        "va_raw": va_raw
    }).dropna()

    if df.empty:
        return pd.DataFrame(columns=["score"])

    df["rank_M6"]  = df["M6"].rank(pct=True)
    df["rank_M12"] = df["M12"].rank(pct=True)
    df["rank_VA"]  = df["va_raw"].rank(pct=True)

    df["score"] = (
        0.33 * df["rank_M6"]
        + 0.33 * df["rank_M12"]
        + 0.20 * df["sma_norm"]
        + 0.14 * df["rank_VA"]
    )

    return df.sort_values("score", ascending=False)

# ============================================================
# MAIN
# ============================================================

def main():
    closes, highs, lows = download_ohlc(TICKERS, HISTORY_START, SCAN_END)

    closes = closes.dropna(how="all", axis=1)
    highs  = highs.reindex(columns=closes.columns)
    lows   = lows.reindex(columns=closes.columns)

    available = closes.columns.tolist()
    missing = sorted(set(TICKERS) - set(available))

    idx = closes.loc[SCAN_START:SCAN_END].index
    fridays = idx[idx.weekday == 4]
    fridays = pd.DatetimeIndex(fridays).unique()

    rows = []
    for d in fridays:
        score_df = compute_trend_score_daily(closes, highs, lows, d, available)
        if score_df.empty:
            continue

        top = score_df.head(TOP_K).copy()
        top.insert(0, "Rank", range(1, len(top) + 1))
        top.insert(0, "WeekEnd", d.date().isoformat())
        top.insert(2, "Ticker", top.index)

        for c in ["score", "M6", "M12", "sma_norm", "va_raw"]:
            if c in top.columns:
                top[c] = top[c].round(6)

        rows.append(top.reset_index(drop=True))

    weekly_top = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        weekly_top.to_excel(writer, sheet_name="Weekly_Top30", index=False)
        pd.DataFrame({"Downloaded": available}).to_excel(writer, sheet_name="Coverage", index=False)
        pd.DataFrame({"Missing": missing}).to_excel(writer, sheet_name="Missing", index=False)

    print(f"Excel erstellt: {OUTPUT_XLSX}")
    print(f"Wochen × Top30-Zeilen: {len(weekly_top)}")
    print(f"Downloaded: {len(available)} | Missing: {len(missing)}")

if __name__ == "__main__":
    main()
