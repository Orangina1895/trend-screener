import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime

# ============================================================
# KONFIGURATION
# ============================================================

START_DATE = "2022-12-23"
END_DATE   = "2023-24-07"

OUTPUT_FILE = "trendscore2025_weekly_top30.xlsx"
TOP_N_EXPORT = 30

# ============================================================
# TICKER-UNIVERSUM (AKTUALISIERT)
# ============================================================

TICKERS = sorted(list(set([
    "AAPL","ABNB","ADBE","ADI","ADP","ADSK","AEP","AMAT","AMD","AMGN","AMZN",
    "APP","ARM","ASML","AVGO","AZN","BIIB","BKNG","BKR","SWKS","CDNS","VRSN",
    "CEG","CHTR","CMCSA","COST","CPRT","CRWD","CSCO","CSGP","CSX","CTAS",
    "CTSH","NTES","DDOG","DXCM","EA","EBAY","EXC","FAST","FTNT","GEHC","GFS",
    "GILD","GOOG","GOOGL","HON","IDXX","INTC","INTU","ISRG","KDP","KHC",
    "KLAC","LIN","LRCX","LULU","MAR","MCHP","BIDU","MDLZ","MELI","META","MNST",
    "MRVL","MSFT","MU","NFLX","NVDA","NXPI","ODFL","ON","ORLY","PANW",
    "PAYX","PCAR","PDD","PEP","PYPL","QCOM","REGN","MTCH","ROST","SBUX",
    "SHOP","SIRI","SNPS","TEAM","TMUS","TRI","TSLA","TTD","TTWO","TXN",
    "VRSK","VRSN","VRTX","WBD","WDAY","XEL","ZS",

    # ➕ NEU
    "ILMN","SMCI","MRNA"
])))

# ❌ ENTFERNT (bewusst nicht enthalten):
# PLTR, MSTR, AXON

print(f"Ticker im Scan: {len(TICKERS)}")

# ============================================================
# DATEN LADEN
# ============================================================

data = yf.download(
    TICKERS,
    start="2023-01-01",  # früher für saubere Indikatoren
    end=END_DATE,
    auto_adjust=True,
    progress=False,
    group_by="ticker"
)

def extract(field):
    return pd.DataFrame({t: data[t][field] for t in TICKERS if t in data})

closes = extract("Close")
highs  = extract("High")
lows   = extract("Low")

# ============================================================
# HILFSFUNKTIONEN
# ============================================================

def compute_atr(high, low, close, window=20):
    prev_close = close.shift(1)
    tr = pd.concat([
        high - low,
        (high - prev_close).abs(),
        (low - prev_close).abs()
    ], axis=0).groupby(level=0).max()
    return tr.rolling(window).mean()

# ============================================================
# WEEKLY TREND SCORE
# ============================================================

def compute_weekly_trend_score(w_closes, w_highs, w_lows, asof):
    hist = w_closes.loc[:asof]
    if len(hist) < 52:
        return None

    last = hist.iloc[-1]

    valid = hist.notna().tail(52).count() == 52
    tickers = valid[valid].index.tolist()
    if not tickers:
        return None

    hist = hist[tickers]
    last = last[tickers]

    m6  = last / hist.iloc[-26] - 1
    m12 = last / hist.iloc[-52] - 1

    sma10 = hist.rolling(10).mean().iloc[-1]
    sma20 = hist.rolling(20).mean().iloc[-1]
    sma40 = hist.rolling(40).mean().iloc[-1]

    atr10 = compute_atr(
        w_highs[tickers], w_lows[tickers], hist, 10
    ).iloc[-1]

    sma_score = (
        (last > sma10).astype(int) +
        (last > sma20).astype(int) +
        (last > sma40).astype(int) +
        (sma10 > sma20).astype(int) +
        (sma20 > sma40).astype(int)
    ) / 5

    vol_adj = m12 / (atr10 / last)

    df = pd.DataFrame({
        "M6": m6,
        "M12": m12,
        "SMA": sma_score,
        "VA": vol_adj
    }).dropna()

    df["rM6"]  = df["M6"].rank(pct=True)
    df["rM12"] = df["M12"].rank(pct=True)
    df["rVA"]  = df["VA"].rank(pct=True)

    df["score"] = (
        0.33 * df["rM6"] +
        0.33 * df["rM12"] +
        0.20 * df["SMA"] +
        0.14 * df["rVA"]
    )

    return df.sort_values("score", ascending=False)

# ============================================================
# WEEKLY LOOP
# ============================================================

w_closes = closes.resample("W-FRI").last()
w_highs  = highs.resample("W-FRI").max()
w_lows   = lows.resample("W-FRI").min()

rows = []

for d in w_closes.index:
    if d < pd.to_datetime(START_DATE) or d > pd.to_datetime(END_DATE):
        continue

    score_df = compute_weekly_trend_score(w_closes, w_highs, w_lows, d)
    if score_df is None:
        continue

    top = score_df.head(TOP_N_EXPORT)

    for rank, (ticker, row) in enumerate(top.iterrows(), start=1):
        rows.append({
            "WeekEnd": d,
            "Rank": rank,
            "Ticker": ticker,
            "score": row["score"],
            "M6": row["M6"],
            "M12": row["M12"],
            "SMA": row["SMA"],
            "VA": row["VA"]
        })

# ============================================================
# EXPORT
# ============================================================

out = pd.DataFrame(rows)
out.sort_values(["WeekEnd", "Rank"], inplace=True)

out.to_excel(OUTPUT_FILE, index=False)
print(f"Excel erstellt: {OUTPUT_FILE}")
