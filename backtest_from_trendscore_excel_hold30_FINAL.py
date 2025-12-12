import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime

# ============================================================
# KONFIGURATION
# ============================================================

INPUT_EXCEL = "trendscore2025_weekly_top30.xlsx"
TOP_N = 15
HOLD_MAX_RANK = 30

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_XLSX = f"backtest_results_top15_hold30_{timestamp}.xlsx"

# ============================================================
# EXCEL LADEN (robust)
# ============================================================

df = pd.read_excel(INPUT_EXCEL, sheet_name=0)

# Spalten normalisieren
df.columns = [c.strip() for c in df.columns]

# Score / score vereinheitlichen
if "Score" not in df.columns and "score" in df.columns:
    df.rename(columns={"score": "Score"}, inplace=True)

required = {"WeekEnd", "Rank", "Ticker", "Score"}
missing = required - set(df.columns)
if missing:
    raise ValueError(f"Fehlende Spalten im Excel: {missing}")

df["WeekEnd"] = pd.to_datetime(df["WeekEnd"])
df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
df["Rank"] = df["Rank"].astype(int)
df["Score"] = df["Score"].astype(float)

df = df.sort_values(["WeekEnd", "Rank"]).reset_index(drop=True)
weeks = sorted(df["WeekEnd"].unique())

# Wochen-Mappings
week_maps = {}
for w in weeks:
    wdf = df[df["WeekEnd"] == w]
    week_maps[w] = {
        "rank": dict(zip(wdf["Ticker"], wdf["Rank"])),
        "score": dict(zip(wdf["Ticker"], wdf["Score"])),
        "ordered": wdf["Ticker"].tolist()
    }

# ============================================================
# KURSDATEN LADEN
# ============================================================

tickers = sorted(df["Ticker"].unique())
start_date = weeks[0].strftime("%Y-%m-%d")

prices = yf.download(
    tickers,
    start=start_date,
    auto_adjust=True,
    progress=False
)["Close"]

if isinstance(prices, pd.Series):
    prices = prices.to_frame()

prices = prices.dropna(how="all")

def price_on_or_after(ticker, date):
    if ticker not in prices.columns:
        return None, None
    s = prices[ticker].dropna()
    if s.empty:
        return None, None
    i = s.index.searchsorted(date)
    if i >= len(s):
        return None, None
    return s.index[i], float(s.iloc[i])

# ============================================================
# BACKTEST
# ============================================================

equity = 1.0
holdings = {}  # ticker -> shares, entry_price
events = []
equity_curve = []

def log_event(date, ticker, action, price, score, entry_price=None):
    trade_return = None
    if action == "EXIT" and entry_price:
        trade_return = (price / entry_price - 1) * 100

    events.append({
        "Date": date,
        "Ticker": ticker,
        "Action": action,
        "Price": price,
        "Score": score,
        "TradeReturn_%": trade_return
    })

def mark_to_market(date):
    total = 0.0
    for t, pos in holdings.items():
        _, px = price_on_or_after(t, date)
        if px is not None:
            total += pos["shares"] * px
    return total

# ----------------------------
# INITIAL: Top-15 kaufen
# ----------------------------
w0 = weeks[0]
alloc = equity / TOP_N

for t in week_maps[w0]["ordered"][:TOP_N]:
    d, px = price_on_or_after(t, w0)
    if px is None:
        continue
    holdings[t] = {"shares": alloc / px, "entry_price": px}
    log_event(d, t, "ENTRY", px, week_maps[w0]["score"].get(t))

equity = mark_to_market(w0)
equity_curve.append({"Date": w0, "Equity": equity, "Holdings": len(holdings)})

# ----------------------------
# WÖCHENTLICHER LOOP
# ----------------------------
for w in weeks[1:]:
    wm = week_maps[w]

    # EXITS: Rank > 30
    for t in list(holdings.keys()):
        if wm["rank"].get(t, 9999) > HOLD_MAX_RANK:
            d, px = price_on_or_after(t, w)
            if px is not None:
                log_event(
                    d, t, "EXIT", px,
                    wm["score"].get(t),
                    holdings[t]["entry_price"]
                )
            holdings.pop(t)

    # NACHKAUF bis wieder 15
    if len(holdings) < TOP_N:
        equity = mark_to_market(w)
        alloc = equity / TOP_N
        for t in wm["ordered"]:
            if t not in holdings:
                d, px = price_on_or_after(t, w)
                if px is None:
                    continue
                holdings[t] = {"shares": alloc / px, "entry_price": px}
                log_event(d, t, "ENTRY", px, wm["score"].get(t))
                if len(holdings) == TOP_N:
                    break

    equity = mark_to_market(w)
    equity_curve.append({"Date": w, "Equity": equity, "Holdings": len(holdings)})

# ----------------------------
# FINALE GLATTSTELLUNG
# ----------------------------
final_date = prices.index.max()

for t in list(holdings.keys()):
    d, px = price_on_or_after(t, final_date)
    if px is not None:
        log_event(
            d, t, "EXIT", px, np.nan,
            holdings[t]["entry_price"]
        )

equity_curve.append({
    "Date": final_date,
    "Equity": mark_to_market(final_date),
    "Holdings": 0
})

# ============================================================
# EXPORT
# ============================================================

events_df = pd.DataFrame(events).sort_values("Date").reset_index(drop=True)
equity_df = pd.DataFrame(equity_curve)

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    events_df.to_excel(writer, sheet_name="Trades", index=False)
    equity_df.to_excel(writer, sheet_name="EquityCurve", index=False)

print("✅ Backtest abgeschlossen")
print(f"Datei: {OUTPUT_XLSX}")
print(f"Trades: {len(events_df)}")
print(f"Finale Equity: {equity_df.iloc[-1]['Equity']:.4f}")
