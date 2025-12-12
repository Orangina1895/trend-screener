import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime

# ============================================================
# KONFIG
# ============================================================

INPUT_EXCEL = "trendscore2025_weekly_top30.xlsx"
SHEET_NAME  = "Weekly_Top30"

TOP_N = 15
HOLD_MAX_RANK = 30

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_XLSX = f"backtest_results_hold30_events_{timestamp}.xlsx"

# ============================================================
# EXCEL LADEN
# ============================================================

df = pd.read_excel(INPUT_EXCEL, sheet_name=SHEET_NAME)

required = {"WeekEnd", "Rank", "Ticker", "score"}
if not required.issubset(df.columns):
    raise ValueError(f"Excel muss enthalten: {required}")

df["WeekEnd"] = pd.to_datetime(df["WeekEnd"])
df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
df["Rank"] = pd.to_numeric(df["Rank"], errors="coerce")
df["score"] = pd.to_numeric(df["score"], errors="coerce")

df = df.dropna(subset=["WeekEnd", "Ticker", "Rank"])
df = df.sort_values(["WeekEnd", "Rank"]).reset_index(drop=True)

weeks = sorted(df["WeekEnd"].unique())
if not weeks:
    raise ValueError("Keine Wochen gefunden.")

# Maps je Woche
week_maps = {}
for w in weeks:
    wdf = df[df["WeekEnd"] == w]
    week_maps[w] = {
        "rank": dict(zip(wdf["Ticker"], wdf["Rank"].astype(int))),
        "score": dict(zip(wdf["Ticker"], wdf["score"])),
        "ordered": wdf["Ticker"].tolist()
    }

# ============================================================
# PREISDATEN
# ============================================================

tickers = sorted(df["Ticker"].unique().tolist())
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
if prices.empty:
    raise RuntimeError("Keine Kursdaten geladen.")

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
# BACKTEST (EVENT-LOG)
# ============================================================

equity = 1.0
holdings = {}   # ticker -> shares
events = []     # ENTRY / EXIT Events
equity_curve = []

def mark_to_market(date):
    total = 0.0
    for t, shares in holdings.items():
        _, px = price_on_or_after(t, date)
        if px is not None:
            total += shares * px
    return total

def log_entry(ticker, date, price, score):
    events.append({
        "Date": date,
        "Ticker": ticker,
        "Action": "ENTRY",
        "Price": price,
        "Score": score,
        "EquityAfter": np.nan
    })

def log_exit(ticker, date, price, score):
    events.append({
        "Date": date,
        "Ticker": ticker,
        "Action": "EXIT",
        "Price": price,
        "Score": score,
        "EquityAfter": np.nan
    })

# ----------------------------
# Woche 1 – Initialkauf
# ----------------------------
w0 = weeks[0]
m0 = week_maps[w0]
alloc = equity / TOP_N

for t in m0["ordered"][:TOP_N]:
    d, px = price_on_or_after(t, w0)
    if px is None:
        continue
    shares = alloc / px
    holdings[t] = shares
    log_entry(t, d, px, m0["score"].get(t, np.nan))

equity = mark_to_market(w0)
equity_curve.append({"Date": w0, "Equity": equity, "Holdings": len(holdings)})

# ----------------------------
# Folgewochen
# ----------------------------
for w in weeks[1:]:
    m = week_maps[w]

    # Exits: Rank > 30
    for t in list(holdings.keys()):
        if m["rank"].get(t, 9999) > HOLD_MAX_RANK:
            d, px = price_on_or_after(t, w)
            if px is not None:
                log_exit(t, d, px, m["score"].get(t, np.nan))
            holdings.pop(t)

    # Nachkäufe
    if len(holdings) < TOP_N:
        equity = mark_to_market(w)
        alloc = equity / TOP_N
        for t in m["ordered"]:
            if t not in holdings:
                d, px = price_on_or_after(t, w)
                if px is None:
                    continue
                holdings[t] = alloc / px
                log_entry(t, d, px, m["score"].get(t, np.nan))
                if len(holdings) == TOP_N:
                    break

    equity = mark_to_market(w)
    equity_curve.append({"Date": w, "Equity": equity, "Holdings": len(holdings)})

# ----------------------------
# Alle offenen Positionen schließen (letzter Handelstag)
# ----------------------------
final_date = prices.index.max()

for t in list(holdings.keys()):
    d, px = price_on_or_after(t, final_date)
    if px is not None:
        log_exit(t, d, px, np.nan)

holdings.clear()

equity = mark_to_market(final_date)
equity_curve.append({"Date": final_date, "Equity": equity, "Holdings": 0})

# ============================================================
# EXPORT
# ============================================================

events_df = pd.DataFrame(events).sort_values("Date").reset_index(drop=True)
equity_df = pd.DataFrame(equity_curve)

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    events_df.to_excel(writer, sheet_name="TradeEvents", index=False)
    equity_df.to_excel(writer, sheet_name="EquityCurve", index=False)

print("=== BACKTEST FERTIG ===")
print(f"Datei: {OUTPUT_XLSX}")
print(f"Events: {len(events_df)}")
print(f"Final Equity: {equity:.4f}")
