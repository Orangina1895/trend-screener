# Hebel_rechner.py
# ------------------------------------------------------------
# Daily 2x Leverage Simulation (path dependent)
# mit Verlusttopf-Steuerlogik (25,6 %)
# ------------------------------------------------------------

import pandas as pd
import numpy as np
import yfinance as yf

# =========================
# KONFIGURATION
# =========================
INPUT_XLSX = "Backtest_result.xlsx"
SHEET_NAME = "TradeEvents"

START_EQUITY = 10000.0
LEVERAGE = 2.0
TAX_RATE = 0.256

# Falls der Einsatz pro Trade in einer bestimmten Spalte steht:
# (bei dir waren das z.B. Werte wie 1159,25)
POSITION_COL = "Unnamed: 9"   # ggf. anpassen!

# =========================
# 1. TRADES LADEN
# =========================
df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME)

df = df.rename(
    columns={
        "Date": "date",
        "Ticker": "ticker",
        "Action": "action"
    }
)

df["date"] = pd.to_datetime(df["date"])
df["action"] = df["action"].str.upper()

df = df.sort_values("date").reset_index(drop=True)

# =========================
# 2. ENTRY / EXIT PAARE (FIFO)
# =========================
open_entries = {}
pairs = []

for _, row in df.iterrows():
    t = row["ticker"]
    a = row["action"]

    if a == "ENTRY":
        open_entries.setdefault(t, []).append(row["date"])

    elif a == "EXIT":
        if t in open_entries and len(open_entries[t]) > 0:
            entry_date = open_entries[t].pop(0)
            exit_date = row["date"]
            pairs.append((t, entry_date, exit_date))

pairs_df = pd.DataFrame(pairs, columns=["ticker", "entry", "exit"])

if pairs_df.empty:
    raise RuntimeError("Keine gültigen ENTRY/EXIT-Paare gefunden.")

# =========================
# 3. POSITIONSGRÖSSE JE EXIT
# =========================
exit_rows = df[df["action"] == "EXIT"].copy()
exit_rows["key"] = list(zip(exit_rows["ticker"], exit_rows["date"]))

size_map = dict(zip(exit_rows["key"], exit_rows[POSITION_COL]))

# =========================
# 4. SIMULATION
# =========================
equity = START_EQUITY
loss_pot = 0.0

results = []

for _, trade in pairs_df.sort_values("exit").iterrows():
    ticker = trade["ticker"]
    entry = trade["entry"]
    exit_ = trade["exit"]

    size = size_map.get((ticker, exit_), np.nan)

    if np.isnan(size):
        # Fallback (sollte idealerweise nicht passieren)
        size = equity * 0.1

    size = float(size)

    # -------------------------
    # Kurse laden
    # -------------------------
    prices = yf.download(
        ticker,
        start=(entry - pd.Timedelta(days=5)).date(),
        end=(exit_ + pd.Timedelta(days=1)).date(),
        auto_adjust=True,
        progress=False
    )["Close"].dropna()

    prices = prices.loc[entry:exit_]

    if len(prices) < 2:
        continue

    # -------------------------
    # Daily Returns + Hebel
    # -------------------------
    daily_returns = prices.pct_change().dropna()
    lev_returns = (daily_returns * LEVERAGE).clip(lower=-1.0)

    # -------------------------
    # PFADABHÄNGIGE ENTWICKLUNG
    # -------------------------
    value = float(size)

    for r in lev_returns.values:
        value *= (1.0 + float(r))

    value = float(value)
    pl = value - size

    # -------------------------
    # STEUERLOGIK (Verlusttopf)
    # -------------------------
    tax = 0.0

    if pl < 0:
        loss_pot += -pl
    else:
        offset = min(loss_pot, pl)
        loss_pot -= offset
        taxable = pl - offset
        tax = taxable * TAX_RATE

    equity += (pl - tax)

    results.append({
        "Ticker": ticker,
        "Entry": entry,
        "Exit": exit_,
        "Position_Start_EUR": round(size, 2),
        "Position_End_EUR": round(value, 2),
        "P/L_EUR": round(pl, 2),
        "Tax_Paid": round(tax, 2),
        "Loss_Pot_After": round(loss_pot, 2),
        "Equity_After": round(equity, 2)
    })

# =========================
# 5. OUTPUT
# =========================
out = pd.DataFrame(results)

summary = pd.DataFrame([{
    "Start_Equity": START_EQUITY,
    "End_Equity": equity,
    "Total_Return_%": (equity / START_EQUITY - 1) * 100,
    "Tax_Rate": TAX_RATE,
    "Leverage": LEVERAGE,
    "Trades": len(out)
}])

with pd.ExcelWriter("Backtest_result_daily_2x.xlsx", engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    out.to_excel(writer, sheet_name="Trades", index=False)

print("Simulation abgeschlossen")
print("End-Equity:", round(equity, 2))
