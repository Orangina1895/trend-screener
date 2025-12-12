import pandas as pd
import numpy as np
import yfinance as yf

TAX_RATE = 0.256
LEVERAGE = 2.0
INPUT_XLSX = "Backtest_result.xlsx"

df = pd.read_excel(INPUT_XLSX, sheet_name="TradeEvents")
df = df.rename(columns={"Date":"date", "Ticker":"ticker", "Action":"action"})
df["date"] = pd.to_datetime(df["date"])

# Pairs: pro Ticker nächstes EXIT nach ENTRY (FIFO)
pairs = []
open_pos = {}  # ticker -> list of entry dates
for _, r in df.sort_values("date").iterrows():
    t = r["ticker"]
    a = str(r["action"]).upper()
    if a == "ENTRY":
        open_pos.setdefault(t, []).append(r["date"])
    elif a == "EXIT" and t in open_pos and open_pos[t]:
        entry_date = open_pos[t].pop(0)
        exit_date = r["date"]
        pairs.append((t, entry_date, exit_date))

pairs_df = pd.DataFrame(pairs, columns=["ticker","entry","exit"])
if pairs_df.empty:
    raise ValueError("Keine ENTRY/EXIT-Paare gefunden.")

# Positionsgröße am Entry: in deiner Datei ist das i.d.R. die Spalte mit dem Einsatz.
# In deinem Sheet ist das sehr wahrscheinlich 'Unnamed: 9' (weil dort z.B. 1159.25 steht).
# Falls du eine saubere Spaltenüberschrift hast, hier ersetzen:
POSITION_COL = "Unnamed: 9"

# Mapping Einsatz je (ticker, exit_date) -> Einsatz
exit_rows = df[df["action"].str.upper() == "EXIT"].copy()
exit_rows["key"] = list(zip(exit_rows["ticker"], exit_rows["date"]))
size_map = dict(zip(exit_rows["key"], exit_rows[POSITION_COL]))

loss_pot = 0.0
equity = 10000.0

rows_out = []

for _, p in pairs_df.sort_values("exit").iterrows():
    ticker, entry, exit_ = p["ticker"], p["entry"], p["exit"]
    size = float(size_map.get((ticker, exit_), np.nan))
    if np.isnan(size):
        # Fallback: wenn Einsatz nicht gefunden wird, nimm Anteil am Depot (z.B. 1/Anzahl)
        size = equity * 0.1  # anpassen

    # Kurse laden (ein paar Tage Puffer)
    px = yf.download(ticker, start=(entry - pd.Timedelta(days=5)).date(),
                     end=(exit_ + pd.Timedelta(days=1)).date(),
                     auto_adjust=True, progress=False)["Close"].dropna()

    px = px.loc[entry:exit_]
    if len(px) < 2:
        continue

    rets = px.pct_change().dropna()
    lev_rets = LEVERAGE * rets
    # optional: Daily Floor bei -100% (damit (1+lev_ret) nicht < 0 wird)
    lev_rets = lev_rets.clip(lower=-1.0)

    # Pfadabhängige Entwicklung des Positionswertes
    value = size
    for d, r2x in lev_rets.items():
        value *= (1.0 + r2x)

    pl = value - size  # realisiert beim Exit

    # Steuer mit Verlusttopf (keine negative Steuer)
    tax = 0.0
    if pl < 0:
        loss_pot += -pl
    else:
        offset = min(loss_pot, pl)
        loss_pot -= offset
        taxable = pl - offset
        tax = taxable * TAX_RATE

    equity += (pl - tax)

    rows_out.append({
        "ticker": ticker,
        "entry": entry,
        "exit": exit_,
        "size_entry": size,
        "value_exit": value,
        "pl": pl,
        "tax_paid": tax,
        "loss_pot_after": loss_pot,
        "equity_after": equity
    })

out = pd.DataFrame(rows_out)
out.to_excel("Backtest_result_daily2x.xlsx", index=False)
print("Fertig. End-Equity:", equity)
