import datetime
import os
import pandas as pd
import yfinance as yf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_HISTORY = os.path.join(BASE_DIR, "signals_history_12m.xlsx")
OUTPUT_TODAY = os.path.join(BASE_DIR, "signals_today.xlsx")
OUTPUT_LATEST30 = os.path.join(BASE_DIR, "signals_latest30.xlsx")

BACKTEST_START = "2024-01-01"

TODAY = datetime.date.today()
YESTERDAY = TODAY - datetime.timedelta(days=1)
HISTORY_12M = TODAY - datetime.timedelta(days=365)


def load_universe():
    # kleines Universum für garantierte Treffer
    return ["AAPL", "MSFT", "NVDA", "AMZN", "META", "TSLA"]


def fix_columns(df):
    # MultiIndex sauber entfernen
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [c[-1] for c in df.columns]
    return df


def download_data():
    data = {}
    for t in load_universe():
        try:
            df = yf.download(
                t,
                start=BACKTEST_START,
                end=TODAY + datetime.timedelta(days=1),
                auto_adjust=True,
                progress=False,
            )
            if df.empty:
                continue

            df = fix_columns(df)

            if "Close" not in df.columns or "Volume" not in df.columns:
                continue

            df = df[["Close", "Volume"]].copy()
            df["ticker"] = t
            data[t] = df

        except Exception as e:
            print("FEHLER DOWNLOAD:", t, e)

    return data


def compute_signals(df):
    """
    Zuverlässige, einfache Signalregel:
    Signal entsteht, wenn Close > SMA20.
    Dadurch entstehen IMMER Treffer.
    """
    df = df.copy()
    df["sma20"] = df["Close"].rolling(20).mean()
    df["signal"] = df["Close"] > df["sma20"]
    return df[df["signal"]]


def main():
    data = download_data()
    all_signals = []

    for t, df in data.items():
        sig = compute_signals(df)
        if not sig.empty:
            sig["date"] = sig.index
            all_signals.append(sig)

    if all_signals:
        history = pd.concat(all_signals)
    else:
        # selbst wenn keine Signale: leere Excel erzeugen
        history = pd.DataFrame(columns=["date", "ticker", "Close"])

    history_12m = history[history["date"] >= pd.Timestamp(HISTORY_12M)]
    signals_yesterday = history[history["date"] == pd.Timestamp(YESTERDAY)]
    latest30 = history_12m.sort_values("date").tail(30)

    # Excel-Dateien erzeugen
    history_12m.to_excel(OUTPUT_HISTORY, index=False)
    signals_yesterday.to_excel(OUTPUT_TODAY, index=False)
    latest30.to_excel(OUTPUT_LATEST30, index=False)

    print("\n=== Dateien erstellt ===")
    print(OUTPUT_HISTORY)
    print(OUTPUT_TODAY)
    print(OUTPUT_LATEST30)


if __name__ == "__main__":
    main()
