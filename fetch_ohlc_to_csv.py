import os
import csv
import pandas as pd
import yfinance as yf
from datetime import datetime

TICKERS = ["SCOOBEEDAY.BO"]
CSV_PATH = "data/daily_ohlc2.csv"

def load_existing_keys():
    """Read existing (ticker, date) pairs to avoid duplicate rows."""
    keys = set()
    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                keys.add((row["ticker"], row["trade_date"]))
    return keys

def fetch_and_append():
    os.makedirs("data", exist_ok=True)
    existing_keys = load_existing_keys()
    file_exists = os.path.exists(CSV_PATH)

    new_rows = []
    for ticker in TICKERS:
        data = yf.download(ticker, period="20y", interval="1d", auto_adjust=False, progress=False)
        if data.empty:
            print(f"No data for {ticker}")
            continue

        if isinstance(data.columns, pd.MultiIndex):
            data.columns = data.columns.get_level_values(0)

        # Loop through the window in case a day or two was missed previously
        for date_idx, row in data.iterrows():
            print(row["Open"])
            trade_date = date_idx.date().isoformat()
            key = (ticker, trade_date)
            if key in existing_keys:
                continue
            new_rows.append({
                "ticker": ticker,
                "trade_date": trade_date,
                "open_price": row["Open"],
                "high_price": row["High"],
                "low_price": row["Low"],
                "close_price": row["Close"],
            })
            existing_keys.add(key)

    if not new_rows:
        print("No new rows to add.")
        return

    with open(CSV_PATH, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "ticker", "trade_date", "open_price", "high_price",
            "low_price", "close_price"
        ])
        if not file_exists:
            writer.writeheader()
        writer.writerows(new_rows)

    print(f"Appended {len(new_rows)} new rows.")

if __name__ == "__main__":
    fetch_and_append()