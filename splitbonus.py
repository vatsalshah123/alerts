import os
import csv
import time
import yfinance as yf

TICKERS = ["TCS.NS", "RELIANCE.NS", "INFY.NS", "TEMBO.NS","ANANDRATHI.NS","NESTLEIND.NS","ZFCVINDIA.NS"]
SPLITS_CSV = "data/splits.csv"


def load_existing_keys(csv_path, key_fields):
    keys = set()
    if os.path.exists(csv_path):
        with open(csv_path, "r", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames and all(k in reader.fieldnames for k in key_fields):
                for row in reader:
                    keys.add(tuple(row[k] for k in key_fields))
            else:
                print(f"Warning: {csv_path} has unexpected headers {reader.fieldnames}, ignoring existing file.")
    return keys

def append_rows(csv_path, rows, fieldnames):
    if not rows:
        return
    file_exists = os.path.exists(csv_path)
    with open(csv_path, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)


def fetch_splits():
    os.makedirs("data", exist_ok=True)
    existing_keys = load_existing_keys(SPLITS_CSV, ["ticker", "action_date"])

    new_rows = []
    for ticker in TICKERS:
        t = yf.Ticker(ticker)
        splits = t.splits

        if splits is not None and not splits.empty:
            for date_idx, ratio in splits.items():
                action_date = date_idx.date().isoformat()
                key = (ticker, action_date)
                if key in existing_keys:
                    continue
                new_rows.append({
                    "ticker": ticker,
                    "action_date": action_date,
                    "ratio": float(ratio)
                    # Note: this covers both true splits and bonus issues —
                    # Yahoo doesn't distinguish between them
                })
                existing_keys.add(key)

        time.sleep(1)  # avoid rate limiting across many tickers

    append_rows(SPLITS_CSV, new_rows, ["ticker", "action_date", "ratio"])
    print(f"Added {len(new_rows)} new split/bonus rows.")


if __name__ == "__main__":
    fetch_splits()