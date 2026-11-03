"""
NSE All Stock List Fetcher
Fetches all equity symbols from NSE bhavcopy and splits them into text files
of 500 symbols each: nse_stocks1.txt (1-500), nse_stocks2.txt (501-1000), ...
Excludes symbols containing ETF, BEES, CASE, NIFTY.
"""

import os
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from nse import NSE

EXCLUDED = ("ETF", "BEES", "CASE", "NIFTY")
CHUNK_SIZE = 500


def is_excluded(symbol: str) -> bool:
    upper = symbol.upper()
    return any(token in upper for token in EXCLUDED)


def get_all_equity_symbols(nse: NSE) -> list[str]:
    for days_back in range(6):
        target_date = datetime.now() - timedelta(days=days_back)
        try:
            path = nse.equityBhavcopy(target_date)
            df = pd.read_csv(path)
            eq = df[(df["SctySrs"] == "EQ") & (df["FinInstrmTp"] == "STK")]
            symbols = sorted(
                str(s) for s in eq["TckrSymb"].tolist() if not is_excluded(str(s))
            )
            if symbols:
                if os.path.exists(path):
                    os.remove(path)
                return symbols
        except Exception:
            continue
    return []


def save_chunks(symbols: list[str]) -> None:
    total_chunks = (len(symbols) + CHUNK_SIZE - 1) // CHUNK_SIZE
    for i in range(total_chunks):
        start_idx = i * CHUNK_SIZE
        end_idx = min((i + 1) * CHUNK_SIZE, len(symbols))
        chunk = symbols[start_idx:end_idx]
        filename = f"nse_stocks{i + 1}.txt"

        with open(filename, "w") as f:
            for s in chunk:
                f.write(f"{s}\n")

        print(f"Created {filename} with {len(chunk)} stocks ({start_idx + 1}-{end_idx})")

    print(f"\nTotal files created: {total_chunks}")


def main():
    print("\nNSE All Stock List Fetcher")
    with NSE(download_folder=Path.cwd(), server=False, timeout=20) as nse:
        stocks = get_all_equity_symbols(nse)

    if not stocks:
        print("No symbols fetched.")
        return

    save_chunks(stocks)


if __name__ == "__main__":
    main()
