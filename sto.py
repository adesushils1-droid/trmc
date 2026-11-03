"""
NSE All Stock List Fetcher
Fetches all equity symbols from NSE bhavcopy and splits them into text files
of 500 symbols each: nse_stocks1.txt (1-500), nse_stocks2.txt (501-1000), ...
Excludes symbols containing ETF, BEES, CASE, NIFTY.
"""

import os
import time
import glob
import shutil
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from nse import NSE

EXCLUDED = ("ETF", "BEES", "CASE", "NIFTY", "BETA", "LIQUID","MIDCAP","NEXT","GOLD","SILVER")
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

        if i < total_chunks - 1:
            print("Resting 60s before next list...")
            time.sleep(00)

    print(f"\nTotal files created: {total_chunks}")
    create_apinse_copies(total_chunks)


def create_apinse_copies(total_chunks: int) -> None:
    """Create apinse1.py, apinse2.py, ... copies from apinse.py, each bound to one nse_stocksN.txt."""
    try:
        with open("apinse.py", "r") as f:
            template = f.read()
    except FileNotFoundError:
        print("apinse.py not found; skipping apinseN.py copy creation.")
        return

    stock_file_line = 'STOCKS_FILE = os.environ.get("STOCKS_FILE")  # e.g. "nse_stocks1.txt" for chunk mode; None = all'
    for i in range(1, total_chunks + 1):
        filename = f"apinse{i}.py"
        content = template.replace(
            stock_file_line,
            f'STOCKS_FILE = "nse_stocks{i}.txt"  # chunk mode: process only nse_stocks{i}.txt',
        )
        with open(filename, "w") as f:
            f.write(content)
        print(f"Created {filename} bound to nse_stocks{i}.txt")

    print(f"Created {total_chunks} apinse*.py copies from apinse.py")


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
