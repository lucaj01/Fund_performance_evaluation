
import yfinance as yf
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

# ==========================================
# 1. USER INPUTS
# ==========================================

print("=== HISTORICAL ADJUSTED PRICES DOWNLOADER ===")

ticker1 = input("Enter first ticker: ").strip().upper()
ticker2 = input("Enter second ticker: ").strip().upper()
ticker3 = input("Enter third ticker: ").strip().upper()

start_date = input("Start date (YYYY-MM-DD): ").strip()
end_date = input("End date (YYYY-MM-DD): ").strip()

tickers = [ticker1, ticker2, ticker3]

# Validate dates
start = datetime.strptime(start_date, "%Y-%m-%d")
end = datetime.strptime(end_date, "%Y-%m-%d")

if end < start:
    raise ValueError("End date must be after start date.")

# Yahoo Finance excludes the end date
end_download = (end + timedelta(days=1)).strftime("%Y-%m-%d")

# ==========================================
# 2. EXCEL FILE
# ==========================================

output_path = Path.home() / "Downloads" / "Historical_Adjusted_Prices.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

    for i, ticker in enumerate(tickers, start=1):

        print(f"\nDownloading {ticker}...")

        data = yf.download(
            ticker,
            start=start_date,
            end=end_download,
            interval="1d",
            auto_adjust=False,
            progress=False
        )

        # Flatten columns if necessary
        if isinstance(data.columns, pd.MultiIndex):
            data.columns = data.columns.get_level_values(0)

        if data.empty:
            print(f"No data found for {ticker}.")
            prices = pd.DataFrame(columns=["Date", "Adj Close"])

        elif "Adj Close" not in data.columns:
            print(f"Adjusted Close not available for {ticker}.")
            prices = pd.DataFrame(columns=["Date", "Adj Close"])

        else:
            # Keep only adjusted closing prices
            prices = data[["Adj Close"]].reset_index()

            # Rename date column
            prices = prices.rename(
                columns={prices.columns[0]: "Date"}
            )

            # Remove timezone if present
            prices["Date"] = pd.to_datetime(
                prices["Date"], utc=True
            ).dt.tz_localize(None)

        # Each ticker gets a separate worksheet
        sheet_name = f"Stock_{i}"

        prices.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

        # Format Excel date column
        worksheet = writer.sheets[sheet_name]
        worksheet.column_dimensions["A"].width = 15
        worksheet.column_dimensions["B"].width = 20

        for cell in worksheet["A"][1:]:
            cell.number_format = "yyyy-mm-dd"

        for cell in worksheet["B"][1:]:
            cell.number_format = "0.0000"

        print(f"{ticker} saved in sheet {sheet_name}")

print("\nDownload completed!")
print(f"Excel file saved at: {output_path}")
