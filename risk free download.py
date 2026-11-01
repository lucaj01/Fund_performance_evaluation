
import pandas as pd
from pathlib import Path

# ==========================================
# 1. INVESTMENT PERIODS
# ==========================================

periods = [
    ("Year 1", "2024-11-01", "DTB1YR", 1),
    ("Year 2", "2025-11-01", "DTB1YR", 1),
    ("Full Period", "2024-11-01", "DGS2", 2)
]

print("=== RISK-FREE RATE DOWNLOADER ===")

# ==========================================
# 2. DOWNLOAD DATA FROM FRED
# ==========================================

def download_fred(series):

    url = (
        "https://fred.stlouisfed.org/graph/"
        f"fredgraph.csv?id={series}"
    )

    data = pd.read_csv(url, na_values=["."])

    data.columns = ["Date", "Rate"]

    data["Date"] = pd.to_datetime(data["Date"])
    data["Rate"] = pd.to_numeric(
        data["Rate"], errors="coerce"
    )

    return data.dropna(subset=["Rate"])


results = []
cache = {}

for period, start_date, series, years in periods:

    print(f"\nDownloading {series} for {period}...")

    if series not in cache:
        cache[series] = download_fred(series)

    data = cache[series]

    # Select latest available rate at investment start
    available = data[
        data["Date"] <= pd.Timestamp(start_date)
    ]

    if available.empty:
        raise ValueError(f"No data for {series}")

    selected = available.iloc[-1]

    rate = float(selected["Rate"]) / 100

    # ======================================
    # 3. CONVERT RATES
    # ======================================

    if series == "DTB1YR":

        # T-Bill quoted on discount basis
        # Assume 364 days to maturity
        days = 364

        holding_return = (
            1 / (1 - rate * days / 360) - 1
        )

        annual_rf = (
            (1 + holding_return) ** (365 / days) - 1
        )

    else:

        # 2Y Treasury yield: semiannual
        # compounding approximation
        annual_rf = (1 + rate / 2) ** 2 - 1

    # Equivalent weekly risk-free return
    weekly_rf = (1 + annual_rf) ** (1 / 52) - 1

    # Approximate cumulative period return
    period_rf = (1 + annual_rf) ** years - 1

    results.append({
        "Period": period,
        "Start Date": pd.Timestamp(start_date),
        "Observation Date": selected["Date"],
        "FRED Series": series,
        "Maturity (Years)": years,
        "Quoted Rate (%)": rate * 100,
        "Annual RF (%)": annual_rf * 100,
        "Weekly RF (%)": weekly_rf * 100,
        "Period RF (%)": period_rf * 100
    })

# ==========================================
# 4. EXPORT TO EXCEL
# ==========================================

result = pd.DataFrame(results)

output_path = (
    Path.home() / "Downloads" /
    "Risk_Free_Assignment.xlsx"
)

with pd.ExcelWriter(
    output_path, engine="openpyxl"
) as writer:

    result.to_excel(
        writer,
        sheet_name="Risk_Free",
        index=False
    )

    ws = writer.sheets["Risk_Free"]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for column in ws.columns:
        letter = column[0].column_letter
        ws.column_dimensions[letter].width = 22

    for row in ws.iter_rows(min_row=2):
        row[1].number_format = "yyyy-mm-dd"
        row[2].number_format = "yyyy-mm-dd"

        for cell in row[5:9]:
            cell.number_format = "0.0000"

print("\n=== RESULTS ===")
print(result.to_string(index=False))

print("\nExcel successfully created!")
print(f"File saved at: {output_path}")
