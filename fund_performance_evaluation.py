
from pathlib import Path
import math
import re

import numpy as np
import pandas as pd
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ==========================================
# 1. CONFIGURATION
# ==========================================

FOLDER = Path(__file__).resolve().parent

OUTPUT = (
    FOLDER / "Fund_Evaluation_Total_Return.xlsx"
)

# Expected worksheets in the price workbook.
# Stock_1/2/3 correspond to:
# WOFNX, ^SP500TR, ^RUOTR.

SHEETS = {
    "WOFNX": ["WOFNX", "Stock_1"],
    "S&P 500 TR": [
        "SP500TR", "^SP500TR", "SP500", "Stock_2"
    ],
    "Russell 2000 Growth TR": [
        "RUOTR", "^RUOTR",
        "RUSSEL2000GWT", "RUSSELL2000GWT",
        "Stock_3"
    ]
}

ASSETS = list(SHEETS)

FUND = "WOFNX"
RUSSELL = "Russell 2000 Growth TR"
SP500 = "S&P 500 TR"

WEEKS_PER_YEAR = 52

PERIODS = {
    "Year 1": ("2024-11-01", "2025-10-31"),
    "Year 2": ("2025-10-31", "2026-10-30"),
    "Full Period": ("2024-11-01", "2026-10-30")
}


# ==========================================
# 2. LOCATE INPUT FILES
# ==========================================

def input_file(filename):

    exact = FOLDER / filename

    if exact.exists():
        return exact

    stem = Path(filename).stem

    matches = sorted(
        FOLDER.glob(
            f"{stem}(*){Path(filename).suffix}"
        )
    )

    if not matches:
        raise FileNotFoundError(
            f"Missing {filename} in {FOLDER}"
        )

    print(f"Using: {matches[0].name}")

    return matches[0]


def column_name(frame, options):

    names = {
        re.sub(r"[^a-z0-9]", "", str(c).lower()): c
        for c in frame.columns
    }

    return next(
        (names[x] for x in options if x in names),
        None
    )


# ==========================================
# 3. LOAD HISTORICAL PRICES
# ==========================================

def read_prices(filename):

    excel = pd.ExcelFile(filename)

    price_series = {}
    used = set()
    checks = []

    for asset, choices in SHEETS.items():

        sheet = next(
            (
                x for x in choices
                if x in excel.sheet_names
                and x not in used
            ),
            None
        )

        if sheet is None:
            raise ValueError(
                f"Missing worksheet for {asset}. "
                f"Available: {excel.sheet_names}"
            )

        used.add(sheet)

        frame = pd.read_excel(
            excel,
            sheet_name=sheet
        )

        date_col = column_name(
            frame, ["date", "datetime"]
        )

        price_col = column_name(
            frame, ["adjclose", "adjustedclose"]
        )

        if date_col is None or price_col is None:
            raise ValueError(
                f"{sheet} requires Date and Adj Close."
            )

        dates = pd.to_datetime(
            frame[date_col],
            errors="coerce",
            utc=True
        ).dt.tz_localize(None).dt.normalize()

        values = pd.to_numeric(
            frame[price_col],
            errors="coerce"
        )

        clean = pd.DataFrame({
            "Date": dates,
            "Value": values
        }).dropna()

        clean = clean.loc[
            clean["Value"] > 0
        ].drop_duplicates(
            "Date", keep="last"
        )

        if clean.empty:
            raise ValueError(
                f"No valid prices in {sheet}"
            )

        price_series[asset] = (
            clean.set_index("Date")["Value"]
            .sort_index()
            .rename(asset)
        )

        checks.append({
            "Check": f"{asset} worksheet",
            "Result": sheet
        })

        checks.append({
            "Check": f"{asset} observations",
            "Result": len(clean)
        })

    # Align observations across all assets
    common = pd.concat(
        list(price_series.values()),
        axis=1,
        join="inner"
    ).dropna()

    common = common.loc[
        "2024-11-01":"2026-10-30"
    ]

    if common.empty:
        raise ValueError(
            "No common observations in the period."
        )

    # Weekly prices: Friday-ending weeks
    weekly = common.resample("W-FRI").last()

    expected = pd.date_range(
        "2024-11-01",
        "2026-10-30",
        freq="W-FRI"
    )

    if (
        not weekly.index.equals(expected)
        or weekly.isna().any().any()
    ):
        raise ValueError(
            "Incomplete weekly prices. "
            "Check the three datasets."
        )

    if len(weekly) != 105:
        raise ValueError(
            "Expected 105 weekly price observations."
        )

    checks.append({
        "Check": "Common daily observations",
        "Result": len(common)
    })

    checks.append({
        "Check": "Weekly prices",
        "Result": len(weekly)
    })

    checks.append({
        "Check": "Input file",
        "Result": filename.name
    })

    checks.append({
        "Check": "Benchmark source verification",
        "Result": (
            "Verify downloaded tickers are "
            "^SP500TR and ^RUOTR"
        )
    })

    return weekly, pd.DataFrame(checks)


# ==========================================
# 4. LOAD RISK-FREE RATES
# ==========================================

def read_risk_free(filename):

    rf = pd.read_excel(
        filename,
        sheet_name="Risk_Free"
    )

    required = [
        "Period",
        "Observation Date",
        "Annual RF (%)",
        "Weekly RF (%)"
    ]

    if any(c not in rf.columns for c in required):
        raise ValueError(
            f"Risk-free file must contain: {required}"
        )

    rf["Period"] = (
        rf["Period"].astype(str).str.strip()
    )

    if rf["Period"].duplicated().any():
        raise ValueError(
            "Duplicate periods in risk-free file."
        )

    rf = rf.set_index("Period")

    for period in PERIODS:

        if period not in rf.index:
            raise ValueError(
                f"Missing risk-free for {period}"
            )

        start_date = (
            "2025-11-01"
            if period == "Year 2"
            else "2024-11-01"
        )

        observation_date = pd.to_datetime(
            rf.loc[period, "Observation Date"]
        )

        if (
            pd.isna(observation_date)
            or observation_date > pd.Timestamp(start_date)
        ):
            raise ValueError(
                f"Invalid risk-free observation "
                f"for {period}"
            )

        for col in [
            "Annual RF (%)",
            "Weekly RF (%)"
        ]:
            value = float(rf.loc[period, col])

            if not np.isfinite(value):
                raise ValueError(
                    f"Invalid {col} for {period}"
                )

    return rf


# ==========================================
# 5. FINANCIAL FUNCTIONS
# ==========================================

def safe_ratio(numerator, denominator):

    if pd.notna(denominator) and denominator > 0:
        return numerator / denominator

    return np.nan


def evaluate(weekly, rf):

    performance = []
    risk_adjusted = []
    comparisons = []
    four_week_tables = []

    for period, (start, end) in PERIODS.items():

        prices = weekly.loc[start:end]

        # Weekly simple returns
        returns = prices.pct_change(
            fill_method=None
        ).iloc[1:]

        expected_n = (
            104 if period == "Full Period" else 52
        )

        if len(returns) != expected_n:
            raise ValueError(
                f"{period}: expected {expected_n} "
                f"weekly returns."
            )

        # Rolling 4-week holding period returns
        rolling_4w = (
            prices.div(prices.shift(4)).sub(1)
        )

        four_week_tables.append(
            rolling_4w.assign(
                Period=period
            ).reset_index()
        )

        # Risk-free at the beginning of the period
        rf_weekly = (
            float(rf.loc[period, "Weekly RF (%)"])
            / 100
        )

        rf_annual = (
            float(rf.loc[period, "Annual RF (%)"])
            / 100
        )

        # Active weekly returns vs Russell Growth TR
        active = (
            returns[FUND] - returns[RUSSELL]
        )

        fund_sd = float(
            returns[FUND].std(ddof=1)
        )

        active_sd = float(
            active.std(ddof=1)
        )

        # Annualized Sharpe ratio
        sharpe = safe_ratio(
            float(returns[FUND].mean()) - rf_weekly,
            fund_sd
        ) * math.sqrt(WEEKS_PER_YEAR)

        # Annualized Information Ratio
        information_ratio = safe_ratio(
            float(active.mean()),
            active_sd
        ) * math.sqrt(WEEKS_PER_YEAR)

        # Tracking error
        tracking_error = (
            active_sd * math.sqrt(WEEKS_PER_YEAR)
        )

        risk_adjusted.append({
            "Period": period,
            "Annual RF": rf_annual,
            "Weekly RF": rf_weekly,
            "Sharpe WOFNX": sharpe,
            "Information Ratio vs Russell TR": information_ratio,
            "Tracking error annualized": tracking_error,
            "Mean weekly active return": float(active.mean())
        })

        # ==================================
        # ASSET PERFORMANCE
        # ==================================

        for asset in ASSETS:

            p = prices[asset]
            r = returns[asset]

            # Cumulative holding period return
            cumulative_return = float(
                p.iloc[-1] / p.iloc[0] - 1
            )

            # Annualized geometric return
            years = (
                p.index[-1] - p.index[0]
            ).days / 365.25

            cagr = (
                (1 + cumulative_return) ** (1 / years)
                - 1
            )

            # Weekly standard deviation
            weekly_sd = float(
                r.std(ddof=1)
            )

            # Annualized volatility
            annual_sd = (
                weekly_sd * math.sqrt(WEEKS_PER_YEAR)
            )

            # Worst rolling four-week return
            window = rolling_4w[asset].dropna()

            worst_return = float(
                window.min()
            )

            worst_end = window.idxmin()

            worst_start = (
                worst_end - pd.Timedelta(weeks=4)
            )

            maximum_loss = max(
                0, -worst_return
            )

            performance.append({
                "Period": period,
                "Asset": asset,
                "Start": p.index[0].date(),
                "End": p.index[-1].date(),
                "N weeks": len(r),
                "Cumulative return": cumulative_return,
                "Annualized return (CAGR)": cagr,
                "Weekly SD": weekly_sd,
                "Annualized SD": annual_sd,
                "Worst rolling 4-week return": worst_return,
                "Maximum 4-week loss": maximum_loss,
                "Worst window start": worst_start.date(),
                "Worst window end": worst_end.date()
            })

        # ==================================
        # BENCHMARK COMPARISON
        # ==================================

        for benchmark in [RUSSELL, SP500]:

            b = returns[benchmark]

            relative_return = (
                prices[FUND].iloc[-1]
                / prices[FUND].iloc[0]
                -
                prices[benchmark].iloc[-1]
                / prices[benchmark].iloc[0]
            )

            correlation = float(
                returns[FUND].corr(b)
            )

            beta = safe_ratio(
                float(returns[FUND].cov(b)),
                float(b.var(ddof=1))
            )

            comparisons.append({
                "Period": period,
                "Benchmark": benchmark,
                "Fund minus benchmark cumulative return (pp)": (
                    relative_return * 100
                ),
                "Correlation": correlation,
                "Beta": beta
            })

    return (
        pd.DataFrame(performance),
        pd.DataFrame(risk_adjusted),
        pd.DataFrame(comparisons),
        pd.concat(four_week_tables, ignore_index=True)
    )


# ==========================================
# 6. AUTOMATIC INTERPRETATION
# ==========================================

def write_comments(performance, ratios):

    comments = []

    for period in PERIODS:

        table = performance.loc[
            performance["Period"].eq(period)
        ].set_index("Asset")

        r = ratios.loc[
            ratios["Period"].eq(period)
        ].iloc[0]

        fund = table.loc[FUND]
        russell = table.loc[RUSSELL]
        sp500 = table.loc[SP500]

        comment = (
            f"WOFNX returned "
            f"{fund['Cumulative return']:.2%}; "
            f"Russell 2000 Growth TR returned "
            f"{russell['Cumulative return']:.2%}; "
            f"S&P 500 TR returned "
            f"{sp500['Cumulative return']:.2%}. "

            f"Annualized volatilities were "
            f"{fund['Annualized SD']:.2%}, "
            f"{russell['Annualized SD']:.2%}, "
            f"and {sp500['Annualized SD']:.2%}, respectively. "

            f"Worst four-week losses were "
            f"{fund['Maximum 4-week loss']:.2%}, "
            f"{russell['Maximum 4-week loss']:.2%}, "
            f"and {sp500['Maximum 4-week loss']:.2%}. "

            f"The fund Sharpe ratio was "
            f"{r['Sharpe WOFNX']:.3f}, "
            f"and its Information Ratio against "
            f"Russell 2000 Growth TR was "
            f"{r['Information Ratio vs Russell TR']:.3f}."
        )

        comments.append({
            "Section": period,
            "Comment": comment
        })

    comments.append({
        "Section": "Benchmark suitability",
        "Comment": (
            "The Russell 2000 Growth Index represents "
            "US small-cap growth stocks, while the S&P 500 "
            "primarily represents large-cap US equities. "
            "The Russell index is therefore more closely "
            "aligned with the small-cap growth mandate "
            "specified in the assignment. However, the "
            "actual portfolio composition and investment "
            "style of WOFNX should be examined before "
            "assessing benchmark suitability for this fund."
        )
    })

    comments.append({
        "Section": "Total Return methodology",
        "Comment": (
            "The analysis uses adjusted NAVs for WOFNX "
            "and Total Return indices for both benchmarks. "
            "S&P 500 Total Return (^SP500TR) and Russell "
            "2000 Growth Total Return (^RUOTR) incorporate "
            "reinvested dividends. This provides a more "
            "consistent basis for performance comparison, "
            "assuming the fund adjusted NAV correctly "
            "reflects distributions."
        )
    })

    comments.append({
        "Section": "Risk-free methodology",
        "Comment": (
            "Risk-free rates are taken at the beginning "
            "of each evaluation horizon. The one-year "
            "periods use the 12-month Treasury Bill "
            "(FRED DTB1YR), while the two-year evaluation "
            "uses the two-year Treasury yield (FRED DGS2) "
            "as a proxy. The latter is not the USD "
            "24-month Euromarket rate mentioned in "
            "the assignment."
        )
    })

    return pd.DataFrame(comments)


# ==========================================
# 7. EXPORT RESULTS TO EXCEL
# ==========================================

def create_report(
    weekly,
    rf,
    performance,
    ratios,
    comparison,
    four_week,
    checks
):

    # Normalized wealth index
    wealth_index = (
        weekly / weekly.iloc[0] * 100
    )

    # Weekly returns
    weekly_returns = weekly.pct_change(
        fill_method=None
    ).iloc[1:]

    # Rolling four-week returns
    rolling = (
        weekly.div(weekly.shift(4)).sub(1)
    )

    comments = write_comments(
        performance, ratios
    )

    methodology = pd.DataFrame([
        (
            "Investment period",
            "November 1, 2024 to October 30, 2026"
        ),
        (
            "Year 1",
            "2024-11-01 to 2025-10-31"
        ),
        (
            "Year 2",
            "2025-10-31 close to 2026-10-30 close; "
            "November 1, 2025 was a Saturday"
        ),
        (
            "Benchmarks",
            "Russell 2000 Growth TR (^RUOTR) "
            "and S&P 500 TR (^SP500TR)"
        ),
        (
            "Weekly observations",
            "Last common available closing price "
            "in each Friday-ending week"
        ),
        (
            "Cumulative return",
            "Final adjusted price / initial adjusted price - 1"
        ),
        (
            "Annualized volatility",
            "Weekly sample standard deviation * sqrt(52)"
        ),
        (
            "Maximum four-week loss",
            "Maximum of zero and negative minimum "
            "rolling four-week holding-period return"
        ),
        (
            "Sharpe Ratio",
            "sqrt(52) * (mean weekly fund return "
            "- weekly risk-free) / weekly fund SD"
        ),
        (
            "Information Ratio",
            "sqrt(52) * mean(weekly fund return "
            "- weekly Russell TR return) / "
            "sample SD of active returns"
        ),
        (
            "Risk-free",
            "DTB1YR for each annual period and DGS2 "
            "as a proxy for the two-year period"
        ),
        (
            "Total Return verification",
            "Check that the benchmark worksheets "
            "contain ^RUOTR and ^SP500TR, not "
            "^RUO and ^GSPC"
        )
    ], columns=[
        "Topic",
        "Definition / qualification"
    ])

    tables = [
        ("Performance", performance),
        ("Ratios", ratios),
        ("Comparison", comparison),
        ("Weekly_Prices", weekly.reset_index()),
        ("Weekly_Returns", weekly_returns.reset_index()),
        ("Wealth_Base_100", wealth_index.reset_index()),
        ("Rolling_4W", rolling.reset_index()),
        ("4W_By_Period", four_week),
        ("Risk_Free_Inputs", rf.reset_index()),
        ("Comments", comments),
        ("Methodology", methodology),
        ("Data_Checks", checks)
    ]

    with pd.ExcelWriter(
        OUTPUT,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd"
    ) as writer:

        for sheet_name, table in tables:

            table.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            ws = writer.sheets[sheet_name]

            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False

            # Header formatting
            for cell in ws[1]:

                cell.fill = PatternFill(
                    "solid",
                    fgColor="17324D"
                )

                cell.font = Font(
                    color="FFFFFF",
                    bold=True
                )

                cell.alignment = Alignment(
                    wrap_text=True,
                    vertical="center"
                )

            ws.row_dimensions[1].height = 32

            percentage_columns = {
                "Cumulative return",
                "Annualized return (CAGR)",
                "Weekly SD",
                "Annualized SD",
                "Worst rolling 4-week return",
                "Maximum 4-week loss",
                "Annual RF",
                "Weekly RF",
                "Tracking error annualized",
                "Mean weekly active return"
            }

            for j, col in enumerate(
                table.columns, start=1
            ):

                letter = get_column_letter(j)

                ws.column_dimensions[letter].width = (
                    100
                    if col in (
                        "Comment",
                        "Definition / qualification"
                    )
                    else min(
                        44,
                        max(16, len(str(col)) + 3)
                    )
                )

                is_percentage = (
                    col in percentage_columns
                    or (
                        sheet_name in (
                            "Weekly_Returns",
                            "Rolling_4W",
                            "4W_By_Period"
                        )
                        and col in ASSETS
                    )
                )

                for row in range(
                    2, ws.max_row + 1
                ):

                    cell = ws.cell(row, j)

                    if (
                        is_percentage
                        and isinstance(
                            cell.value, (int, float)
                        )
                    ):
                        cell.number_format = "0.00%"

                    elif col in (
                        "Sharpe WOFNX",
                        "Information Ratio vs Russell TR",
                        "Beta",
                        "Correlation"
                    ):
                        cell.number_format = "0.000"

                    if col == "Comment":

                        cell.alignment = Alignment(
                            vertical="top",
                            wrap_text=True
                        )

                        ws.row_dimensions[row].height = 78

            # ==================================
            # PERFORMANCE CHART
            # ==================================

            if sheet_name == "Wealth_Base_100":

                chart = LineChart()

                chart.title = (
                    "Growth of 100: Fund vs TR Benchmarks"
                )

                chart.y_axis.title = "Portfolio Value"
                chart.x_axis.title = "Week"

                chart.height = 11
                chart.width = 19

                chart.add_data(
                    Reference(
                        ws,
                        min_col=2,
                        max_col=4,
                        min_row=1,
                        max_row=ws.max_row
                    ),
                    titles_from_data=True
                )

                chart.set_categories(
                    Reference(
                        ws,
                        min_col=1,
                        min_row=2,
                        max_row=ws.max_row
                    )
                )

                ws.add_chart(chart, "F2")

            # ==================================
            # FOUR-WEEK ROLLING RETURNS CHART
            # ==================================

            if sheet_name == "Rolling_4W":

                chart = LineChart()

                chart.title = (
                    "Rolling 4-Week Returns"
                )

                chart.height = 10
                chart.width = 18

                chart.add_data(
                    Reference(
                        ws,
                        min_col=2,
                        max_col=4,
                        min_row=1,
                        max_row=ws.max_row
                    ),
                    titles_from_data=True
                )

                chart.set_categories(
                    Reference(
                        ws,
                        min_col=1,
                        min_row=2,
                        max_row=ws.max_row
                    )
                )

                ws.add_chart(chart, "F2")


# ==========================================
# 8. MAIN PROGRAM
# ==========================================

def main():

    print(
        "=== FUND PERFORMANCE EVALUATION ==="
    )

    # Locate files
    price_file = input_file(
        "Historical_Adjusted_Prices.xlsx"
    )

    risk_file = input_file(
        "Risk_Free_Assignment.xlsx"
    )

    print("\nLoading historical prices...")

    weekly, checks = read_prices(
        price_file
    )

    print("Loading risk-free rates...")

    rf = read_risk_free(
        risk_file
    )

    print("Calculating performance metrics...")

    (
        performance,
        ratios,
        comparison,
        four_week
    ) = evaluate(
        weekly,
        rf
    )

    print("Creating Excel report...")

    create_report(
        weekly,
        rf,
        performance,
        ratios,
        comparison,
        four_week,
        checks
    )

    print("\n=== PERFORMANCE SUMMARY ===")

    print(
        performance[[
            "Period",
            "Asset",
            "Cumulative return",
            "Annualized SD",
            "Maximum 4-week loss"
        ]].to_string(index=False)
    )

    print("\n=== RISK-ADJUSTED PERFORMANCE ===")

    print(
        ratios[[
            "Period",
            "Sharpe WOFNX",
            "Information Ratio vs Russell TR"
        ]].to_string(index=False)
    )

    print("\nAnalysis completed successfully!")

    print(f"\nExcel file saved at:\n{OUTPUT}")

    print(
        "\nIMPORTANT: Verify that the Excel "
        "contains the newly downloaded "
        "^SP500TR and ^RUOTR series."
    )


if __name__ == "__main__":
    main()
