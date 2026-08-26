"""
ZERO-TOUCH DATA PIPELINE — Messy Raw Data -> Executive Excel Report
--------------------------------------------------------------------
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("data_pipeline")

# --------------------------------------------------------------------------
# CONFIG (overridable via CLI flags — see main())
# --------------------------------------------------------------------------
DEFAULT_INPUT = "sample_data/messy_sales_export.csv"
DEFAULT_OUTPUT = "executive_report.xlsx"
DATE_COLUMNS = ["order_date"]
NUMERIC_COLUMNS = ["quantity", "unit_price", "revenue"]
KEY_GROUP_COLUMN = "region"
REVENUE_COLUMN = "revenue"
REQUIRED_COLUMNS = {REVENUE_COLUMN}  # analysis hard-fails without these

# pandas >= 2.2 deprecated the "M" resample alias in favor of "ME".
# Pick whichever the installed version actually supports instead of
# hardcoding one and breaking on the other.
_PANDAS_MAJOR_MINOR = tuple(int(x) for x in pd.__version__.split(".")[:2])
MONTH_END_ALIAS = "ME" if _PANDAS_MAJOR_MINOR >= (2, 2) else "M"


@dataclass
class DataQualityReport:
    rows_in: int = 0
    rows_out: int = 0
    duplicates_removed: int = 0
    missing_values_filled: dict = field(default_factory=dict)
    type_coercion_failures: dict = field(default_factory=dict)
    outliers_flagged: int = 0
    revenue_recomputed: int = 0

    def summary_lines(self) -> list[str]:
        lines = [
            f"Rows ingested: {self.rows_in}",
            f"Rows in final clean dataset: {self.rows_out}",
            f"Duplicate rows removed: {self.duplicates_removed}",
            f"Outlier rows flagged for review: {self.outliers_flagged}",
            f"Revenue rows recomputed from quantity*unit_price: {self.revenue_recomputed}",
        ]
        for col, n in self.missing_values_filled.items():
            lines.append(f"  - '{col}': {n} missing values filled/imputed")
        for col, n in self.type_coercion_failures.items():
            lines.append(f"  - '{col}': {n} values could not be parsed and were set to NaN")
        return lines


class DataPipeline:
    def __init__(self, input_path: str):
        self.input_path = Path(input_path)
        self.report = DataQualityReport()
        self.df_raw: pd.DataFrame | None = None
        self.df_clean: pd.DataFrame | None = None

    # ------------------------------------------------------------- ingest
    def ingest(self) -> pd.DataFrame:
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        if self.input_path.suffix.lower() == ".csv":
            df = pd.read_csv(self.input_path)
        else:
            df = pd.read_excel(self.input_path)
        if df.empty:
            raise ValueError("Input file contains no rows.")
        self.df_raw = df
        self.report.rows_in = len(df)
        return df

    # -------------------------------------------------------------- clean
    def clean(self) -> pd.DataFrame:
        df = self.df_raw.copy()

        # 1. Normalize column names
        df.columns = [re.sub(r"\s+", "_", c.strip().lower()) for c in df.columns]

        missing_required = REQUIRED_COLUMNS - set(df.columns)
        if missing_required:
            raise ValueError(
                f"Input is missing required column(s): {sorted(missing_required)}. "
                f"Found columns: {list(df.columns)}"
            )

        # 2. Strip whitespace from all string/object columns
        for col in df.select_dtypes(include=["object", "string"]).columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": np.nan, "": np.nan, "None": np.nan})

        # 3. Coerce numeric columns (strips currency symbols/commas first)
        for col in NUMERIC_COLUMNS:
            if col not in df.columns:
                continue
            before_na = df[col].isna().sum()
            cleaned = (
                df[col].astype(str)
                .str.replace(r"[$,€£]", "", regex=True)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(cleaned, errors="coerce")
            failures = df[col].isna().sum() - before_na
            if failures > 0:
                self.report.type_coercion_failures[col] = int(failures)

        # 4. Parse dates
        for col in DATE_COLUMNS:
            if col not in df.columns:
                continue
            df[col] = pd.to_datetime(df[col], errors="coerce", format="mixed")

        # 5. Recompute revenue BEFORE the generic fill step, so we can tell
        #    "recomputed from clean inputs" apart from "filled with a median
        #    because everything was missing". This ordering fix is the main
        #    correctness bug from the original script.
        recomputed_mask = pd.Series(False, index=df.index)
        if {"quantity", "unit_price"}.issubset(df.columns):
            have_inputs = df["quantity"].notna() & df["unit_price"].notna()
            expected_revenue = df["quantity"] * df["unit_price"]
            if "revenue" in df.columns:
                mismatch = have_inputs & ((df["revenue"] - expected_revenue).abs() > 0.01)
                missing_rev = have_inputs & df["revenue"].isna()
                recomputed_mask = mismatch | missing_rev
                df.loc[recomputed_mask, "revenue"] = expected_revenue[recomputed_mask]
            else:
                df["revenue"] = expected_revenue
                recomputed_mask = have_inputs
        self.report.revenue_recomputed = int(recomputed_mask.sum())

        # 6. Fill / flag remaining missing values (strategy differs by column type)
        for col in df.columns:
            n_missing = int(df[col].isna().sum())
            if n_missing == 0:
                continue
            if col in NUMERIC_COLUMNS:
                fill_value = df[col].median()
                df[col] = df[col].fillna(fill_value)
                self.report.missing_values_filled[col] = n_missing
            elif col not in DATE_COLUMNS:
                df[col] = df[col].fillna("UNKNOWN")
                self.report.missing_values_filled[col] = n_missing

        # 7. Remove exact duplicate rows
        before = len(df)
        df = df.drop_duplicates()
        self.report.duplicates_removed = before - len(df)

        # 8. Flag statistical outliers in revenue (IQR method) — not removed,
        #    just flagged, since a human should decide if a $50k sale is real.
        if REVENUE_COLUMN in df.columns:
            q1, q3 = df[REVENUE_COLUMN].quantile([0.25, 0.75])
            iqr = q3 - q1
            lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            df["flag_outlier"] = ~df[REVENUE_COLUMN].between(lower, upper)
            self.report.outliers_flagged = int(df["flag_outlier"].sum())

        df = df.reset_index(drop=True)
        self.report.rows_out = len(df)
        self.df_clean = df
        return df

    # ------------------------------------------------------------ analyze
    def analyze(self) -> dict:
        df = self.df_clean
        summary: dict = {}
        if REVENUE_COLUMN in df.columns:
            summary["total_revenue"] = round(float(df[REVENUE_COLUMN].sum()), 2)
            summary["avg_order_value"] = round(float(df[REVENUE_COLUMN].mean()), 2)
            summary["order_count"] = int(len(df))

        if KEY_GROUP_COLUMN in df.columns and REVENUE_COLUMN in df.columns:
            by_group = (
                df.groupby(KEY_GROUP_COLUMN)[REVENUE_COLUMN]
                .sum().sort_values(ascending=False).round(2)
            )
            total = by_group.sum()
            pct_of_total = (by_group / total * 100).round(1) if total else by_group * 0
            summary["revenue_by_group"] = by_group
            summary["revenue_by_group_pct"] = pct_of_total

        if "order_date" in df.columns:
            monthly = (
                df.dropna(subset=["order_date"])
                .set_index("order_date")[REVENUE_COLUMN]
                .resample(MONTH_END_ALIAS).sum().round(2)
            )
            mom_pct = monthly.pct_change().mul(100).round(1)
            summary["monthly_revenue"] = monthly
            summary["monthly_mom_pct"] = mom_pct

        return summary

    # ------------------------------------------------------------- export
    def export_excel(self, summary: dict, out_path: str):
        wb = Workbook()
        header_font = Font(bold=True, size=14, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="2F5496")
        subheader_font = Font(bold=True)
        currency_fmt = '"$"#,##0.00'
        pct_fmt = "0.0%"

        # ---- Sheet 1: Executive Summary ----
        ws = wb.active
        ws.title = "Summary"
        ws["A1"] = "Executive Data Summary"
        ws["A1"].font = header_font
        ws["A1"].fill = header_fill
        ws.merge_cells("A1:D1")

        row = 3
        ws[f"A{row}"] = "Metric"
        ws[f"B{row}"] = "Value"
        for c in ("A", "B"):
            ws[f"{c}{row}"].font = subheader_font
        row += 1
        metric_labels = {
            "total_revenue": ("Total Revenue", currency_fmt),
            "avg_order_value": ("Average Order Value", currency_fmt),
            "order_count": ("Total Orders", None),
        }
        for key, (label, fmt) in metric_labels.items():
            if key in summary:
                ws[f"A{row}"] = label
                cell = ws[f"B{row}"]
                cell.value = summary[key]
                if fmt:
                    cell.number_format = fmt
                row += 1

        if "revenue_by_group" in summary:
            row += 1
            ws[f"A{row}"] = f"Revenue by {KEY_GROUP_COLUMN.title()}"
            ws[f"A{row}"].font = subheader_font
            row += 1
            ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = KEY_GROUP_COLUMN.title(), "Revenue", "% of Total"
            for c in ("A", "B", "C"):
                ws[f"{c}{row}"].font = subheader_font
            row += 1
            chart_start_row = row
            pct = summary.get("revenue_by_group_pct")
            for group_name, val in summary["revenue_by_group"].items():
                ws[f"A{row}"] = group_name
                ws[f"B{row}"] = val
                ws[f"B{row}"].number_format = currency_fmt
                if pct is not None:
                    ws[f"C{row}"] = float(pct[group_name]) / 100
                    ws[f"C{row}"].number_format = pct_fmt
                row += 1
            chart_end_row = row - 1

            chart = BarChart()
            chart.title = f"Revenue by {KEY_GROUP_COLUMN.title()}"
            chart.y_axis.title = "Revenue"
            data = Reference(ws, min_col=2, min_row=chart_start_row, max_row=chart_end_row)
            cats = Reference(ws, min_col=1, min_row=chart_start_row, max_row=chart_end_row)
            chart.add_data(data, titles_from_data=False)
            chart.set_categories(cats)
            ws.add_chart(chart, f"E{chart_start_row}")
            row += 2

        if "monthly_revenue" in summary:
            ws[f"A{row}"] = "Monthly Revenue Trend"
            ws[f"A{row}"].font = subheader_font
            row += 1
            ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = "Month", "Revenue", "MoM % Change"
            for c in ("A", "B", "C"):
                ws[f"{c}{row}"].font = subheader_font
            row += 1
            trend_start = row
            mom = summary.get("monthly_mom_pct")
            for month, val in summary["monthly_revenue"].items():
                ws[f"A{row}"] = month.strftime("%Y-%m")
                ws[f"B{row}"] = val
                ws[f"B{row}"].number_format = currency_fmt
                mchange = mom[month] if mom is not None else None
                if mchange is not None and not pd.isna(mchange):
                    ws[f"C{row}"] = float(mchange) / 100
                    ws[f"C{row}"].number_format = pct_fmt
                row += 1
            trend_end = row - 1

            line = LineChart()
            line.title = "Monthly Revenue Trend"
            data = Reference(ws, min_col=2, min_row=trend_start, max_row=trend_end)
            cats = Reference(ws, min_col=1, min_row=trend_start, max_row=trend_end)
            line.add_data(data, titles_from_data=False)
            line.set_categories(cats)
            ws.add_chart(line, f"E{trend_start}")

        for col, width in zip("ABCDE", (28, 16, 14, 4, 40)):
            ws.column_dimensions[col].width = width

        # ---- Sheet 2: Data Quality Report ----
        ws2 = wb.create_sheet("Data Quality Report")
        ws2["A1"] = "Data Quality Report"
        ws2["A1"].font = header_font
        ws2["A1"].fill = header_fill
        ws2.merge_cells("A1:B1")
        for i, line in enumerate(self.report.summary_lines(), start=3):
            ws2[f"A{i}"] = line
        ws2.column_dimensions["A"].width = 75

        # ---- Sheet 3: Cleaned Data ----
        ws3 = wb.create_sheet("Cleaned Data")
        for r in dataframe_to_rows(self.df_clean, index=False, header=True):
            ws3.append(r)
        for cell in ws3[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2F5496")
        for i, col in enumerate(self.df_clean.columns, start=1):
            ws3.column_dimensions[get_column_letter(i)].width = 16
        ws3.freeze_panes = "A2"

        wb.save(out_path)
        logger.info("Executive report written to: %s", out_path)

    # ---------------------------------------------------------------- run
    def run(self, out_path: str):
        logger.info("Ingesting raw data from %s", self.input_path)
        self.ingest()
        logger.info("%d rows loaded", self.report.rows_in)

        logger.info("Cleaning data...")
        self.clean()
        logger.info(
            "%d clean rows (%d duplicates removed, %d revenue values recomputed)",
            self.report.rows_out, self.report.duplicates_removed, self.report.revenue_recomputed,
        )

        logger.info("Analyzing...")
        summary = self.analyze()

        logger.info("Exporting formatted Excel workbook...")
        self.export_excel(summary, out_path)

        logger.info("DATA QUALITY REPORT")
        for line in self.report.summary_lines():
            logger.info("  %s", line)


def main():
    parser = argparse.ArgumentParser(description="Clean a messy sales export and produce an executive Excel report.")
    parser.add_argument("--input", "-i", default=DEFAULT_INPUT, help="Path to input CSV/XLSX file.")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help="Path to write the output .xlsx report.")
    args = parser.parse_args()

    try:
        DataPipeline(args.input).run(args.output)
    except (FileNotFoundError, ValueError) as e:
        logger.error(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
