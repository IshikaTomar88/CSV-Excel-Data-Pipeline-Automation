"""
 ZERO-TOUCH DATA PIPELINE — Messy Raw Data -> Executive Excel Report
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
INPUT_FILE = "sample_data/messy_sales_export.csv"
OUTPUT_XLSX = "executive_report.xlsx"
DATE_COLUMNS = ["order_date"]
NUMERIC_COLUMNS = ["quantity", "unit_price", "revenue"]
KEY_GROUP_COLUMN = "region"          # what to summarize revenue by
REVENUE_COLUMN = "revenue"


@dataclass
class DataQualityReport:
    rows_in: int = 0
    rows_out: int = 0
    duplicates_removed: int = 0
    missing_values_filled: dict = field(default_factory=dict)
    type_coercion_failures: dict = field(default_factory=dict)
    outliers_flagged: int = 0

    def summary_lines(self) -> list[str]:
        lines = [
            f"Rows ingested: {self.rows_in}",
            f"Rows in final clean dataset: {self.rows_out}",
            f"Duplicate rows removed: {self.duplicates_removed}",
            f"Outlier rows flagged for review: {self.outliers_flagged}",
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
        if self.input_path.suffix.lower() == ".csv":
            df = pd.read_csv(self.input_path)
        else:
            df = pd.read_excel(self.input_path)
        self.df_raw = df
        self.report.rows_in = len(df)
        return df

    # -------------------------------------------------------------- clean
    def clean(self) -> pd.DataFrame:
        df = self.df_raw.copy()

        # 1. Normalize column names: "Order Date " -> "order_date"
        df.columns = [
            re.sub(r"\s+", "_", c.strip().lower()) for c in df.columns
        ]

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

        # 5. Fill / flag missing values (strategy differs by column type)
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

        # 6. Recompute revenue if missing/inconsistent (quantity * unit_price)
        if {"quantity", "unit_price"}.issubset(df.columns):
            expected_revenue = df["quantity"] * df["unit_price"]
            if "revenue" in df.columns:
                mismatch = (df["revenue"] - expected_revenue).abs() > 0.01
                df.loc[mismatch, "revenue"] = expected_revenue[mismatch]
            else:
                df["revenue"] = expected_revenue

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
        summary = {}
        if REVENUE_COLUMN in df.columns:
            summary["total_revenue"] = round(float(df[REVENUE_COLUMN].sum()), 2)
            summary["avg_order_value"] = round(float(df[REVENUE_COLUMN].mean()), 2)
        if KEY_GROUP_COLUMN in df.columns and REVENUE_COLUMN in df.columns:
            by_group = (
                df.groupby(KEY_GROUP_COLUMN)[REVENUE_COLUMN]
                .sum().sort_values(ascending=False).round(2)
            )
            summary["revenue_by_group"] = by_group
        if "order_date" in df.columns:
            monthly = (
                df.dropna(subset=["order_date"])
                .set_index("order_date")[REVENUE_COLUMN]
                .resample("ME").sum().round(2)
            )
            summary["monthly_revenue"] = monthly
        return summary

    # ------------------------------------------------------------- export
    def export_excel(self, summary: dict, out_path: str):
        wb = Workbook()

        # ---- Sheet 1: Executive Summary ----
        ws = wb.active
        ws.title = "Summary"
        header_font = Font(bold=True, size=14, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="2F5496")
        ws["A1"] = "Executive Data Summary"
        ws["A1"].font = header_font
        ws["A1"].fill = header_fill
        ws.merge_cells("A1:C1")

        row = 3
        ws[f"A{row}"] = "Metric"; ws[f"B{row}"] = "Value"
        for c in ("A", "B"):
            ws[f"{c}{row}"].font = Font(bold=True)
        row += 1
        for key in ("total_revenue", "avg_order_value"):
            if key in summary:
                ws[f"A{row}"] = key.replace("_", " ").title()
                ws[f"B{row}"] = summary[key]
                row += 1

        if "revenue_by_group" in summary:
            row += 1
            ws[f"A{row}"] = f"Revenue by {KEY_GROUP_COLUMN.title()}"
            ws[f"A{row}"].font = Font(bold=True)
            row += 1
            chart_start_row = row
            for group_name, val in summary["revenue_by_group"].items():
                ws[f"A{row}"] = group_name
                ws[f"B{row}"] = val
                row += 1
            chart_end_row = row - 1

            chart = BarChart()
            chart.title = f"Revenue by {KEY_GROUP_COLUMN.title()}"
            chart.y_axis.title = "Revenue"
            data = Reference(ws, min_col=2, min_row=chart_start_row, max_row=chart_end_row)
            cats = Reference(ws, min_col=1, min_row=chart_start_row, max_row=chart_end_row)
            chart.add_data(data, titles_from_data=False)
            chart.set_categories(cats)
            ws.add_chart(chart, f"D{chart_start_row}")

        for col, width in zip("ABCD", (28, 16, 16, 16)):
            ws.column_dimensions[col].width = width

        # ---- Sheet 2: Data Quality Report ----
        ws2 = wb.create_sheet("Data Quality Report")
        ws2["A1"] = "Data Quality Report"
        ws2["A1"].font = header_font
        ws2["A1"].fill = header_fill
        ws2.merge_cells("A1:B1")
        for i, line in enumerate(self.report.summary_lines(), start=3):
            ws2[f"A{i}"] = line
        ws2.column_dimensions["A"].width = 70

        # ---- Sheet 3: Cleaned Data ----
        ws3 = wb.create_sheet("Cleaned Data")
        for r in dataframe_to_rows(self.df_clean, index=False, header=True):
            ws3.append(r)
        for cell in ws3[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2F5496")
        for i, col in enumerate(self.df_clean.columns, start=1):
            ws3.column_dimensions[get_column_letter(i)].width = 16

        wb.save(out_path)
        print(f"\n✅ Executive report written to: {out_path}")

    # ---------------------------------------------------------------- run
    def run(self):
        print("Ingesting raw data...")
        self.ingest()
        print(f"  -> {self.report.rows_in} rows loaded")

        print("Cleaning data...")
        self.clean()
        print(f"  -> {self.report.rows_out} clean rows "
              f"({self.report.duplicates_removed} duplicates removed)")

        print("Analyzing...")
        summary = self.analyze()

        print("Exporting formatted Excel workbook...")
        self.export_excel(summary, OUTPUT_XLSX)

        print("\n" + "=" * 60)
        print("DATA QUALITY REPORT")
        for line in self.report.summary_lines():
            print(" ", line)
        print("=" * 60)


if __name__ == "__main__":
    if not Path(INPUT_FILE).exists():
        print(f"Input file not found: {INPUT_FILE}\n"
              f"Run generate_sample_data.py first to create a demo file.")
        sys.exit(1)
    DataPipeline(INPUT_FILE).run()
