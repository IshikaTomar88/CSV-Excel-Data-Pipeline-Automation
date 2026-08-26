"""
FILE-BASED ENTERPRISE DATA PIPELINE & REPORTING SUITE
Production-Grade Streamlit Application for Automated CSV/Excel Transformation
"""

import io
from dataclasses import dataclass, field
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --------------------------------------------------------------------------
# PAGE CONFIGURATION & ENTERPRISE STYLING
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Data Transformation & Executive Reporter",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main { background-color: #0f172a; color: #f8fafc; }
        .stMetric { background-color: #1e293b; padding: 15px; border-radius: 10px; border: 1px solid #334155; }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; }
        .stTabs [data-baseweb="tab"] { background-color: #1e293b; border-radius: 6px; color: #cbd5e1; padding: 10px 20px; font-weight: 600; }
        .stTabs [aria-selected="true"] { background-color: #2563eb !important; color: white !important; }
        div.stButton > button { background-color: #2563eb; color: white; font-weight: bold; border-radius: 6px; padding: 0.5rem 1rem; border: none; width: 100%; }
        div.stButton > button:hover { background-color: #1d4ed8; }
    </style>
    """,
    unsafe_allow_html=True,
)

COLUMN_ALIASES = {
    "revenue": ["revenue", "total_sales", "sales", "amount", "total_amount", "gmv"],
    "order_date": ["order_date", "date", "transaction_date", "invoice_date", "timestamp"],
    "quantity": ["quantity", "qty", "units", "volume"],
    "unit_price": ["unit_price", "price", "rate", "cost_per_unit"],
    "region": ["region", "market", "territory", "location", "zone", "branch"],
}

@dataclass
class DataQualityReport:
    rows_in: int = 0
    rows_out: int = 0
    duplicates_removed: int = 0
    missing_values_filled: dict[str, int] = field(default_factory=dict)
    type_coercion_failures: dict[str, int] = field(default_factory=dict)
    outliers_flagged: int = 0
    revenue_recomputed: int = 0

    def summary_lines(self) -> list[str]:
        lines = [
            f"Total rows ingested: {self.rows_in}",
            f"Total rows in final clean dataset: {self.rows_out}",
            f"Exact duplicate rows dropped: {self.duplicates_removed}",
            f"Statistical outliers flagged for review (using IQR method): {self.outliers_flagged}",
            f"Revenue values recalculated from quantity * unit_price: {self.revenue_recomputed}",
        ]
        for col, n in self.missing_values_filled.items():
            lines.append(f"  - Column '{col}': {n} missing values imputed with column median / fallback")
        for col, n in self.type_coercion_failures.items():
            lines.append(f"  - Column '{col}': {n} malformed values coerced safely to NaN")
        return lines


class EnterprisePipeline:
    def __init__(self, uploaded_file):
        self.uploaded_file = uploaded_file
        self.report = DataQualityReport()
        self.df_raw: pd.DataFrame | None = None
        self.df_clean: pd.DataFrame | None = None

    def ingest(self) -> pd.DataFrame:
        filename = self.uploaded_file.name.lower()
        if filename.endswith(".csv"):
            df = pd.read_csv(self.uploaded_file)
        elif filename.endswith((".xls", ".xlsx")):
            df = pd.read_excel(self.uploaded_file)
        else:
            raise ValueError("Unsupported file format. Please upload a CSV or Excel file.")

        if df.empty:
            raise ValueError("The uploaded file contains no rows.")
        
        self.df_raw = df
        self.report.rows_in = len(df)
        return df

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        import re
        df.columns = [re.sub(r"\s+", "_", c.strip().lower()) for c in df.columns]
        rename_map = {}
        for standard_name, aliases in COLUMN_ALIASES.items():
            if standard_name not in df.columns:
                for col in df.columns:
                    if col in aliases:
                        rename_map[col] = standard_name
                        break
        if rename_map:
            df = df.rename(columns=rename_map)
        return df

    def clean(self) -> pd.DataFrame:
        df = self.df_raw.copy()
        df = self._normalize_columns(df)

        if "revenue" not in df.columns:
            if {"quantity", "unit_price"}.issubset(df.columns):
                df["revenue"] = pd.to_numeric(df["quantity"], errors="coerce") * pd.to_numeric(df["unit_price"], errors="coerce")
            else:
                raise ValueError("Dataset lacks a 'revenue' column and cannot derive it without quantity/unit_price.")

        for col in df.select_dtypes(include=["object", "string"]).columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": np.nan, "": np.nan, "None": np.nan})

        numeric_cols = ["quantity", "unit_price", "revenue"]
        for col in numeric_cols:
            if col in df.columns:
                before_na = df[col].isna().sum()
                cleaned = df[col].astype(str).str.replace(r"[$,€£]", "", regex=True).str.replace(",", "", regex=False).str.strip()
                df[col] = pd.to_numeric(cleaned, errors="coerce")
                failures = df[col].isna().sum() - before_na
                if failures > 0:
                    self.report.type_coercion_failures[col] = int(failures)

        if "order_date" in df.columns:
            df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")

        recomputed_mask = pd.Series(False, index=df.index)
        if {"quantity", "unit_price"}.issubset(df.columns):
            have_inputs = df["quantity"].notna() & df["unit_price"].notna()
            expected_revenue = df["quantity"] * df["unit_price"]
            mismatch = have_inputs & ((df["revenue"] - expected_revenue).abs() > 0.01)
            missing_rev = have_inputs & df["revenue"].isna()
            recomputed_mask = mismatch | missing_rev
            df.loc[recomputed_mask, "revenue"] = expected_revenue[recomputed_mask]
        self.report.revenue_recomputed = int(recomputed_mask.sum())

        for col in df.columns:
            n_missing = int(df[col].isna().sum())
            if n_missing == 0:
                continue
            if col in numeric_cols:
                df[col] = df[col].fillna(df[col].median() if not df[col].isna().all() else 0.0)
                self.report.missing_values_filled[col] = n_missing
            elif col != "order_date":
                df[col] = df[col].fillna("UNKNOWN")
                self.report.missing_values_filled[col] = n_missing

        before = len(df)
        df = df.drop_duplicates()
        self.report.duplicates_removed = before - len(df)

        if "revenue" in df.columns and len(df) > 5:
            q1, q3 = df["revenue"].quantile([0.25, 0.75])
            iqr = q3 - q1
            lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            df["flag_outlier"] = ~df["revenue"].between(lower, upper)
            self.report.outliers_flagged = int(df["flag_outlier"].sum())
        else:
            df["flag_outlier"] = False

        self.df_clean = df.reset_index(drop=True)
        self.report.rows_out = len(self.df_clean)
        return self.df_clean

    def analyze(self) -> dict:
        df = self.df_clean
        summary = {}
        if "revenue" in df.columns:
            summary["total_revenue"] = round(float(df["revenue"].sum()), 2)
            summary["avg_order_value"] = round(float(df["revenue"].mean()), 2)
            summary["order_count"] = int(len(df))

        group_col = "region" if "region" in df.columns else (df.select_dtypes(include=["object", "string"]).columns[0] if len(df.select_dtypes(include=["object", "string"]).columns) > 0 else None)
        
        if group_col and "revenue" in df.columns:
            by_group = df.groupby(group_col)["revenue"].sum().sort_values(ascending=False).round(2)
            total = by_group.sum()
            summary["group_column_name"] = group_col
            summary["revenue_by_group"] = by_group
            summary["revenue_by_group_pct"] = (by_group / total * 100).round(1) if total else by_group * 0

        if "order_date" in df.columns and df["order_date"].notna().sum() > 0:
            monthly = df.dropna(subset=["order_date"]).set_index("order_date")["revenue"].resample("ME").sum().round(2)
            summary["monthly_revenue"] = monthly
            summary["monthly_mom_pct"] = monthly.pct_change().mul(100).round(1)

        return summary

    def build_excel_bytes(self, summary: dict) -> bytes:
        wb = Workbook()
        header_font = Font(bold=True, size=14, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="2F5496")
        subheader_font = Font(bold=True)
        currency_fmt = '"$"#,##0.00'
        pct_fmt = "0.0%"

        ws = wb.active
        ws.title = "Executive Summary"
        ws["A1"] = "Executive Data Summary & Performance"
        ws["A1"].font = header_font
        ws["A1"].fill = header_fill
        ws.merge_cells("A1:D1")

        row = 3
        ws[f"A{row}"] = "Metric"
        ws[f"B{row}"] = "Value"
        for c in ("A", "B"): ws[f"{c}{row}"].font = subheader_font
        row += 1

        metrics = [
            ("total_revenue", "Total Revenue", currency_fmt),
            ("avg_order_value", "Average Order Value", currency_fmt),
            ("order_count", "Total Clean Orders", None)
        ]
        for key, label, fmt in metrics:
            if key in summary:
                ws[f"A{row}"] = label
                cell = ws[f"B{row}"]
                cell.value = summary[key]
                if fmt: cell.number_format = fmt
                row += 1

        group_col = summary.get("group_column_name", "Group").title()
        if "revenue_by_group" in summary:
            row += 2
            ws[f"A{row}"] = f"Revenue Breakdown by {group_col}"
            ws[f"A{row}"].font = subheader_font
            row += 1
            ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = group_col, "Revenue", "% of Total"
            for c in ("A", "B", "C"): ws[f"{c}{row}"].font = subheader_font
            row += 1
            chart_start = row
            pct = summary.get("revenue_by_group_pct")
            for name, val in summary["revenue_by_group"].items():
                ws[f"A{row}"] = str(name)
                ws[f"B{row}"] = val
                ws[f"B{row}"].number_format = currency_fmt
                if pct is not None and name in pct:
                    ws[f"C{row}"] = float(pct[name]) / 100
                    ws[f"C{row}"].number_format = pct_fmt
                row += 1
            chart_end = row - 1

            chart = BarChart()
            chart.title = f"Revenue by {group_col}"
            chart.y_axis.title = "Revenue ($)"
            chart.add_data(Reference(ws, min_col=2, min_row=chart_start, max_row=chart_end), titles_from_data=False)
            chart.set_categories(Reference(ws, min_col=1, min_row=chart_start, max_row=chart_end))
            ws.add_chart(chart, f"E{chart_start}")
            row += 2

        if "monthly_revenue" in summary and not summary["monthly_revenue"].empty:
            ws[f"A{row}"] = "Monthly Revenue Trend"
            ws[f"A{row}"].font = subheader_font
            row += 1
            ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = "Month", "Revenue", "MoM % Change"
            for c in ("A", "B", "C"): ws[f"{c}{row}"].font = subheader_font
            row += 1
            trend_start = row
            mom = summary.get("monthly_mom_pct")
            for month, val in summary["monthly_revenue"].items():
                ws[f"A{row}"] = month.strftime("%Y-%m")
                ws[f"B{row}"] = val
                ws[f"B{row}"].number_format = currency_fmt
                mchange = mom.get(month) if mom is not None else None
                if mchange is not None and not pd.isna(mchange):
                    ws[f"C{row}"] = float(mchange) / 100
                    ws[f"C{row}"].number_format = pct_fmt
                row += 1
            trend_end = row - 1

            line = LineChart()
            line.title = "Monthly Revenue Trend"
            line.add_data(Reference(ws, min_col=2, min_row=trend_start, max_row=trend_end), titles_from_data=False)
            line.set_categories(Reference(ws, min_col=1, min_row=trend_start, max_row=trend_end))
            ws.add_chart(line, f"E{trend_start}")

        for col, width in zip("ABCDE", (28, 16, 14, 4, 45)):
            ws.column_dimensions[col].width = width

        ws2 = wb.create_sheet("Data Quality Report")
        ws2["A1"] = "Pipeline Data Quality & Health Audit"
        ws2["A1"].font = header_font
        ws2["A1"].fill = header_fill
        ws2.merge_cells("A1:B1")
        for i, line in enumerate(self.report.summary_lines(), start=3):
            ws2[f"A{i}"] = line
        ws2.column_dimensions["A"].width = 75

        ws3 = wb.create_sheet("Cleaned Raw Data")
        for r in dataframe_to_rows(self.df_clean, index=False, header=True):
            ws3.append(r)
        for cell in ws3[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2F5496")
        for i, _ in enumerate(self.df_clean.columns, start=1):
            ws3.column_dimensions[get_column_letter(i)].width = 16
        ws3.freeze_panes = "A2"

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()


# STREAMLIT UI
st.sidebar.title("⚡ Pipeline Studio")
st.sidebar.markdown("---")
st.sidebar.info(
    "**Description:** An offline file transformation utility built to clean "
    "messy CSV/Excel client datasets and export formatted executive reports."
)
st.sidebar.markdown("### Developer")
st.sidebar.write("**Ishika Tomar**")

st.title("📊 File-Based Data Transformation & Executive Reporting Pipeline")
st.markdown("Upload raw CSV or Excel client exports to normalize schemas, clean missing values, flag statistical outliers, and export formatted workbooks.")

uploaded_file = st.file_uploader("Upload Raw Client Dataset (CSV or Excel)", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        with st.spinner("Executing data normalization and validation pipeline..."):
            pipeline = EnterprisePipeline(uploaded_file)
            pipeline.ingest()
            df_clean = pipeline.clean()
            summary = pipeline.analyze()
            excel_bytes = pipeline.build_excel_bytes(summary)

        st.success("✅ Pipeline Executed Successfully.")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Revenue", f"${summary.get('total_revenue', 0):,.2f}")
        with col2:
            st.metric("Average Order Value", f"${summary.get('avg_order_value', 0):,.2f}")
        with col3:
            st.metric("Total Clean Orders", f"{summary.get('order_count', 0):,}")
        with col4:
            st.metric("Outliers Flagged", f"{pipeline.report.outliers_flagged}")

        st.markdown("---")

        tab1, tab2, tab3 = st.tabs(["📊 Analytics Breakdown", "🧹 Data Quality Audit Log", "🔍 Cleaned Dataset Preview"])

        with tab1:
            st.subheader("Executive Performance Overview")
            if "revenue_by_group" in summary:
                col_a, col_b = st.columns(2)
                with col_a:
                    st.write(f"**Revenue Breakdown by {summary.get('group_column_name', 'Group').title()}**")
                    st.dataframe(summary["revenue_by_group"], use_container_width=True)
                with col_b:
                    if "monthly_revenue" in summary and not summary["monthly_revenue"].empty:
                        st.write("**Monthly Revenue Trend**")
                        st.line_chart(summary["monthly_revenue"])
            else:
                st.info("No grouping column detected for regional aggregation.")

        with tab2:
            st.subheader("Data Health & Integrity Audit Log")
            st.markdown("This log explicitly lists how the pipeline processed raw anomalies:")
            for line in pipeline.report.summary_lines():
                st.text(line)

        with tab3:
            st.subheader("Cleaned Dataset")
            st.dataframe(df_clean, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Export Workbook")
        st.download_button(
            label="Download Formatted Executive Report (.xlsx)",
            data=excel_bytes,
            file_name="executive_client_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Pipeline Error: {str(e)}")
else:
    st.markdown(
        """
        > **Instructions:** Upload a standard tabular CSV or Excel file. The script handles column alias mapping, data type coercion, missing data imputation, and outlier flagging automatically.
        """
    )
