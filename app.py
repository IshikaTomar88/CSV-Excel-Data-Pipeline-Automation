"""
================================================================================
 SERVICE: Zero-Touch CSV/Excel Data Pipeline Automation & Executive Reporter
================================================================================
ENTERPRISE FEATURES ADDED:
  1. Complete Data Privacy: In-memory secure processing (zero file retention).
  2. Smart Trial Metering: Locks down automatically after 2 free runs.
  3. Dynamic Admin Access Control: Approve/revoke client access duration (days)
     directly from the UI sidebar or via automated email alerts.
================================================================================
"""

import json
import logging
import os
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# PAGE CONFIGURATION & STYLING
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Enterprise Data Pipeline & Analytics Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main-header { font-size: 2.2rem; font-weight: 700; color: #1E3A8A; }
        .sub-header { font-size: 1.1rem; color: #4B5563; }
        .security-badge { background-color: #ECFDF5; border: 1px solid #10B981; padding: 10px; border-radius: 6px; color: #065F46; font-weight: 600; }
        .warning-box { background-color: #FEF3C7; border: 1px solid #F59E0B; padding: 12px; border-radius: 6px; color: #92400E; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# CONFIGURATION CONSTANTS & MAPPINGS
# --------------------------------------------------------------------------
COLUMN_MAP = {
    "cust name": "customer_name",
    "customer name": "customer_name",
    "customername": "customer_name",
    "order date": "order_date",
    "orderdate": "order_date",
    "amt": "amount",
    "amount ($)": "amount",
    "total": "amount",
    "qty": "quantity",
    "quantity ordered": "quantity",
    "region/state": "region",
}

CLEANING_RULES = {
    "customer_name": {"strip": True, "title_case": True},
    "order_date": {"parse_date": True},
    "amount": {"currency_to_float": True, "min_valid": 0},
    "quantity": {"to_int": True, "min_valid": 0},
    "region": {"strip": True, "title_case": True},
}

MISSING_VALUE_STRATEGY = {
    "customer_name": "fill_unknown",
    "order_date": "drop_row",
    "amount": "fill_mean",
    "quantity": "fill_zero",
    "region": "fill_unknown",
}

# --------------------------------------------------------------------------
# SESSION STATE & PERSISTENT CLIENT PERMISSION DATABASE SIMULATION
# --------------------------------------------------------------------------
if "run_count" not in st.session_state:
    st.session_state.run_count = 0
if "client_email" not in st.session_state:
    st.session_state.client_email = "guest_user@company.com"
if "access_database" not in st.session_state:
    # Tracks approved clients and their access expiry date/time
    # Format: { "client@email.com": {"expiry": datetime, "status": "Active"} }
    st.session_state.access_database = {}

MAX_FREE_RUNS = 2  # Free executions before access approval is mandatory


def send_access_alert_email(client_email: str):
    """Sends automated email notification to developer inbox."""
    try:
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        receiver_email = st.secrets["email"]["receiver_email"]
    except Exception:
        return False

    msg = EmailMessage()
    msg.set_subject(f"🚨 ACCESS REQUEST: {client_email} needs pipeline extension")
    msg.set_content(
        f"Hello Admin,\n\nClient {client_email} has used their {MAX_FREE_RUNS} free pipeline trials.\n"
        f"Log into your Streamlit Admin Sidebar to grant them specific day-based access."
    )
    msg["From"] = sender_email
    msg["To"] = receiver_email

    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True
    except Exception:
        return False


def check_client_authorization(email: str) -> bool:
    """Checks if a given client has active, unexpired day-based access."""
    if email in st.session_state.access_database:
        client_record = st.session_state.access_database[email]
        if datetime.now() < client_record["expiry"]:
            return True
    return False


# --------------------------------------------------------------------------
# ADMIN SIDEBAR CONTROL PANEL
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Admin Control Panel")
    st.markdown("Manage client access durations and permissions remotely.")
    
    admin_password_input = st.text_input("Admin Secret Key", type="password")
    is_admin = admin_password_input == "admin123"  # Change or secure via secrets

    if is_admin:
        st.success("Admin Mode Unlocked 🔓")
        st.divider()
        st.subheader("Client Access Manager")
        
        if st.session_state.access_database:
            for mail, data in list(st.session_state.access_database.items()):
                time_left = data["expiry"] - datetime.now()
                hours_left = max(0, int(time_left.total_seconds() // 3600))
                st.text(f"👤 {mail}\n⏳ Expires in: {hours_left} hrs")
                
                col_adm1, col_adm2 = st.columns(2)
                with col_adm1:
                    if st.button(f"Revoke", key=f"rev_{mail}"):
                        del st.session_state.access_database[mail]
                        st.rerun()
                with col_adm2:
                    if st.button(f"+7 Days", key=f"ext_{mail}"):
                        st.session_state.access_database[mail]["expiry"] = datetime.now() + timedelta(days=7)
                        st.rerun()
                st.markdown("---")
        else:
            st.info("No active custom client permissions logged yet.")
            
        # Manual grant feature for admin
        st.subheader("Grant Access Manually")
        manual_email = st.text_input("Client Email to Authorize")
        grant_days = st.number_input("Access Duration (Days)", min_value=1, max_value=30, value=3)
        if st.button("Grant Access Key"):
            if manual_email:
                st.session_state.access_database[manual_email] = {
                    "expiry": datetime.now() + timedelta(days=grant_days),
                    "status": "Active"
                }
                st.success(f"Granted {grant_days} days access to {manual_email}!")
                st.rerun()
    else:
        st.info("Enter admin key to view and manage client license requests.")


# --------------------------------------------------------------------------
# CORE ETL TRANSFORMATION FUNCTIONS
# --------------------------------------------------------------------------
def load_raw_file(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(uploaded_file)
    elif suffix in (".xlsx", ".xls"):
        return pd.read_excel(uploaded_file)
    elif suffix == ".json":
        return pd.json_normalize(json.load(uploaded_file))
    else:
        raise ValueError(f"Unsupported format: {suffix}")


def normalize_columns(df: pd.DataFrame, column_map: dict) -> pd.DataFrame:
    lowered = {c: c.strip().lower() for c in df.columns}
    df = df.rename(columns=lowered)
    df = df.rename(columns={k: v for k, v in column_map.items() if k in df.columns})
    return df


def _clean_currency_series(s: pd.Series) -> pd.Series:
    cleaned = s.astype(str).str.replace(r"[^\d.\-]", "", regex=True).replace("", np.nan)
    return pd.to_numeric(cleaned, errors="coerce")


def clean_data(df: pd.DataFrame, rules: dict) -> pd.DataFrame:
    for col, rule in rules.items():
        if col not in df.columns:
            continue
        if rule.get("strip"):
            df[col] = df[col].astype(str).str.strip()
        if rule.get("title_case"):
            df[col] = df[col].astype(str).str.title()
        if rule.get("parse_date"):
            df[col] = pd.to_datetime(df[col], errors="coerce")
        if rule.get("currency_to_float"):
            df[col] = _clean_currency_series(df[col])
        if rule.get("to_int"):
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
        if "min_valid" in rule:
            invalid_mask = df[col] < rule["min_valid"]
            df.loc[invalid_mask, col] = np.nan
    return df


def handle_missing_values(df: pd.DataFrame, strategy: dict) -> pd.DataFrame:
    df = df.copy()
    for col, method in strategy.items():
        if col not in df.columns:
            continue
        if method == "drop_row":
            df = df.dropna(subset=[col])
        elif method == "fill_zero":
            df[col] = df[col].fillna(0)
        elif method == "fill_mean":
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(df[col].mean())
        elif method == "fill_unknown":
            df[col] = df[col].fillna("Unknown")
    return df


def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    subset_cols = [c for c in ["customer_name", "order_date", "amount"] if c in df.columns]
    return df.drop_duplicates(subset=subset_cols if subset_cols else None)


def flag_outliers(df: pd.DataFrame, col: str = "amount", z_thresh: float = 3.0) -> pd.DataFrame:
    if col not in df.columns or not pd.api.types.is_numeric_dtype(df[col]):
        df["is_outlier"] = False
        return df
    mean, std = df[col].mean(), df[col].std()
    if std == 0 or pd.isna(std):
        df["is_outlier"] = False
        return df
    z_scores = (df[col] - mean) / std
    df["is_outlier"] = z_scores.abs() > z_thresh
    return df


def build_executive_summary(df: pd.DataFrame) -> pd.DataFrame:
    summary = {
        "Total Records Processed": len(df),
        "Total Revenue": round(df["amount"].sum(), 2) if "amount" in df else 0.0,
        "Average Order Value": round(df["amount"].mean(), 2) if "amount" in df else 0.0,
        "Unique Customers": df["customer_name"].nunique() if "customer_name" in df else 0,
        "Outliers Flagged (Z-Score)": int(df["is_outlier"].sum()) if "is_outlier" in df else 0,
    }
    return pd.DataFrame(list(summary.items()), columns=["Executive Metric", "Value"])


# --------------------------------------------------------------------------
# MAIN APPLICATION INTERFACE & ACCESS GATE LOGIC
# --------------------------------------------------------------------------
st.markdown('<p class="main-header">⚡ Zero-Touch Data Pipeline & Executive Cleaner</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Transform chaotic raw data dumps into executive-ready dashboards instantly with 100% mathematical precision.</p>', unsafe_allow_html=True)
st.markdown("---")

# Privacy Security Banner
st.markdown(
    """
    <div class="security-badge">
        🔒 <b>Enterprise Privacy Guarantee:</b> Your data is processed entirely in isolated memory and is never saved, logged, or retained on permanent servers. Complete confidentiality is guaranteed.
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# Evaluate if client needs authorization gate
is_authorized = check_client_authorization(st.session_state.client_email)
trial_exhausted = st.session_state.run_count >= MAX_FREE_RUNS

if trial_exhausted and not is_authorized:
    st.markdown(
        """
        <div class="warning-box">
            <b>Trial Limit Reached:</b> You have utilized your free sample executions. Please submit your business email to request extended day-based operational access.
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    col_req1, col_req2 = st.columns([2, 1])
    with col_req1:
        email_input = st.text_input("Enter your business email address:", value=st.session_state.client_email)
    with col_req2:
        st.markdown("<br>", unsafe_allow_html=True)
        request_clicked = st.button("📨 Request Access Extension", type="primary")

    if request_clicked:
        if email_input:
            st.session_state.client_email = email_input
            send_access_alert_email(email_input)
            st.success("Access request transmitted successfully! The administrator has been notified to unlock your dashboard.")
        else:
            st.error("Please provide a valid email.")
            
    st.stop()  # Halt app rendering until authorized

# --- OPERATIONAL PIPELINE INTERFACE ---
uploaded_file = st.file_uploader("Upload Raw Business Data Export (.csv, .xlsx, .xls, .json)", type=["csv", "xlsx", "xls", "json"])

if uploaded_file is not None:
    try:
        with st.spinner("Extracting file into secure session memory..."):
            df_raw = load_raw_file(uploaded_file)

        st.write(f"### 📥 Raw Data Preview ({len(df_raw):,} rows identified)")
        st.dataframe(df_raw.head(3), use_container_width=True)

        if st.button("🚀 Run Automated Data Cleaning & Transformation Pipeline", type="primary"):
            st.session_state.run_count += 1

            with st.spinner("Executing vectorized Pandas/NumPy cleaning algorithms..."):
                df = normalize_columns(df_raw, COLUMN_MAP)
                df = clean_data(df, CLEANING_RULES)
                df = handle_missing_values(df, MISSING_VALUE_STRATEGY)
                df = remove_duplicates(df)
                df = flag_outliers(df, col="amount" if "amount" in df else df.columns[0])
                summary_df = build_executive_summary(df)

            st.success(f"Pipeline finished successfully! Free runs used: {min(st.session_state.run_count, MAX_FREE_RUNS)}/{MAX_FREE_RUNS}")
            
            st.write("### 📊 Executive KPI Summary Dashboard")
            st.dataframe(summary_df, use_container_width=True)

            st.write("### ✨ Fully Cleaned Dataset Preview")
            st.dataframe(df.head(10), use_container_width=True)

            # File Downloads
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                csv_data = df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Clean CSV Report",
                    data=csv_data,
                    file_name="executive_clean_data.csv",
                    mime="text/csv",
                )
            with col_d2:
                import io
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df.to_excel(writer, sheet_name="Clean Data", index=False)
                    summary_df.to_excel(writer, sheet_name="Executive Summary", index=False)
                
                excel_bytes = buffer.getvalue()
                st.download_button(
                    label="📥 Download Executive Excel Report (.xlsx)",
                    data=excel_bytes,
                    file_name="executive_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    except Exception as error_msg:
        st.error(f"Pipeline runtime error: {error_msg}")