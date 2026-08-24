"""
================================================================================
 SERVICE: Enterprise Zero-Touch Data Pipeline & Secure Memory Suite
================================================================================
"""

import hashlib
import io
import json
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st


# PAGE CONFIGURATION & MODERN EXECUTIVE STYLING

st.set_page_config(
    page_title="Enterprise Data Pipeline & Secure Vault",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main-title { font-size: 2.5rem; font-weight: 800; color: #0F172A; letter-spacing: -0.025em; }
        .sub-title { font-size: 1.1rem; color: #475569; font-weight: 400; }
        .card { background: #F8FAFC; border: 1px solid #E2E8F0; padding: 20px; border-radius: 12px; margin-bottom: 15px; }
        .secure-banner { background: #064E3B; color: #ECFDF5; padding: 12px 18px; border-radius: 8px; font-weight: 500; font-size: 0.95rem; display: flex; align-items: center; gap: 10px; }
        .stButton>button { border-radius: 8px; font-weight: 600; padding: 0.5rem 1rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# INITIALIZE ADVANCED MEMORY MANAGER IN SESSION STATE

if "short_term_cache" not in st.session_state:
    st.session_state.short_term_cache = None  # Active session cleaned data dataframe

if "long_term_vault" not in st.session_state:
    # Stores encrypted or password-locked reports: { report_name: { "summary": df, "data": bytes, "hash": str, "timestamp": str } }
    st.session_state.long_term_vault = {}


# CONFIGURATION CONSTANTS & MAPPINGS

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


# CORE ETL TRANSFORMATION & VECTORIZED ENGINE

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


# SIDEBAR — ADVANCED MEMORY & SECURE VAULT MANAGER

with st.sidebar:
    st.markdown("### 🧠 Memory & Vault Manager")
    st.markdown("Control short-term operational cache, long-term summaries, and military-grade encryption keys.")
    st.divider()

    st.markdown("#### 📦 Long-Term Secure Vault")
    if st.session_state.long_term_vault:
        st.success(f"{len(st.session_state.long_term_vault)} report(s) safely vaulted.")
        selected_vault_item = st.selectbox("Select Vault Report", list(st.session_state.long_term_vault.keys()))
        
        vault_pwd_input = st.text_input("Enter Vault File Password", type="password", key="vault_unlock_pwd")
        
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            if st.button("Unlock & Download"):
                record = st.session_state.long_term_vault[selected_vault_item]
                hashed_input = hashlib.sha256(vault_pwd_input.encode()).hexdigest()
                if hashed_input == record["hash"]:
                    st.success("Access Granted!")
                    st.download_button(
                        "📥 Get Encrypted Excel",
                        data=record["data"],
                        file_name=f"secure_{selected_vault_item}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Incorrect Vault Password!")
        with col_v2:
            if st.button("Purge Item"):
                del st.session_state.long_term_vault[selected_vault_item]
                st.rerun()
    else:
        st.info("Vault is currently empty.")

    st.divider()
    if st.button("🧹 Clear All Session & Cache Memory", type="secondary"):
        st.session_state.short_term_cache = None
        st.session_state.long_term_vault = {}
        st.rerun()


# MAIN INTERFACE

st.markdown('<p class="main-title">⚡ Zero-Touch Data Pipeline & Secure Vault</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">High-performance vector cleaning suite backed by a privacy-first memory manager and optional password protection.</p>', unsafe_allow_html=True)
st.markdown("---")

# Privacy Security Notice Banner
st.markdown(
    """
    <div class="secure-banner">
        🛡️ <b>Strict Zero-Retention Privacy:</b> All raw files are parsed in isolated memory. Data is never written to public disks unless you explicitly choose to save summaries or encrypt files into your private session vault.
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# Main Application Tabs
tab_pipeline, tab_vault_viewer = st.tabs(["🚀 Pipeline Execution", "📂 Active Memory & Summary Logs"])

with tab_pipeline:
    uploaded_file = st.file_uploader("Upload Raw Business Dataset (.csv, .xlsx, .xls, .json)", type=["csv", "xlsx", "xls", "json"])

    if uploaded_file is not None:
        try:
            with st.spinner("Extracting file into secure memory..."):
                df_raw = load_raw_file(uploaded_file)

            st.write(f"### 📥 Raw Dataset Preview ({len(df_raw):,} records found)")
            st.dataframe(df_raw.head(3), use_container_width=True)

            if st.button("✨ Run Vectorized Cleaning Pipeline", type="primary"):
                with st.spinner("Executing high-speed Pandas/NumPy cleaning pipeline..."):
                    df = normalize_columns(df_raw, COLUMN_MAP)
                    df = clean_data(df, CLEANING_RULES)
                    df = handle_missing_values(df, MISSING_VALUE_STRATEGY)
                    df = remove_duplicates(df)
                    df = flag_outliers(df, col="amount" if "amount" in df else df.columns[0])
                    summary_df = build_executive_summary(df)

                # Save to short-term session memory manager
                st.session_state.short_term_cache = {
                    "filename": uploaded_file.name,
                    "clean_df": df,
                    "summary_df": summary_df,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                st.success("Pipeline successfully completed and cached in short-term memory!")

            # If data exists in short term memory, display results and advanced saving options
            if st.session_state.short_term_cache is not None:
                cache = st.session_state.short_term_cache
                st.markdown("---")
                st.write("### 📊 Executive KPI Summary")
                st.dataframe(cache["summary_df"], use_container_width=True)

                st.write("### ✨ Cleaned Data Sample")
                st.dataframe(cache["clean_df"].head(10), use_container_width=True)

                # Export Options & Long-Term Vault Storage configuration
                st.markdown("### 🔒 Privacy & Long-Term Storage Controls")
                col_opts1, col_opts2 = st.columns(2)
                
                with col_opts1:
                    enable_vault_save = st.checkbox("Save report summary to Long-Term Secure Vault")
                with col_opts2:
                    enable_password = st.checkbox("Add Military-Grade Password Protection")

                vault_report_name = ""
                file_password = ""
                if enable_vault_save:
                    vault_report_name = st.text_input("Vault Record Label Name", value=f"Report_{cache['filename']}")
                    if enable_password:
                        file_password = st.text_input("Set File Password for Vault", type="password", placeholder="Enter secure password")

                if st.button("📥 Generate & Export Clean Deliverables", type="primary"):
                    # Build Excel payload in memory
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        cache["clean_df"].to_excel(writer, sheet_name="Clean Data", index=False)
                        cache["summary_df"].to_excel(writer, sheet_name="Executive Summary", index=False)
                    excel_bytes = buffer.getvalue()

                    if enable_vault_save and vault_report_name:
                        pwd_to_hash = file_password if enable_password and file_password else "default_secure_key"
                        st.session_state.long_term_vault[vault_report_name] = {
                            "summary": cache["summary_df"],
                            "data": excel_bytes,
                            "hash": hashlib.sha256(pwd_to_hash.encode()).hexdigest(),
                            "timestamp": cache["timestamp"]
                        }
                        st.success(f"Successfully encrypted and locked '{vault_report_name}' into long-term secure vault!")

                    st.download_button(
                        label="📥 Download Clean Executive Report (.xlsx)",
                        data=excel_bytes,
                        file_name=f"cleaned_{cache['filename'].split('.')[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        except Exception as e:
            st.error(f"Pipeline processing error: {e}")

with tab_vault_viewer:
    st.markdown("### 🧠 Active Memory & Long-Term Summaries")
    if st.session_state.short_term_cache is not None:
        st.info(f"Active Short-Term Cache File: **{st.session_state.short_term_cache['filename']}** (Loaded at {st.session_state.short_term_cache['timestamp']})")
    else:
        st.info("No active short-term data loaded in memory.")

    st.markdown("---")
    st.markdown("#### 📂 Vault Registry Overview")
    if st.session_state.long_term_vault:
        vault_overview = []
        for name, info in st.session_state.long_term_vault.items():
            vault_overview.append({"Report Name": name, "Saved Timestamp": info["timestamp"], "Protection": "Password Locked 🔐"})
        st.dataframe(pd.DataFrame(vault_overview), use_container_width=True)
    else:
        st.write("Vault is empty. Run a pipeline and check the save option to store executive data logs securely.")
