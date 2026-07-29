import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Social Investment Managers & Advisors LLC - CFO Dashboard",
    page_icon="📊",
    layout="wide",
)

# --- HEADER & EXECUTIVE BRANDING ---
st.title("📊 Social Investment Managers & Advisors LLC")
st.markdown("### Executive Financial Performance, Position & Cashflow Dashboard")
st.markdown(
    "**Developed by: Abdul Rehman — VP Finance & CFO**  \n*CFO Data"
    " Analytics Workspace | Focus: Profitability, YoY Variances, Cash Position &"
    " Expense Analysis*"
)
st.divider()


# --- HELPER: GENERATE DOWNLOADABLE UNIFIED TEMPLATE ---
def generate_sample_template():
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    # Single unified template sheet matching your exact General Ledger schema
    td_df = pd.DataFrame([
        [
            "Classification",
            "Distribution account",
            "Transaction date",
            "Transaction type",
            "Num",
            "Name",
            "Description",
            "Split",
            "Amount",
            "Balance",
        ],
        [
            "Asset",
            "Chase Bank - Checking 0102",
            "2026-01-02",
            "Expense",
            "",
            "Javed Rizvi",
            "Dividend 2025",
            "Retained Earnings",
            -17157.54,
            83006.26,
        ],
        [
            "Expense",
            "Advertising and Marketing",
            "2026-01-14",
            "Expense",
            "",
            "Mind Reflections",
            "",
            "Chase Bank - Checking 0102",
            -483.13,
            451597.01,
        ],
    ])
    td_df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)

  output.seek(0)
  return output.getvalue()


# --- SIDEBAR FORM FOR UNIFIED UPLOAD & TEMPLATE DOWNLOAD ---
st.sidebar.image(
    "https://img.icons8.com/color/96/combo-chart--v1.png", width=60
)
st.sidebar.header("Reports Management")

template_bytes = generate_sample_template()
st.sidebar.download_button(
    label="📥 Download Unified Template",
    data=template_bytes,
    file_name="Uploading Template.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ),
    help=(
        "Download this standard template, fill in your GL data, and upload it"
        " below."
    ),
)

st.sidebar.divider()

with st.sidebar.form("upload_form"):
  st.markdown("Upload your standardized General Ledger file:")
  uploaded_file = st.file_uploader(
      "Uploading Template (.xlsx)", type=["xlsx", "xls"], key="gl_in"
  )

  submitted = st.form_submit_button(
      "🚀 Process & Load Dashboard", type="primary", use_container_width=True
  )

if submitted:
  if uploaded_file is not None:
    st.session_state["gl_file"] = uploaded_file

gl_file = st.session_state.get("gl_file", None)

if not gl_file:
  st.info(
      "👋 **Welcome CFO!** Download the unified template above, populate your"
      " General Ledger data, upload it in the sidebar form, and click **'Process"
      " & Load Dashboard'**."
  )
  st.stop()


# --- ROBUST UNIFIED GENERAL LEDGER DATA LOADER ---
@st.cache_data
def load_gl_data(file):
  try:
    df_data = pd.read_excel(file)

    # Standardize column names
    df_data.columns = [str(col).strip() for col in df_data.columns]

    # Ensure required columns exist
    required_cols = [
        "Classification",
        "Distribution account",
        "Transaction date",
        "Amount",
    ]
    for req in required_cols:
      if req not in df_data.columns:
        st.error(f"Missing required column in uploaded template: {req}")
        return None

    # Parse Dates and Amounts
    df_data["Transaction date"] = pd.to_datetime(
        df_data["Transaction date"], errors="coerce"
    )
    df_data = df_data.dropna(subset=["Transaction date"])

    df_data["Amount"] = pd.to_numeric(df_data["Amount"], errors="coerce").fillna(
        0.0
    )
    if "Balance" in df_data.columns:
      df_data["Balance"] = pd.to_numeric(
          df_data["Balance"], errors="coerce"
      ).fillna(0.0)

    # Add helper fields for filtering and grouping
    df_data["Year"] = df_data["Transaction date"].dt.year
    df_data["Month-Year"] = (
        df_data["Transaction date"].dt.to_period("M").astype(str)
    )

    # Map missing text fields safely
    for col, default_val in [
        ("Name", "Unassigned"),
        ("Transaction type", "General"),
        ("Description", ""),
        ("Split", ""),
        ("Num", ""),
    ]:
      if col not in df_data.columns:
        df_data[col] = default_val
      else:
        df_data[col] = df_data[col].fillna(default_val)

    return df_data
  except Exception as e:
    st.error(f"Error loading General Ledger data: {e}")
    return None


df_gl = load_gl_data(gl_file) if gl_file else None

# --- SIDEBAR DYNAMIC FILTERS ---
selected_year = 2026
selected_classification = "All Classifications"
selected_ledger = "All Accounts"
selected_vendor = "All Vendors"

if df_gl is not None:
  st.sidebar.subheader("🔍 CFO Filter Controls")
  years = sorted(df_gl["Year"].unique(), reverse=True)
  selected_year = st.sidebar.selectbox("Reporting Year", years)
  df_filtered = df_gl[df_gl["Year"] == selected_year]

  classifications = ["All Classifications"] + sorted(
      df_filtered["Classification"].astype(str).unique().tolist()
  )
  selected_classification = st.sidebar.selectbox(
      "Filter by Statement Classification", classifications
  )
  if selected_classification != "All Classifications":
    df_filtered = df_filtered[
        df_filtered["Classification"] == selected_classification
    ]

  ledgers = ["All Accounts"] + sorted(
      df_filtered["Distribution account"].astype(str).unique().tolist()
  )
  selected_ledger = st.sidebar.selectbox("Filter by Distribution Account", ledgers)
  if selected_ledger != "All Accounts":
    df_filtered = df_filtered[
        df_filtered["Distribution account"] == selected_ledger
    ]

  vendors = ["All Vendors"] + sorted(
      df_filtered["Name"].astype(str).unique().tolist()
  )
  selected_vendor = st.sidebar.selectbox("Filter by Vendor / Payee", vendors)
  if selected_vendor != "All Vendors":
    df_filtered = df_filtered[df_filtered["Name"] == selected_vendor]
else:
  df_filtered = None

# --- EXECUTIVE FINANCIAL KPIS (Calculated directly from GL Data) ---
st.subheader("📌 Executive Financial Position & Performance")

col1, col2, col3, col4, col5 = st.columns(5)

# Calculate metrics dynamically from the full GL upload
total_rev = (
    df_gl[df_gl["Classification"] == "Revenue"]["Amount"].sum()
    if df_gl is not None
    else 0.0
)
total_exp = (
    df_gl[df_gl["Classification"] == "Expense"]["Amount"].sum()
    if df_gl is not None
    else 0.0
)
net_income = total_rev + total_exp  # expenses are negative amounts

total_assets = (
    df_gl[df_gl["Classification"] == "Asset"]["Balance"].iloc[-1]
    if df_gl is not None and not df_gl[df_gl["Classification"] == "Asset"].empty
    else 0.0
)
total_liabilities = (
    df_gl[df_gl["Classification"] == "Liability"]["Balance"].iloc[-1]
    if df_gl is not None
    and not df_gl[df_gl["Classification"] == "Liability"].empty
    else 0.0
)

with col1:
  st.metric(label="Total Revenue (YTD)", value=f"${total_rev:,.2f}")
with col2:
  st.metric(label="Net Income (YTD)", value=f"${net_income:,.2f}")
with col3:
  st.metric(
      label="Total Outflows / Expenses", value=f"${abs(total_exp):,.2f}"
  )
with col4:
  st.metric(label="Total Tracked Assets", value=f"${total_assets:,.2f}")
with col5:
  st.metric(label="Total Tracked Liabilities", value=f"${total_liabilities:,.2f}")

st.divider()

# --- SECTION 1: ACCOUNT CLASSIFICATION & BREAKDOWN ---
st.subheader(
    "📈 General Ledger Breakdown by Classification & Account Dynamics"
)

if df_filtered is not None:
  p1, p2 = st.columns(2)
  with p1:
    class_summary = (
        df_filtered.groupby("Classification")["Amount"].sum().reset_index()
    )
    fig_class = px.bar(
        class_summary,
        x="Classification",
        y="Amount",
        title=f"Net Movement by Financial Classification ({selected_year})",
        labels={"Amount": "Net Amount ($)", "Classification": "Category"},
        color="Classification",
        color_discrete_sequence=px.colors.qualitative.Prism,
    )
    st.plotly_chart(fig_class, use_container_width=True)

  with p2:
    top_accounts = (
        df_filtered.groupby("Distribution account")["Amount"]
        .sum()
        .reset_index()
    )
    top_accounts = top_accounts.sort_values(
        by="Amount", ascending=True
    ).head(10)
    fig_acc = px.bar(
        top_accounts,
        x="Amount",
        y="Distribution account",
        orientation="h",
        title=f"Top Distribution Accounts by Activity ({selected_year})",
        labels={"Amount": "Amount ($)", "Distribution account": "Account"},
        color="Amount",
        color_continuous_scale="Blues",
    )
    st.plotly_chart(fig_acc, use_container_width=True)

  st.divider()

  # --- SECTION 2: VENDOR & CASHFLOW MONITORING ---
  st.subheader(
      f"📉 Vendor Outflows & Monthly Cashflow Dynamics ({selected_year})"
  )

  income_df = df_filtered[df_filtered["Amount"] > 0]
  expense_df = df_filtered[df_filtered["Amount"] < 0]

  e1, e2 = st.columns(2)
  with e1:
    vendor_exp = (
        expense_df.groupby("Name")["Amount"].sum().abs().reset_index()
    )
    vendor_exp = vendor_exp.sort_values(by="Amount", ascending=False).head(10)
    fig_vendor = px.bar(
        vendor_exp,
        x="Amount",
        y="Name",
        orientation="h",
        title=f"Top Payees / Vendors ({selected_year})",
        labels={"Amount": "Total Outflow ($)", "Name": "Vendor Name"},
        color="Amount",
        color_continuous_scale="Reds",
    )
    fig_vendor.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_vendor, use_container_width=True)

  with e2:
    monthly_in = income_df.groupby("Month-Year")["Amount"].sum().reset_index()
    monthly_in["Type"] = "Cash Inflows"
    monthly_out = (
        expense_df.groupby("Month-Year")["Amount"].sum().abs().reset_index()
    )
    monthly_out["Type"] = "Cash Outflows"
    cf_trend = pd.concat([monthly_in, monthly_out])

    fig_cf = px.bar(
        cf_trend,
        x="Month-Year",
        y="Amount",
        color="Type",
        barmode="group",
        title=f"Monthly Cash Inflows vs Outflows ({selected_year})",
        labels={"Amount": "Amount ($)", "Month-Year": "Month"},
        color_discrete_map={
            "Cash Inflows": "#2ecc71",
            "Cash Outflows": "#e74c3c",
        },
    )
    st.plotly_chart(fig_cf, use_container_width=True)

  with st.expander("📋 View Detailed Filtered Transaction Ledger"):
    st.markdown(
        df_filtered.to_html(index=False, classes="table table-striped"),
        unsafe_allow_html=True,
    )
else:
  st.info("Upload your 'Uploading Template.xlsx' file to view analytics.")
