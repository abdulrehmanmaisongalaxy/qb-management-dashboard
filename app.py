import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Social Investment Managers & Advisors LLC",
    page_icon="📊",
    layout="wide",
)

# --- HEADER & BRANDING ---
st.title("📊 Social Investment Managers & Advisors LLC")
st.markdown("### Management Dashboard")
st.markdown(
    "**Developed by: Abdul Rehman — VP Finance & CFO**  \n*Historical Data"
    " Analytics Workspace | Focus: Expense Analysis & Cashflow Monitoring*"
)
st.divider()

# Sidebar File Upload & Template info
st.sidebar.image(
    "https://img.icons8.com/color/96/combo-chart--v1.png", width=60
)
st.sidebar.header("Data Management")
uploaded_file = st.sidebar.file_uploader(
    "Upload Standardized Template (.xlsx)", type=["xlsx", "xls"]
)

with st.sidebar.expander("ℹ️ Template Field Guide"):
  st.markdown("""
    Your uploaded file should contain these exact columns:
    * **Ledger Name** (e.g., Bank / Account Name)
    * **Transaction date** (MM/DD/YYYY)
    * **Transaction type** (Expense, Transfer, Payment, etc.)
    * **Name** (Vendor or Customer)
    * **Description** (Transaction notes)
    * **Amount** (Positive for Inflows, Negative for Outflows/Expenses)
    """)

if uploaded_file is None:
  st.info(
      "👋 **Welcome!** Please upload your standardized Excel template using"
      " the sidebar to launch the dashboard."
  )
  st.stop()


@st.cache_data
def load_template_data(file):
  try:
    xls = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xls.sheet_names[0])

    # Strip whitespace from column headers
    df.columns = [str(col).strip() for col in df.columns]

    # Required columns check
    if "Transaction date" not in df.columns or "Amount" not in df.columns:
      st.error(
          "Error: Uploaded file is missing required columns ('Transaction"
          " date' or 'Amount')."
      )
      return None

    # Date and Amount formatting
    df["Date"] = pd.to_datetime(df["Transaction date"], errors="coerce")
    df = df.dropna(subset=["Date"])

    df["Amount"] = pd.to_numeric(
        df["Amount"].astype(str).str.replace(",", "").str.replace("$", ""),
        errors="coerce",
    )
    df = df.dropna(subset=["Amount"])

    # Extract time fields
    df["Year"] = df["Date"].dt.year
    df["Month-Year"] = df["Date"].dt.to_period("M").astype(str)
    df["Month Name"] = df["Date"].dt.month_name()

    # Handle optional columns gracefully
    if "Ledger Name" not in df.columns:
      df["Ledger Name"] = "Primary Account"
    else:
      df["Ledger Name"] = df["Ledger Name"].fillna("Primary Account")

    if "Name" not in df.columns:
      df["Name"] = "Unassigned"
    else:
      df["Name"] = df["Name"].fillna("Unassigned")

    if "Transaction type" not in df.columns:
      df["Transaction type"] = "General"

    return df
  except Exception as e:
    st.error(f"Error reading file: {e}")
    return None


df = load_template_data(uploaded_file)

if df is not None:
  # --- SIDEBAR FILTERS ---
  st.sidebar.divider()
  st.sidebar.subheader("🔍 Filter Controls")

  years = sorted(df["Year"].unique(), reverse=True)
  selected_year = st.sidebar.selectbox("Select Year", years)

  df_filtered = df[df["Year"] == selected_year]

  # Ledger Account Filter
  ledgers = ["All"] + sorted(
      df_filtered["Ledger Name"].astype(str).unique().tolist()
  )
  selected_ledger = st.sidebar.selectbox("Filter by Account / Ledger", ledgers)
  if selected_ledger != "All":
    df_filtered = df_filtered[df_filtered["Ledger Name"] == selected_ledger]

  # Vendor / Customer Filter
  vendors = ["All"] + sorted(df_filtered["Name"].astype(str).unique().tolist())
  selected_vendor = st.sidebar.selectbox("Filter by Vendor/Customer", vendors)
  if selected_vendor != "All":
    df_filtered = df_filtered[df_filtered["Name"] == selected_vendor]

  # --- KPI METRICS ---
  expense_df = df_filtered[df_filtered["Amount"] < 0]
  total_expenses = abs(expense_df["Amount"].sum())

  cash_in_df = df_filtered[df_filtered["Amount"] > 0]
  total_inflows = cash_in_df["Amount"].sum()

  net_cashflow = total_inflows - total_expenses

  kpi1, kpi2, kpi3, kpi4 = st.columns(4)
  with kpi1:
    st.metric(
        label="Total Inflows",
        value=f"${total_inflows:,.2f}",
        delta="Cash In",
    )
  with kpi2:
    st.metric(
        label="Total Expenses",
        value=f"${total_expenses:,.2f}",
        delta="Outflows",
        delta_color="inverse",
    )
  with kpi3:
    st.metric(
        label="Net Cashflow",
        value=f"${net_cashflow:,.2f}",
        delta="Surplus" if net_cashflow >= 0 else "Deficit",
    )
  with kpi4:
    st.metric(label="Total Transactions", value=len(df_filtered))

  st.divider()

  # --- SECTION 1: EXPENSE ANALYSIS ---
  st.subheader("📉 Comprehensive Expense Analysis")

  e_col1, e_col2 = st.columns(2)

  with e_col1:
    # Top Vendors by Expense
    vendor_exp = (
        expense_df.groupby("Name")["Amount"].sum().abs().reset_index()
    )
    vendor_exp = vendor_exp.sort_values(by="Amount", ascending=False).head(10)

    fig_vendor = px.bar(
        vendor_exp,
        x="Amount",
        y="Name",
        orientation="h",
        title="Top 10 Expense Vendors / Payees",
        labels={"Amount": "Total Expense ($)", "Name": "Vendor / Name"},
        color="Amount",
        color_continuous_scale="Reds",
    )
    fig_vendor.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_vendor, use_container_width=True)

  with e_col2:
    # Monthly Expense Trend
    monthly_exp = (
        expense_df.groupby("Month-Year")["Amount"].sum().abs().reset_index()
    )
    fig_m_exp = px.line(
        monthly_exp,
        x="Month-Year",
        y="Amount",
        markers=True,
        title="Monthly Expense Trend",
        labels={"Amount": "Total Expense ($)", "Month-Year": "Month"},
    )
    fig_m_exp.update_traces(line_color="#e74c3c", line_width=3)
    st.plotly_chart(fig_m_exp, use_container_width=True)

  st.divider()

  # --- SECTION 2: CASHFLOW MONITORING ---
  st.subheader("💰 Cashflow Monitoring & Liquidity")

  c1, c2 = cash_in_df.groupby("Month-Year")["Amount"].sum().reset_index()
  c1_df = cash_in_df.groupby("Month-Year")["Amount"].sum().reset_index()
  c1_df["Type"] = "Cash Inflows"

  c2_df = (
      expense_df.groupby("Month-Year")["Amount"].sum().abs().reset_index()
  )
  c2_df["Type"] = "Cash Outflows"

  cashflow_trend = pd.concat([c1_df, c2_df])

  fig_cf = px.bar(
      cashflow_trend,
      x="Month-Year",
      y="Amount",
      color="Type",
      barmode="group",
      title="Monthly Cash Inflows vs. Outflows",
      labels={"Amount": "Amount ($)", "Month-Year": "Month"},
      color_discrete_map={
          "Cash Inflows": "#2ecc71",
          "Cash Outflows": "#e74c3c",
      },
  )
  st.plotly_chart(fig_cf, use_container_width=True)

  # --- SECTION 3: DETAILED TABLE ---
  with st.expander("📋 View Filtered Transaction Ledger"):
    st.dataframe(df_filtered, use_container_width=True)
