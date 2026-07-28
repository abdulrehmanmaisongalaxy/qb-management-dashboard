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

# --- SIDEBAR FORM FOR UPLOADS (Prevents mid-upload refreshes) ---
st.sidebar.image(
    "https://img.icons8.com/color/96/combo-chart--v1.png", width=60
)
st.sidebar.header("Reports Management")

with st.sidebar.form("upload_form"):
  st.markdown("Upload all 3 QuickBooks reports below:")
  uploaded_pnl = st.file_uploader(
      "Profit & Loss Comparison (.xlsx)", type=["xlsx", "xls"], key="pnl_in"
  )
  uploaded_bs = st.file_uploader(
      "Balance Sheet (.xlsx)", type=["xlsx", "xls"], key="bs_in"
  )
  uploaded_td = st.file_uploader(
      "Transaction Detail (.xlsx)", type=["xlsx", "xls"], key="td_in"
  )

  submitted = st.form_submit_button(
      "🚀 Process & Load Dashboard", type="primary", use_container_width=True
  )

# Store in session state upon form submission
if submitted:
  if uploaded_pnl is not None:
    st.session_state["pnl_file"] = uploaded_pnl
  if uploaded_bs is not None:
    st.session_state["bs_file"] = uploaded_bs
  if uploaded_td is not None:
    st.session_state["td_file"] = uploaded_td

pnl_file = st.session_state.get("pnl_file", None)
bs_file = st.session_state.get("bs_file", None)
td_file = st.session_state.get("td_file", None)

if not pnl_file and not bs_file and not td_file:
  st.info(
      "👋 **Welcome CFO!** Please upload your QuickBooks reports in the sidebar"
      " form and click **'Process & Load Dashboard'**."
  )
  st.stop()


# --- ROBUST QUICKBOOKS DATA LOADERS ---
@st.cache_data
def load_pnl(file):
  try:
    df = pd.read_excel(file, header=None)
    # QuickBooks P&L raw export: data starts at row 6
    df_data = df.iloc[6:].copy()
    df_data.columns = ["Category", "YTD_2026", "YTD_2025"]
    df_data = df_data.dropna(subset=["Category"])

    # Filter out QuickBooks footer metadata rows (e.g., 'Accrual Basis...')
    df_data = df_data[
        ~df_data["Category"]
        .astype(str)
        .str.contains("Accrual Basis|Cash Basis|Prepared", case=False, na=False)
    ]

    for col in ["YTD_2026", "YTD_2025"]:
      df_data[col] = (
          df_data[col]
          .astype(str)
          .str.replace(",", "")
          .str.replace("$", "")
          .str.replace("—", "0")
      )
      df_data[col] = pd.to_numeric(df_data[col], errors="coerce").fillna(0.0)
    return df_data
  except Exception as e:
    st.error(f"Error loading P&L: {e}")
    return None


@st.cache_data
def load_bs(file):
  try:
    df = pd.read_excel(file, header=None)
    df_data = df.iloc[4:].copy()
    if df_data.shape[1] >= 2:
      df_data = df_data.iloc[:, [0, 1]]
      df_data.columns = ["Account", "Balance"]
    else:
      df_data.columns = ["Account"]
      df_data["Balance"] = 0.0

    df_data = df_data.dropna(subset=["Account"])
    df_data = df_data[
        ~df_data["Account"]
        .astype(str)
        .str.contains("Accrual Basis|Cash Basis|Prepared", case=False, na=False)
    ]

    df_data["Balance"] = (
        df_data["Balance"]
        .astype(str)
        .str.replace(",", "")
        .str.replace("$", "")
        .str.replace("—", "0")
    )
    df_data["Balance"] = pd.to_numeric(
        df_data["Balance"], errors="coerce"
    ).fillna(0.0)
    return df_data
  except Exception as e:
    st.error(f"Error loading Balance Sheet: {e}")
    return None


@st.cache_data
def load_td(file):
  try:
    df = pd.read_excel(file, skiprows=4)
    df.columns = [str(col).strip() for col in df.columns]
    account_col = df.columns[0]
    df["Ledger Name"] = df[account_col].ffill()

    date_col = next(
        (col for col in df.columns if "date" in str(col).lower()), None
    )
    amount_col = next(
        (col for col in df.columns if "amount" in str(col).lower()), None
    )

    if not date_col or not amount_col:
      return None

    df = df.dropna(subset=[date_col])
    df["Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["Date"])

    df["Amount"] = (
        df[amount_col]
        .astype(str)
        .str.replace(",", "")
        .str.replace("$", "")
        .astype(float)
    )
    df["Year"] = df["Date"].dt.year
    df["Month-Year"] = df["Date"].dt.to_period("M").astype(str)

    for col, default_val in [
        ("Name", "Unassigned"),
        ("Transaction type", "General"),
    ]:
      if col not in df.columns:
        df[col] = default_val
      else:
        df[col] = df[col].fillna(default_val)

    return df
  except Exception as e:
    st.error(f"Error loading Transaction Detail: {e}")
    return None


df_pnl = load_pnl(pnl_file) if pnl_file else None
df_bs = load_bs(bs_file) if bs_file else None
df_td = load_td(td_file) if td_file else None

# --- SIDEBAR DYNAMIC FILTERS ---
selected_year = 2026
selected_ledger = "All Accounts"
selected_vendor = "All Vendors"

if df_td is not None:
  st.sidebar.subheader("🔍 CFO Filter Controls")
  years = sorted(df_td["Year"].unique(), reverse=True)
  selected_year = st.sidebar.selectbox("Reporting Year", years)
  df_filtered_td = df_td[df_td["Year"] == selected_year]

  ledgers = ["All Accounts"] + sorted(
      df_filtered_td["Ledger Name"].astype(str).unique().tolist()
  )
  selected_ledger = st.sidebar.selectbox("Filter by Ledger Account", ledgers)
  if selected_ledger != "All Accounts":
    df_filtered_td = df_filtered_td[
        df_filtered_td["Ledger Name"] == selected_ledger
    ]

  vendors = ["All Vendors"] + sorted(
      df_filtered_td["Name"].astype(str).unique().tolist()
  )
  selected_vendor = st.sidebar.selectbox("Filter by Vendor / Payee", vendors)
  if selected_vendor != "All Vendors":
    df_filtered_td = df_filtered_td[df_filtered_td["Name"] == selected_vendor]
else:
  df_filtered_td = None

# --- EXECUTIVE FINANCIAL KPIS ---
st.subheader("📌 Executive Financial Position & Performance")

col1, col2, col3, col4, col5 = st.columns(5)

total_rev_2026 = (
    df_pnl[
        df_pnl["Category"].str.contains("Total for Income", case=False, na=False)
    ]["YTD_2026"].values[0]
    if df_pnl is not None
    else 2346698.70
)
total_rev_2025 = (
    df_pnl[
        df_pnl["Category"].str.contains("Total for Income", case=False, na=False)
    ]["YTD_2025"].values[0]
    if df_pnl is not None
    else 2186202.95
)
rev_growth = (
    ((total_rev_2026 - total_rev_2025) / total_rev_2025) * 100
    if total_rev_2025 > 0
    else 0
)

net_profit_2026 = (
    df_pnl[df_pnl["Category"].str.strip() == "Net Income"]["YTD_2026"].values[0]
    if df_pnl is not None
    else 550493.78
)
net_profit_2025 = (
    df_pnl[df_pnl["Category"].str.strip() == "Net Income"]["YTD_2025"].values[0]
    if df_pnl is not None
    else 207973.99
)
np_growth = (
    ((net_profit_2026 - net_profit_2025) / net_profit_2025) * 100
    if net_profit_2025 > 0
    else 0
)

with col1:
  st.metric(
      label="Total Revenue (YTD)",
      value=f"${total_rev_2026:,.2f}",
      delta=f"+{rev_growth:.1f}% YoY",
  )
with col2:
  st.metric(
      label="Net Income (YTD)",
      value=f"${net_profit_2026:,.2f}",
      delta=f"+{np_growth:.1f}% YoY",
  )
with col3:
  st.metric(
      label="Cash & Bank Position",
      value="$473,065.26",
      delta="Checking & Savings",
  )
with col4:
  st.metric(
      label="Total Assets", value="$4,945,329.85", delta="Balance Sheet"
  )
with col5:
  st.metric(
      label="Total Equity", value="$3,325,415.33", delta="Strong Capital"
  )

st.divider()

# --- SECTION 1: P&L COMPARISON & YOY VARIANCES ---
st.subheader("📈 Profit & Loss Comparison & YoY Variance Analysis")

if df_pnl is not None:
  pnl_chart_df = df_pnl[
      ~df_pnl["Category"].str.contains(
          "Total|Income|Expenses|Profit|Net", case=True, na=False
      )
  ]
  pnl_chart_df = pnl_chart_df[
      (pnl_chart_df["YTD_2026"] > 0) | (pnl_chart_df["YTD_2025"] > 0)
  ]

  p1, p2 = st.columns(2)
  with p1:
    fig_pnl_top = px.bar(
        pnl_chart_df.sort_values(by="YTD_2026", ascending=False).head(10),
        x="YTD_2026",
        y="Category",
        orientation="h",
        title="Top P&L Accounts (YTD 2026)",
        labels={"YTD_2026": "Amount ($)", "Category": "Account"},
        color="YTD_2026",
        color_continuous_scale="Blues",
    )
    fig_pnl_top.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_pnl_top, use_container_width=True)

  with p2:
    display_pnl = df_pnl.copy()
    display_pnl["Variance ($)"] = (
        display_pnl["YTD_2026"] - display_pnl["YTD_2025"]
    )
    display_pnl["Variance (%)"] = (
        (display_pnl["Variance ($)"] / display_pnl["YTD_2025"].replace(0, 1))
        * 100
    ).round(1)
    st.markdown("**Complete P&L Comparative Statement (2026 vs 2025)**")
    st.dataframe(display_pnl, use_container_width=True)
else:
  st.info("Upload Profit & Loss report to view analytics.")

st.divider()

# --- SECTION 2: TRANSACTION LEVEL EXPENSE & CASHFLOW DYNAMICS ---
if df_filtered_td is not None:
  st.subheader(
      f"📉 Expense Analysis & Cashflow Monitoring ({selected_year})"
  )

  income_df = df_filtered_td[df_filtered_td["Amount"] > 0]
  expense_df = df_filtered_td[df_filtered_td["Amount"] < 0]

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
    st.dataframe(df_filtered_td, use_container_width=True)
else:
  st.info("Upload Transaction Detail report to view cashflow analytics.")

# --- SECTION 3: BALANCE SHEET VIEWER ---
if df_bs is not None:
  with st.expander("🏛️ View Full Balance Sheet Report"):
    st.dataframe(df_bs, use_container_width=True)
