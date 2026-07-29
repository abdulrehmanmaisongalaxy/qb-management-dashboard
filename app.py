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


# --- HELPER: GENERATE DOWNLOADABLE TEMPLATE ---
def generate_sample_template():
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    # 1. PNL Sheet template
    pnl_df = pd.DataFrame([
        ["Social Investment Managers & Advisors LLC", "", ""],
        ["Profit & Loss Comparison", "", ""],
        ["For the period ended", "", ""],
        ["", "", ""],
        ["Accrual Basis", "", ""],
        ["Category", "YTD_2026", "YTD_2025"],
        ["Income", 0.0, 0.0],
        ["Consulting Income", 203434.33, 50704.00],
        ["Grant Income", 218750.00, 348147.56],
        ["Total for Income", 2346698.70, 2186202.95],
        ["Gross Profit", 2346698.70, 2186202.95],
        ["Expenses", 0.0, 0.0],
        ["Software Expenses", 1578.89, 74.99],
        ["Employee Salaries", 1282550.36, 1180557.52],
        ["Total for Expenses", 1701154.83, 1753331.46],
        ["Net Income", 550493.78, 207973.99],
    ])
    pnl_df.to_excel(
        writer, sheet_name="Profit and Loss Comparison", index=False, header=False
    )

    # 2. Balance Sheet template
    bs_df = pd.DataFrame([
        ["Social Investment Managers & Advisors LLC", ""],
        ["Balance Sheet", ""],
        ["As of July 2026", ""],
        ["", ""],
        ["Account", "Balance"],
        ["Checking & Savings", 473065.26],
        ["Total Assets", 4945329.85],
        ["Total Equity", 3325415.33],
    ])
    bs_df.to_excel(
        writer, sheet_name="Balance Sheet", index=False, header=False
    )

    # 3. Updated Transaction Detail template (Matching your General Ledger schema)
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
    td_df.to_excel(
        writer, sheet_name="Transaction Detail", index=False, header=False
    )

  output.seek(0)
  return output.getvalue()


# --- SIDEBAR FORM FOR UPLOADS & TEMPLATE DOWNLOAD ---
st.sidebar.image(
    "https://img.icons8.com/color/96/combo-chart--v1.png", width=60
)
st.sidebar.header("Reports Management")

template_bytes = generate_sample_template()
st.sidebar.download_button(
    label="📥 Download Standard Template",
    data=template_bytes,
    file_name="CFO_Dashboard_Template.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ),
    help=(
        "Download this template, fill in your data, and upload the files"
        " below."
    ),
)

st.sidebar.divider()

with st.sidebar.form("upload_form"):
  st.markdown("Upload all 3 QuickBooks reports below:")
  uploaded_pnl = st.file_uploader(
      "Profit & Loss Comparison (.xlsx)", type=["xlsx", "xls"], key="pnl_in"
  )
  uploaded_bs = st.file_uploader(
      "Balance Sheet (.xlsx)", type=["xlsx", "xls"], key="bs_in"
  )
  uploaded_td = st.file_uploader(
      "Transaction Detail / GL (.xlsx)", type=["xlsx", "xls"], key="td_in"
  )

  submitted = st.form_submit_button(
      "🚀 Process & Load Dashboard", type="primary", use_container_width=True
  )

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
      "👋 **Welcome CFO!** Download the standard template above, populate your"
      " numbers, upload them in the sidebar form, and click **'Process & Load"
      " Dashboard'**."
  )
  st.stop()


# --- ROBUST QUICKBOOKS DATA LOADERS ---
@st.cache_data
def load_pnl(file):
  try:
    df = pd.read_excel(file, header=None)
    header_row_idx = None
    for idx, row in df.iterrows():
      row_str = row.astype(str).str.lower().values
      if any("category" in s or "account" in s for s in row_str):
        header_row_idx = idx
        break

    if header_row_idx is None:
      header_row_idx = 5

    df_data = df.iloc[header_row_idx + 1 :].copy()

    if df_data.shape[1] >= 3:
      df_data = df_data.iloc[:, [0, 1, 2]]
      df_data.columns = ["Category", "YTD_2026", "YTD_2025"]
    elif df_data.shape[1] == 2:
      df_data.columns = ["Category", "YTD_2026"]
      df_data["YTD_2025"] = 0.0
    else:
      return None

    df_data = df_data.dropna(subset=["Category"])
    df_data = df_data[
        ~df_data["Category"]
        .astype(str)
        .str.contains(
            "Accrual Basis|Cash Basis|Prepared|Table|Report",
            case=False,
            na=False,
        )
    ]

    for col in ["YTD_2026", "YTD_2025"]:
      if col in df_data.columns:
        df_data[col] = (
            df_data[col]
            .astype(str)
            .str.replace(",", "")
            .str.replace("$", "")
            .str.replace("—", "0")
            .str.strip()
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
    # Read the cleaned uploader template directly using the header row
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
        st.error(f"Missing required column in Transaction Detail: {req}")
        return None

    # Parse Dates and Amounts
    df_data["Transaction date"] = pd.to_datetime(
        df_data["Transaction date"], errors="coerce"
    )
    df_data = df_data.dropna(subset=["Transaction date"])

    df_data["Amount"] = pd.to_numeric(df_data["Amount"], errors="coerce").fillna(
        0.0
    )

    # Add helper fields for filtering and grouping
    df_data["Year"] = df_data["Transaction date"].dt.year
    df_data["Month-Year"] = (
        df_data["Transaction date"].dt.to_period("M").astype(str)
    )

    # Map missing text fields
    if "Name" not in df_data.columns:
      df_data["Name"] = "Unassigned"
    else:
      df_data["Name"] = df_data["Name"].fillna("Unassigned")

    if "Transaction type" not in df_data.columns:
      df_data["Transaction type"] = "General"
    else:
      df_data["Transaction type"] = df_data["Transaction type"].fillna(
          "General"
      )

    return df_data
  except Exception as e:
    st.error(f"Error loading Transaction Detail: {e}")
    return None


df_pnl = load_pnl(pnl_file) if pnl_file else None
df_bs = load_bs(bs_file) if bs_file else None
df_td = load_td(td_file) if td_file else None

# --- SIDEBAR DYNAMIC FILTERS ---
selected_year = 2026
selected_classification = "All Classifications"
selected_ledger = "All Accounts"
selected_vendor = "All Vendors"

if df_td is not None:
  st.sidebar.subheader("🔍 CFO Filter Controls")
  years = sorted(df_td["Year"].unique(), reverse=True)
  selected_year = st.sidebar.selectbox("Reporting Year", years)
  df_filtered_td = df_td[df_td["Year"] == selected_year]

  classifications = ["All Classifications"] + sorted(
      df_filtered_td["Classification"].astype(str).unique().tolist()
  )
  selected_classification = st.sidebar.selectbox(
      "Filter by Statement Classification", classifications
  )
  if selected_classification != "All Classifications":
    df_filtered_td = df_filtered_td[
        df_filtered_td["Classification"] == selected_classification
    ]

  ledgers = ["All Accounts"] + sorted(
      df_filtered_td["Distribution account"].astype(str).unique().tolist()
  )
  selected_ledger = st.sidebar.selectbox("Filter by Distribution Account", ledgers)
  if selected_ledger != "All Accounts":
    df_filtered_td = df_filtered_td[
        df_filtered_td["Distribution account"] == selected_ledger
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
  exclude_keywords = (
      "Total|Income|Expenses|Profit|Net|Gross|Operating|Earnings"
  )
  pnl_chart_df = df_pnl[
      ~df_pnl["Category"].astype(str).str.contains(exclude_keywords, case=True)
  ]
  pnl_chart_df = pnl_chart_df[
      (pnl_chart_df["YTD_2026"] != 0) | (pnl_chart_df["YTD_2025"] != 0)
  ]

  p1, p2 = st.columns(2)
  with p1:
    if not pnl_chart_df.empty:
      top_pnl = (
          pnl_chart_df.sort_values(by="YTD_2026", ascending=False)
          .head(10)
          .copy()
      )
      fig_pnl_top = px.bar(
          top_pnl,
          x="YTD_2026",
          y="Category",
          orientation="h",
          title="Top P&L Line-Item Accounts (YTD 2026)",
          labels={"YTD_2026": "Amount ($)", "Category": "Account"},
          color="YTD_2026",
          color_continuous_scale="Blues",
      )
      fig_pnl_top.update_layout(yaxis={"categoryorder": "total ascending"})
      st.plotly_chart(fig_pnl_top, use_container_width=True)
    else:
      st.info("No individual line items available for charting.")

  with p2:
    display_pnl = df_pnl.copy()
    if "YTD_2025" in display_pnl.columns and "YTD_2026" in display_pnl.columns:
      display_pnl["Variance ($)"] = (
          display_pnl["YTD_2026"] - display_pnl["YTD_2025"]
      )
      display_pnl["Variance (%)"] = (
          (display_pnl["Variance ($)"] / display_pnl["YTD_2025"].replace(0, 1))
          * 100
      ).round(1)

    st.markdown("**Complete P&L Comparative Statement (2026 vs 2025)**")
    st.markdown(
        display_pnl.to_html(index=False, classes="table table-striped"),
        unsafe_allow_html=True,
    )
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
    st.markdown(
        df_filtered_td.to_html(index=False, classes="table table-striped"),
        unsafe_allow_html=True,
    )
else:
  st.info("Upload Transaction Detail / General Ledger report to view cashflow analytics.")

# --- SECTION 3: BALANCE SHEET VIEWER ---
if df_bs is not None:
  with st.expander("🏛️ View Full Balance Sheet Report"):
    st.markdown(
        df_bs.to_html(index=False, classes="table table-striped"),
        unsafe_allow_html=True,
    )
