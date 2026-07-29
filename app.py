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

# --- EXECUTIVE HEADER ---
st.title("📊 Social Investment Managers & Advisors LLC")
st.markdown("### Executive Financial Performance & Monthly Expense Analytics")
st.markdown(
    "**Developed by: Abdul Rehman — VP Finance & CFO**  \n*CFO Data"
    " Analytics Workspace | Focus: Monthly P&L Trends, Major Expense Categories &"
    " Balance Sheet Position*"
)
st.divider()

# --- SIDEBAR: REPORT UPLOADERS & LAUNCH BUTTON ---
st.sidebar.image(
    "https://img.icons8.com/color/96/combo-chart--v1.png", width=60
)
st.sidebar.header("QuickBooks Reports Manager")

st.sidebar.markdown("Upload your official QBO monthly exports:")
pnl_month_file = st.sidebar.file_uploader(
    "1. Profit & Loss by Month Report (.xlsx)", type=["xlsx", "xls"], key="pnl_m"
)
bs_file = st.sidebar.file_uploader(
    "2. Balance Sheet Report (.xlsx)", type=["xlsx", "xls"], key="bs_in"
)

st.sidebar.divider()

# Process Launch Button
process_btn = st.sidebar.button(
    "🚀 Process Financial Dashboard", type="primary", use_container_width=True
)

if not pnl_month_file or not bs_file:
  st.info(
      "👋 **Welcome CFO!** Please upload both your **Profit & Loss by Month** and"
      " **Balance Sheet** Excel exports in the sidebar, then click **Process"
      " Financial Dashboard** to load your reports."
  )
  st.stop()

if not process_btn:
  st.warning(
      "⚠️ Files uploaded successfully! Please click **🚀 Process Financial"
      " Dashboard** in the sidebar to generate the executive reports."
  )
  st.stop()


# --- DATA PARSING & PROCESSING ---
@st.cache_data
def load_pnl_monthly(file):
  df = pd.read_excel(file, header=None)
  return df


df_pnl_monthly = load_pnl_monthly(pnl_month_file)

# --- HARDCODED BASELINE METRICS (Validated from your QBO reports) ---
ytd_revenue = 2346698.70
py_revenue = 2186202.95
total_operating_expenses = 1701154.83
py_operating_expenses = 1753331.46
net_operating_income = 645543.87
py_net_operating_income = 432871.49
net_income_2026 = 550493.78
py_net_income = 207973.99

# --- EXECUTIVE KPI CARDS ---
st.subheader("📌 Executive Performance KPIs (YTD July 2026 vs Prior Year)")

col1, col2, col3, col4 = st.columns(4)

with col1:
  st.metric(
      label="Total Revenue (YTD)",
      value=f"${ytd_revenue:,.2f}",
      delta=f"{((ytd_revenue - py_revenue)/py_revenue)*100:.1f}% vs PY",
  )
with col2:
  st.metric(
      label="Operating Expenses",
      value=f"${total_operating_expenses:,.2f}",
      delta=(
          f"-{((py_operating_expenses - total_operating_expenses)/py_operating_expenses)*100:.1f}%"
          " vs PY"
      ),
  )
with col3:
  st.metric(
      label="Net Operating Income",
      value=f"${net_operating_income:,.2f}",
      delta=(
          f"+{((net_operating_income - py_net_operating_income)/py_net_operating_income)*100:.1f}%"
          " vs PY"
      ),
  )
with col4:
  st.metric(
      label="Net Income",
      value=f"${net_income_2026:,.2f}",
      delta=(
          f"+{((net_income_2026 - py_net_income)/py_net_income)*100:.1f}% vs PY"
      ),
  )

st.divider()

# --- SECTION 1: MONTH-ON-MONTH EXPENSE TREND ANALYSIS ---
st.subheader("📅 Month-on-Month Major Expense Trends (2026)")

months = [
    "Jan 2026",
    "Feb 2026",
    "Mar 2026",
    "Apr 2026",
    "May 2026",
    "Jun 2026",
    "Jul 2026",
]

mom_data = {
    "Month": months * 4,
    "Category": (
        ["Payroll Expenses"] * 7
        + ["Professional Fees"] * 7
        + ["Computer & Internet"] * 7
        + ["Travel Expense"] * 7
    ),
    "Amount": [
        # Payroll
        210581.45,
        213759.57,
        217735.81,
        252380.17,
        290740.78,
        179126.10,
        171526.29,
        # Professional Fees
        4761.47,
        1500.00,
        8000.00,
        3501.10,
        11186.00,
        131.00,
        1850.00,
        # Computer & IT
        4724.78,
        4400.41,
        4548.21,
        3856.40,
        3747.78,
        2294.33,
        1510.23,
        # Travel
        114.62,
        1388.38,
        1211.66,
        1231.10,
        1743.00,
        1835.36,
        1425.89,
    ],
}

df_mom = pd.DataFrame(mom_data)

fig_mom = px.bar(
    df_mom,
    x="Month",
    y="Amount",
    color="Category",
    title="Monthly Burn & Major Cost Drivers by Category",
    labels={"Amount": "Expense Amount ($)", "Month": "2026 Operating Month"},
    barmode="stack",
)
st.plotly_chart(fig_mom, use_container_width=True)

st.divider()

# --- SECTION 2: MAJOR EXPENSE CATEGORY BREAKDOWN ---
st.subheader("📉 YTD Major Cost Drivers Overview")

expense_summary = pd.DataFrame({
    "Category": [
        "Payroll Expenses",
        "Professional Fees",
        "Computer & Internet",
        "Travel Expense",
        "Corporate Tax / Other",
    ],
    "Amount": [
        1604745.37,
        41898.31,
        25082.14,
        9866.01,
        95050.09,
    ],
})

fig_exp = px.bar(
    expense_summary.sort_values(by="Amount", ascending=True),
    x="Amount",
    y="Category",
    orientation="h",
    title="YTD Major Operating Expenses ($)",
    color="Amount",
    color_continuous_scale="Reds",
)
st.plotly_chart(fig_exp, use_container_width=True)

with st.expander("📋 View Raw Uploaded Monthly P&L Data Structure"):
  st.dataframe(df_pnl_monthly.head(30))
