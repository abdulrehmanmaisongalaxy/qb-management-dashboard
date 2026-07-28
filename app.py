import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Management Dashboard - Expense & Cashflow",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Executive Management Dashboard")
st.markdown(
    "**Expense Analysis & Cashflow Monitoring** | Upload your QuickBooks"
    " Transaction Detail report."
)
st.divider()

# Sidebar File Upload & Template info
st.sidebar.header("Data Management")
uploaded_file = st.sidebar.file_uploader(
    "Upload QuickBooks Excel Report (.xlsx)", type=["xlsx", "xls"]
)

with st.sidebar.expander("ℹ️ Expected Report Template"):
  st.markdown("""
    Your QuickBooks **Transaction Detail by Account** report should include:
    * **Transaction Date**
    * **Transaction Type**
    * **Account** (or section headers)
    * **Name** (Vendor / Customer)
    * **Split** (Offset account)
    * **Amount**
    """)

if uploaded_file is None:
  st.info(
      "👈 Please upload your QuickBooks Excel report using the sidebar to"
      " launch the dashboard."
  )
  st.stop()


@st.cache_data
def load_and_clean_data(file):
  try:
    xls = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xls.sheet_names[0], skiprows=4)

    # Drop completely empty rows
    df = df.dropna(how="all")

    # Ensure column names are strings
    df.columns = [str(col).strip() for col in df.columns]

    # Find columns safely
    account_col = df.columns[
        0
    ]  # QuickBooks usually places Account names in the first column
    date_col = next(
        (col for col in df.columns if "date" in str(col).lower()), None
    )
    amount_col = next(
        (col for col in df.columns if "amount" in str(col).lower()), None
    )
    name_col = next(
        (
            col
            for col in df.columns
            if "name" in str(col).lower() and "class" not in str(col).lower()
        ),
        None,
    )
    type_col = next(
        (col for col in df.columns if "type" in str(col).lower()), None
    )

    if not date_col or not amount_col:
      st.error(
          "Could not detect 'Transaction date' or 'Amount' columns in this"
          " file format."
      )
      return None

    # Forward fill account names from section headers
    df["Account"] = df[account_col].ffill()

    # Filter out header rows where date is missing or not a valid date string
    df = df.dropna(subset=[date_col])
    df["Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["Date"])

    # Clean Amount
    df["Amount"] = (
        df[amount_col]
        .astype(str)
        .str.replace(",", "")
        .str.replace("$", "")
        .astype(float)
    )

    df["Year"] = df["Date"].dt.year
    df["Month-Year"] = df["Date"].dt.to_period("M").astype(str)
    df["Month Name"] = df["Date"].dt.month_name()

    return df, name_col, type_col
  except Exception as e:
    st.error(f"Error processing file: {e}")
    return None


result = load_and_clean_data(uploaded_file)
if result:
  df, name_col, type_col = result

  # --- SIDEBAR FILTERS ---
  st.sidebar.divider()
  st.sidebar.subheader("🔍 Dashboard Filters")

  years = sorted(df["Year"].unique(), reverse=True)
  selected_year = st.sidebar.selectbox("Select Year", years)

  df_filtered = df[df["Year"] == selected_year]

  # Optional Vendor / Name filter
  if name_col:
    all_names = ["All"] + sorted(
        df_filtered[name_col].dropna().astype(str).unique().tolist()
    )
    selected_name = st.sidebar.selectbox("Filter by Vendor/Customer", all_names)
    if selected_name != "All":
      df_filtered = df_filtered[df_filtered[name_col].astype(str) == selected_name]

  # --- KPI METRICS ---
  expense_df = df_filtered[df_filtered["Amount"] < 0]
  total_expenses = abs(expense_df["Amount"].sum())

  cash_in_df = df_filtered[df_filtered["Amount"] > 0]
  total_inflows = cash_in_df["Amount"].sum()

  net_cashflow = total_inflows - total_expenses

  col1, col2, col3, col4 = st.columns(4)
  with col1:
    st.metric(
        label="Total Inflows",
        value=f"${total_inflows:,.2f}",
        delta="Cash In",
    )
  with col2:
    st.metric(
        label="Total Expenses",
        value=f"${total_expenses:,.2f}",
        delta="Outflows",
        delta_color="inverse",
    )
  with col3:
    st.metric(
        label="Net Cashflow",
        value=f"${net_cashflow:,.2f}",
        delta="Surplus" if net_cashflow >= 0 else "Deficit",
    )
  with col4:
    st.metric(label="Filtered Transactions", value=len(df_filtered))

  st.divider()

  # --- SECTION 1: EXPENSE ANALYSIS ---
  st.subheader("📉 Expense Analysis")

  exp_col1, exp_col2 = st.columns(2)

  with exp_col1:
    # Top Expense Accounts
    acc_exp = (
        expense_df.groupby("Account")["Amount"].sum().abs().reset_index()
    )
    acc_exp = acc_exp.sort_values(by="Amount", ascending=False).head(10)

    fig_acc = px.bar(
        acc_exp,
        x="Amount",
        y="Account",
        orientation="h",
        title="Top 10 Expense Accounts",
        labels={"Amount": "Total Expense ($)", "Account": "Account Name"},
        color="Amount",
        color_continuous_scale="Reds",
    )
    fig_acc.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_acc, use_container_width=True)

  with exp_col2:
    # Monthly Expense Trend
    monthly_exp = (
        expense_df.groupby("Month-Year")["Amount"].sum().abs().reset_index()
    )
    fig_trend = px.line(
        monthly_exp,
        x="Month-Year",
        y="Amount",
        markers=True,
        title="Monthly Expense Trend",
        labels={"Amount": "Expenses ($)", "Month-Year": "Month"},
    )
    fig_trend.update_traces(line_color="#e74c3c", line_width=3)
    st.plotly_chart(fig_trend, use_container_width=True)

  st.divider()

  # --- SECTION 2: CASHFLOW MONITORING ---
  st.subheader("💰 Cashflow Monitoring")

  monthly_in = cash_in_df.groupby("Month-Year")["Amount"].sum().reset_index()
  monthly_in["Type"] = "Cash In"

  monthly_out = (
      expense_df.groupby("Month-Year")["Amount"].sum().abs().reset_index()
  )
  monthly_out["Type"] = "Cash Out"

  cashflow_monthly = pd.concat([monthly_in, monthly_out])

  fig_cf = px.bar(
      cashflow_monthly,
      x="Month-Year",
      y="Amount",
      color="Type",
      barmode="group",
      title="Monthly Cash In vs. Cash Out",
      labels={"Amount": "Amount ($)", "Month-Year": "Month"},
      color_discrete_map={"Cash In": "#2ecc71", "Cash Out": "#e74c3c"},
  )
  st.plotly_chart(fig_cf, use_container_width=True)

  # --- SECTION 3: TRANSACTION DETAILS TABLE ---
  with st.expander("🔍 View Detailed Transactions Table"):
    st.dataframe(df_filtered, use_container_width=True)
