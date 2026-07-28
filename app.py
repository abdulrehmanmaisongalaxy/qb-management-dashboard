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

# Custom Styling
st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .metric-card { background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_data(uploaded_file):
  try:
    # Read Excel file, skipping header rows typically found in QuickBooks reports
    # QB Transaction Detail reports usually have title rows at the top. We inspect sheet names first.
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=4)
    return df
  except Exception as e:
    st.error(f"Error reading file: {e}")
    return None


def clean_qb_data(df):
  # Drop completely empty rows
  df = df.dropna(how="all")

  # Standardize column names (strip whitespace)
  df.columns = df.columns.str.strip()

  # Identify necessary columns based on standard QuickBooks Transaction Detail report
  # Typical columns: Date, Transaction Type, Num, Name, Memo/Description, Split, Amount, Balance, Account
  date_col = next((col for col in df.columns if "date" in col.lower()), None)
  amount_col = next(
      (col for col in df.columns if "amount" in col.lower()), None
  )
  account_col = next(
      (col for col in df.columns if "account" in col.lower()), None
  )
  type_col = next(
      (col for col in df.columns if "type" in col.lower() or "txn" in col.lower()),
      None,
  )

  if not date_col or not amount_col:
    st.error(
        "Could not automatically detect 'Date' or 'Amount' columns. Please check"
        " your report format."
    )
    return None

  # Clean Date
  df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
  df = df.dropna(subset=[date_col])

  # Clean Amount
  df[amount_col] = (
      df[amount_col]
      .astype(str)
      .str.replace(",", "")
      .str.replace("$", "")
      .astype(float)
  )

  # Extract Month-Year for grouping
  df["Month-Year"] = df[date_col].dt.to_period("M").astype(str)
  df["Year"] = df[date_col].dt.year
  df["Month"] = df[date_col].dt.month_name()

  return df, date_col, amount_col, account_col, type_col


# Dashboard Header
st.title("📊 Executive Management Dashboard")
st.markdown(
    "**Expense Analysis & Cashflow Monitoring** | Upload your monthly QuickBooks"
    " Transaction Detail report below."
)
st.divider()

# Sidebar File Upload
st.sidebar.header("Data Management")
uploaded_file = st.sidebar.file_uploader(
    "Upload QuickBooks Excel Report (.xlsx)", type=["xlsx", "xls"]
)

# Use default sample or uploaded file
if uploaded_file is not None:
  raw_df = load_data(uploaded_file)
else:
  st.sidebar.info(
      "Awaiting file upload. Please upload your QuickBooks report to view the"
      " dashboard."
  )
  # Stop execution until file is uploaded
  st.stop()

if raw_df is not None:
  processed_data = clean_qb_data(raw_df)
  if processed_data:
    df, date_col, amount_col, account_col, type_col = processed_data

    # Sidebar Filters
    st.sidebar.divider()
    st.sidebar.subheader("Filters")

    years = sorted(df["Year"].unique(), reverse=True)
    selected_year = st.sidebar.selectbox("Select Year", years)

    df_filtered = df[df["Year"] == selected_year]

    # --- KPI METRICS ---
    total_volume = df_filtered[amount_col].sum()

    # Define Expenses as negative amounts or specific accounts (Adjust based on QB convention)
    expense_df = df_filtered[df_filtered[amount_col] < 0]
    total_expenses = abs(expense_df[amount_col].sum())

    cash_in_df = df_filtered[df_filtered[amount_col] > 0]
    total_cash_in = cash_in_df[amount_col].sum()

    net_cashflow = total_cash_in - total_expenses

    col1, col2, col3, col4 = st.columns(4)
    with col1:
      st.metric(
          label="Total Inflows",
          value=f"${total_cash_in:,.2f}",
          delta="Cash In",
      )
    with col2:
      st.metric(
          label="Total Expenses",
          value=f"${total_expenses:,.2f}",
          delta="-Outflows",
          delta_color="inverse",
      )
    with col3:
      st.metric(
          label="Net Cashflow",
          value=f"${net_cashflow:,.2f}",
          delta="Surrogate" if net_cashflow >= 0 else "Deficit",
      )
    with col4:
      st.metric(label="Transactions Analyzed", value=len(df_filtered))

    st.divider()

    # --- SECTION 1: EXPENSE ANALYSIS ---
    st.subheader("📉 Expense Analysis")

    exp_col1, exp_col2 = st.columns(2)

    with exp_col1:
      if account_col:
        # Group expenses by account
        account_expenses = (
            expense_df.groupby(account_col)[amount_col]
            .sum()
            .abs()
            .reset_index()
        )
        account_expenses = account_expenses.sort_values(
            by=amount_col, ascending=False
        ).head(10)

        fig_acc = px.bar(
            account_expenses,
            x=amount_col,
            y=account_col,
            orientation="h",
            title="Top 10 Expense Accounts",
            labels={amount_col: "Amount ($)", account_col: "Account"},
            color=amount_col,
            color_continuous_scale="Reds",
        )
        fig_acc.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_acc, use_container_width=True)
      else:
        st.warning("Account column not found for breakdown.")

    with exp_col2:
      # Monthly Expense Trend
      monthly_exp = (
          expense_df.groupby("Month-Year")[amount_col].sum().abs().reset_index()
      )
      fig_trend = px.line(
          monthly_exp,
          x="Month-Year",
          y=amount_col,
          markers=True,
          title="Monthly Expense Trend",
          labels={amount_col: "Total Expenses ($)", "Month-Year": "Month"},
      )
      fig_trend.update_traces(line_color="#e74c3c", line_width=3)
      st.plotly_chart(fig_trend, use_container_width=True)

    st.divider()

    # --- SECTION 2: CASHFLOW MONITORING ---
    st.subheader("💰 Cashflow Monitoring (Inflows vs Outflows)")

    # Aggregate monthly cash in and cash out
    monthly_in = (
        cash_in_df.groupby("Month-Year")[amount_col].sum().reset_index()
    )
    monthly_in["Type"] = "Cash In"

    monthly_out = (
        expense_df.groupby("Month-Year")[amount_col]
        .sum()
        .abs()
        .reset_index()
    )
    monthly_out["Type"] = "Cash Out"

    cashflow_monthly = pd.concat([monthly_in, monthly_out])

    fig_cf = px.bar(
        cashflow_monthly,
        x="Month-Year",
        y=amount_col,
        color="Type",
        barmode="group",
        title="Monthly Cash In vs. Cash Out",
        labels={amount_col: "Amount ($)", "Month-Year": "Month"},
        color_discrete_map={"Cash In": "#2ecc71", "Cash Out": "#e74c3c"},
    )
    st.plotly_chart(fig_cf, use_container_width=True)

    # --- SECTION 3: TRANSACTION DETAILS TABLE ---
    with st.expander("🔍 View Raw Transaction Details"):
      st.dataframe(df_filtered, use_container_width=True)
