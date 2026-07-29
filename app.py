import re
from dataclasses import dataclass, field

import openpyxl
import pandas as pd
import streamlit as st

# =============================================================================
# PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="Social Investment Managers & Advisors LLC - CFO Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Social Investment Managers & Advisors LLC")
st.markdown("### Executive Financial Performance & Monthly Expense Analytics")
st.markdown(
    "**Developed by: Abdul Rehman — VP Finance & CFO**  \n*CFO Data"
    " Analytics Workspace | Focus: Monthly P&L Trends, Major Expense Categories &"
    " Balance Sheet Position*"
)
st.divider()

MONTH_RE = re.compile(
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*'?\s*(\d{2,4})",
    re.IGNORECASE,
)
TOTAL_LABEL_RE = re.compile(r"^total\s+", re.IGNORECASE)
NET_LABELS = {
    "net income",
    "net operating income",
    "net other income",
    "gross profit",
}

# =============================================================================
# DATA MODEL
# =============================================================================
@dataclass
class ParsedReport:
    # ordered list of (month_label, column_index) actually found in the file
    month_cols: list = field(default_factory=list)
    total_col: int | None = None
    # rows: list of dicts {label, indent, is_total, is_net, section, values: {col: val}}
    rows: list = field(default_factory=list)

    def get_row(self, label_pred):
        for r in self.rows:
            if label_pred(r["label"]):
                return r
        return None

    def total_of(self, exact_label):
        return self.get_row(lambda l: l.strip().lower() == exact_label.lower())

    def value_for(self, row, col=None):
        if row is None:
            return None
        col = col if col is not None else self.total_col
        if col is not None and col in row["values"]:
            return row["values"][col]
        # fall back to summing the month columns
        return sum(row["values"].get(c, 0) or 0 for _, c in self.month_cols)


# =============================================================================
# CORE PARSER (indentation-aware, reads QBO xlsx exports directly via openpyxl)
# =============================================================================
def parse_qbo_report(file) -> ParsedReport:
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active

    report = ParsedReport()

    # --- 1. locate the header row (contains month labels and/or "Total") ---
    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=15):
        month_hits = 0
        total_hit = False
        for cell in row:
            if cell.value is None:
                continue
            val = str(cell.value).strip()
            if MONTH_RE.search(val):
                month_hits += 1
            if val.lower() == "total":
                total_hit = True
        if month_hits >= 1 or total_hit:
            header_row_idx = row[0].row
            if month_hits >= 1:
                break  # prefer a row with actual month labels over a bare "Total" row

    if header_row_idx is None:
        raise ValueError(
            "Could not find a header row with month labels or a 'Total' column. "
            "This file may not be a standard QBO 'P&L by Month' or 'Balance Sheet' export."
        )

    for cell in ws[header_row_idx]:
        if cell.value is None:
            continue
        val = str(cell.value).strip()
        m = MONTH_RE.search(val)
        if m:
            month_name = m.group(1).title()
            report.month_cols.append((month_name, cell.column))
        elif val.lower() == "total":
            report.total_col = cell.column

    # --- 2. walk data rows, using indent level to build hierarchy ---
    current_section = None
    for row in ws.iter_rows(min_row=header_row_idx + 1, max_row=ws.max_row):
        label_cell = row[0]
        label = label_cell.value
        if label is None or str(label).strip() == "":
            continue
        label = str(label).strip()

        indent = 0
        try:
            if label_cell.alignment and label_cell.alignment.indent:
                indent = int(label_cell.alignment.indent)
        except Exception:
            indent = 0

        is_total = bool(TOTAL_LABEL_RE.match(label))
        is_net = label.lower() in NET_LABELS

        if indent == 0 and not is_total and not is_net:
            current_section = label

        values = {}
        for _, col in report.month_cols:
            v = ws.cell(row=label_cell.row, column=col).value
            values[col] = v if isinstance(v, (int, float)) else 0
        if report.total_col is not None:
            v = ws.cell(row=label_cell.row, column=report.total_col).value
            values[report.total_col] = v if isinstance(v, (int, float)) else 0

        report.rows.append(
            {
                "label": label,
                "indent": indent,
                "is_total": is_total,
                "is_net": is_net,
                "section": current_section,
                "values": values,
            }
        )

    # --- fallback: if no indentation info was preserved at all, infer indent
    # from position relative to section headers / total rows ---
    if all(r["indent"] == 0 for r in report.rows):
        section = None
        for r in report.rows:
            if r["is_total"] or r["is_net"]:
                continue
            if r["label"] in ("Income", "Cost of Goods Sold", "Expenses",
                               "Other Income", "Other Expenses",
                               "Assets", "Liabilities", "Equity",
                               "Liabilities and Equity"):
                section = r["label"]
                r["indent"] = 0
            else:
                r["indent"] = 1
                r["section"] = section

    return report


def line_items_under(report: ParsedReport, section_name: str):
    """Return non-total, indented rows belonging to a given top-level section."""
    return [
        r
        for r in report.rows
        if r["section"] == section_name and not r["is_total"] and not r["is_net"] and r["indent"] >= 1
    ]


# =============================================================================
# SIDEBAR — UPLOADERS
# =============================================================================
st.sidebar.image("https://img.icons8.com/color/96/combo-chart--v1.png", width=60)
st.sidebar.header("QuickBooks Reports Manager")
st.sidebar.markdown("Upload your official QBO exports:")

pnl_month_file = st.sidebar.file_uploader(
    "1. Profit & Loss by Month Report (.xlsx)", type=["xlsx", "xls"], key="pnl_m"
)
bs_file = st.sidebar.file_uploader(
    "2. Balance Sheet Report (.xlsx)", type=["xlsx", "xls"], key="bs_in"
)
pnl_py_file = st.sidebar.file_uploader(
    "3. (Optional) Prior-Year P&L Report — for YoY deltas",
    type=["xlsx", "xls"],
    key="pnl_py",
)

st.sidebar.divider()
process_btn = st.sidebar.button(
    "🚀 Process Financial Dashboard", type="primary", use_container_width=True
)

if not pnl_month_file or not bs_file:
    st.info(
        "👋 **Welcome CFO!** Please upload your **Profit & Loss by Month** and"
        " **Balance Sheet** Excel exports in the sidebar, then click **Process"
        " Financial Dashboard**."
    )
    st.stop()

if not process_btn:
    st.warning(
        "⚠️ Files uploaded. Click **🚀 Process Financial Dashboard** in the"
        " sidebar to generate the executive reports."
    )
    st.stop()

# =============================================================================
# PARSE UPLOADED FILES
# =============================================================================
try:
    pnl = parse_qbo_report(pnl_month_file)
except Exception as e:
    st.error(f"Couldn't parse the P&L by Month file: {e}")
    st.stop()

try:
    bs = parse_qbo_report(bs_file)
except Exception as e:
    st.error(f"Couldn't parse the Balance Sheet file: {e}")
    st.stop()

pnl_py = None
if pnl_py_file:
    try:
        pnl_py = parse_qbo_report(pnl_py_file)
    except Exception as e:
        st.warning(f"Prior-year file uploaded but couldn't be parsed, skipping YoY deltas: {e}")
        pnl_py = None

total_income_row = pnl.total_of("Total Income")
total_cogs_row = pnl.total_of("Total Cost of Goods Sold")
total_expenses_row = pnl.total_of("Total Expenses")
net_op_income_row = pnl.get_row(lambda l: l.strip().lower() == "net operating income")
net_income_row = pnl.get_row(lambda l: l.strip().lower() == "net income")

ytd_revenue = pnl.value_for(total_income_row) or 0
total_operating_expenses = pnl.value_for(total_expenses_row) or 0
net_operating_income = pnl.value_for(net_op_income_row) or (ytd_revenue - total_operating_expenses)
net_income = pnl.value_for(net_income_row) or net_operating_income

py_vals = {}
if pnl_py:
    py_vals["revenue"] = pnl_py.value_for(pnl_py.total_of("Total Income")) or 0
    py_vals["opex"] = pnl_py.value_for(pnl_py.total_of("Total Expenses")) or 0
    py_vals["noi"] = pnl_py.value_for(
        pnl_py.get_row(lambda l: l.strip().lower() == "net operating income")
    ) or 0
    py_vals["ni"] = pnl_py.value_for(
        pnl_py.get_row(lambda l: l.strip().lower() == "net income")
    ) or 0


def pct_delta(current, prior):
    if not prior:
        return None
    return (current - prior) / prior * 100


# =============================================================================
# EXECUTIVE KPI CARDS
# =============================================================================
st.subheader("📌 Executive Performance KPIs (YTD)")
col1, col2, col3, col4 = st.columns(4)

kpis = [
    (col1, "Total Revenue (YTD)", ytd_revenue, py_vals.get("revenue")),
    (col2, "Operating Expenses", total_operating_expenses, py_vals.get("opex")),
    (col3, "Net Operating Income", net_operating_income, py_vals.get("noi")),
    (col4, "Net Income", net_income, py_vals.get("ni")),
]

for col, label, value, prior in kpis:
    delta = pct_delta(value, prior)
    with col:
        st.metric(
            label=label,
            value=f"${value:,.2f}",
            delta=f"{delta:+.1f}% vs PY" if delta is not None else None,
        )

if not pnl_py_file:
    st.caption("ℹ️ Upload a prior-year P&L report in the sidebar to see YoY deltas above.")

st.divider()

# =============================================================================
# SECTION 1: MONTH-ON-MONTH EXPENSE TREND (built from actual parsed rows)
# =============================================================================
st.subheader("📅 Month-on-Month Major Expense Trends")

expense_items = line_items_under(pnl, "Expenses")
if not expense_items:
    st.info("No expense line items were detected under an 'Expenses' section in the uploaded file.")
else:
    month_labels = [m for m, _ in pnl.month_cols]
    month_cols_idx = [c for _, c in pnl.month_cols]

    trend_df = pd.DataFrame(
        {
            "Expense Category": [r["label"] for r in expense_items],
            **{
                month_labels[i]: [r["values"].get(month_cols_idx[i], 0) for r in expense_items]
                for i in range(len(month_labels))
            },
            "YTD Total": [
                sum(r["values"].get(c, 0) for c in month_cols_idx) for r in expense_items
            ],
        }
    )
    trend_df = trend_df.sort_values("YTD Total", ascending=False).reset_index(drop=True)

    display_df = trend_df.copy()
    for c in display_df.columns[1:]:
        display_df[c] = display_df[c].apply(lambda x: f"${x:,.2f}")

    st.dataframe(display_df, use_container_width=True, hide_index=True)

st.divider()

# =============================================================================
# SECTION 2: MAJOR EXPENSE CATEGORY BREAKDOWN (top cost drivers, YTD)
# =============================================================================
st.subheader("📉 YTD Major Cost Drivers Overview")

if expense_items:
    top_drivers = trend_df[["Expense Category", "YTD Total"]].head(8)
    st.bar_chart(top_drivers.set_index("Expense Category"))
    driver_display = top_drivers.copy()
    driver_display["YTD Total"] = driver_display["YTD Total"].apply(lambda x: f"${x:,.2f}")
    st.dataframe(driver_display, use_container_width=True, hide_index=True)

st.divider()

# =============================================================================
# SECTION 3: BALANCE SHEET POSITION
# =============================================================================
st.subheader("🏦 Balance Sheet Position")

total_assets = bs.value_for(bs.total_of("Total Assets"))
total_liabilities = bs.value_for(bs.total_of("Total Liabilities"))
total_equity = bs.value_for(bs.total_of("Total Equity"))
cash_row = bs.get_row(lambda l: "cash" in l.lower() and "total" not in l.lower())
cash_position = bs.value_for(cash_row) if cash_row else None

bcol1, bcol2, bcol3, bcol4 = st.columns(4)
with bcol1:
    st.metric("Total Assets", f"${total_assets:,.2f}" if total_assets else "N/A")
with bcol2:
    st.metric("Total Liabilities", f"${total_liabilities:,.2f}" if total_liabilities else "N/A")
with bcol3:
    st.metric("Total Equity", f"${total_equity:,.2f}" if total_equity else "N/A")
with bcol4:
    st.metric("Cash Position", f"${cash_position:,.2f}" if cash_position else "N/A")

st.divider()

with st.expander("📋 View Parsed P&L Rows (debug / verification)"):
    debug_df = pd.DataFrame(
        [
            {
                "Label": r["label"],
                "Section": r["section"],
                "Indent": r["indent"],
                "Is Total": r["is_total"],
                "YTD/Total": pnl.value_for(r),
            }
            for r in pnl.rows
        ]
    )
    st.dataframe(debug_df, use_container_width=True, hide_index=True)

with st.expander("📋 View Parsed Balance Sheet Rows (debug / verification)"):
    debug_bs_df = pd.DataFrame(
        [
            {
                "Label": r["label"],
                "Section": r["section"],
                "Indent": r["indent"],
                "Is Total": r["is_total"],
                "Value": bs.value_for(r),
            }
            for r in bs.rows
        ]
    )
    st.dataframe(debug_bs_df, use_container_width=True, hide_index=True)
