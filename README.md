# QuickBooks CFO Management Dashboard

Web-based executive dashboard for Social Investment Managers & Advisors LLC —
expense analysis, cost-driver breakdowns, and balance sheet monitoring, built
directly from QuickBooks Online exports.

## What it does

Upload two QBO reports and the dashboard automatically:
- Computes YTD revenue, operating expenses, net operating income, and net
  income, with prior-year deltas compared over the *same* YTD period (not a
  full prior year against a partial current year).
- Breaks down expenses by category, month-by-month, ranked by YTD spend.
- Extracts Total Assets, Total Liabilities, Total Equity, Cash, Accounts
  Receivable, Working Capital, and Current Ratio from the Balance Sheet.
- Everything is parsed live from the uploaded files — nothing is hardcoded.

## Updating data monthly

1. In QuickBooks Online, run these two reports:
   - **Profit and Loss by Month**, for the year to date, with the
     "Previous Year" comparison columns turned on
     (Customize report → Comparison periods → check *Previous year*, and
     make sure the layout keeps a monthly column breakdown).
   - **Balance Sheet**, as of today's date.
2. Export both as Excel (`.xlsx`) from QuickBooks.
3. Open the deployed web app link.
4. In the sidebar, upload:
   - **1. Profit & Loss by Month Report** → the P&L export from step 1.
   - **2. Balance Sheet Report** → the Balance Sheet export from step 2.
5. Click **🚀 Process Financial Dashboard**. The dashboard updates
   immediately from whatever you uploaded.

## Notes

- The parser reads QuickBooks' actual row indentation and `"Total for X"`
  subtotal labels to tell section headers, line items, and subtotals apart —
  so it's tied to how QBO formats these two specific report types, not to
  any hardcoded account names or numbers.
- If a future export doesn't parse cleanly (e.g. QuickBooks changes the
  report layout, or a differently-configured report is uploaded), the app
  will show a specific parsing error rather than silently displaying wrong
  numbers. Two "verification" expanders at the bottom of the dashboard show
  every row the parser extracted, useful for double-checking against the
  source file.
- Prior-year comparison relies on the P&L export having the "Previous Year"
  comparison columns enabled — without them, KPI cards will show YTD figures
  without a YoY delta.

## Deployment

Deployed on Render as a Python web service:
- `requirements.txt` — pinned dependencies (streamlit, pandas, openpyxl).
- `render.yaml` — Render service definition; runs
  `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`.
- `.streamlit/config.toml` — Streamlit server config (headless mode, CORS/XSRF
  settings, max upload size). Note: the port in `config.toml` is overridden
  at runtime by the `--server.port=$PORT` flag in `render.yaml`, since Render
  assigns the port dynamically.
