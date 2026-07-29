import pandas as pd


def parse_pnl_file(uploaded_file):
  """Reads and parses the QuickBooks Profit & Loss by Month export."""
  try:
    df = pd.read_excel(uploaded_file, header=None)
    return df
  except Exception as e:
    raise ValueError(f"Error parsing Profit & Loss file: {e}")


def parse_balance_sheet_file(uploaded_file):
  """Reads and parses the QuickBooks Balance Sheet export."""
  try:
    df = pd.read_excel(uploaded_file, header=None)
    return df
  except Exception as e:
    raise ValueError(f"Error parsing Balance Sheet file: {e}")
