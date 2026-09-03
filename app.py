from datetime import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="63 Offices Daily Report Generator", layout="centered")

st.title("📊 Daily Digital Transactions Report Generator")
st.write(
    "Upload your raw daily data file below to instantly generate the clean,"
    " formatted Excel report for all 63 offices."
)

uploaded_file = st.file_uploader(
    "Upload DailyData Excel or CSV File", type=["xlsx", "xls", "csv"]
)

office_list_path = "OfficeList.xlsx"

try:
  office_df = pd.read_excel(office_list_path)
except Exception as e:
  st.error(f"Error loading OfficeList.xlsx from repository: {e}")
  office_df = None

if uploaded_file is not None and office_df is not None:
  try:
    if uploaded_file.name.endswith(".csv"):
      daily_df = pd.read_csv(uploaded_file)
    else:
      daily_df = pd.read_excel(uploaded_file)

    st.success("Files loaded successfully! Processing your report...")

    # --- SAFE DYNAMIC MERGING LOGIC ---
    # Automatically find common columns or match by index if names differ
    common_cols = [
        col for col in daily_df.columns if col in office_df.columns
    ]

    if len(common_cols) > 0:
      # Use the first matching column (like Office ID)
      match_col = common_cols[0]
      office_df[match_col] = office_df[match_col].astype(str)
      daily_df[match_col] = daily_df[match_col].astype(str)
      merged_df = pd.merge(office_df, daily_df, on=match_col, how="left").fillna(
          0
      )
    else:
      # Fallback: if columns don't match by name, join side-by-side or keep office list intact
      merged_df = office_df.copy()
      for col in daily_df.columns:
        if col not in merged_df.columns:
          merged_df[col] = daily_df[col]
      merged_df = merged_df.fillna(0)

    st.write("Preview of Formatted Report:", merged_df.head())

    # --- GENERATE EXCEL DOWNLOAD FILE ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      merged_df.to_excel(writer, index=False, sheet_name="Daily Report")
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Download Formatted Daily Report",
        data=processed_data,
        file_name=f"Daily_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  except Exception as e:
    st.error(f"Error processing the uploaded file: {e}")
