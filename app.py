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

# File uploader supporting both Excel and CSV formats
uploaded_file = st.file_uploader(
    "Upload DailyData Excel or CSV File", type=["xlsx", "xls", "csv"]
)

# Automatically load the office list reference file from the GitHub repository
office_list_path = "OfficeList.xlsx"

try:
  office_df = pd.read_excel(office_list_path)
except Exception as e:
  st.error(f"Error loading OfficeList.xlsx from repository: {e}")
  office_df = None

if uploaded_file is not None and office_df is not None:
  try:
    # Conditional reading based on file type
    if uploaded_file.name.endswith(".csv"):
      daily_df = pd.read_csv(uploaded_file)
    else:
      daily_df = pd.read_excel(uploaded_file)

    st.success("Files loaded successfully! Processing your report...")

    # --- DATA TRANSFORMATION & MERGING LOGIC ---
    # Merge the uploaded daily transactions with your official 63 offices reference list.
    # Adjust the key column names ('Office ID' or 'Office Name') to match your exact file structure.
    if "Office ID" in daily_df.columns and "Office ID" in office_df.columns:
      merged_df = pd.merge(
          office_df, daily_df, on="Office ID", how="left"
      ).fillna(0)
    else:
      # Fallback merge or raw processing if column names differ
      merged_df = pd.merge(
          office_df, daily_df, on="Office Name", how="left"
      ).fillna(0)

    # Display preview of processed report
    st.write("Preview of Formatted Report:", merged_df.head())

    # --- GENERATE EXCEL DOWNLOAD FILE ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      merged_df.to_excel(writer, index=False, sheet_name="Daily Report")
    processed_data = output.getvalue()

    # Download button trigger
    st.download_button(
        label="📥 Download Formatted Daily Report",
        data=processed_data,
        file_name=f"Daily_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  except Exception as e:
    st.error(f"Error processing the uploaded file: {e}")
