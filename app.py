from datetime import datetime, timedelta
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="63 Offices Daily Report Generator", layout="centered")

st.title("📊 Daily Digital Transactions Report Generator")
st.write(
    "Upload your raw daily data file below to instantly generate the clean,"
    " formatted Excel report."
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

    # --- INSERT YOUR MERGING & TRANSFORMATION LOGIC HERE ---
    # Example: processing placeholder dataframe for download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      daily_df.to_excel(writer, index=False, sheet_name="Daily Report")
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Download Formatted Daily Report",
        data=processed_data,
        file_name=f"Daily_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  except Exception as e:
    st.error(f"Error processing the uploaded file: {e}")
