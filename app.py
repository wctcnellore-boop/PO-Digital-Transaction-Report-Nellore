import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="63 Offices Daily Report Generator", layout="centered")

st.title("📊 Daily Digital Transactions Report Generator")
st.write("Upload your raw daily data file below to instantly generate the clean, formatted Excel report.")

# Only Daily Data file uploader (Office list is loaded automatically from GitHub)
uploaded_file = st.file_uploader(
    "Upload DailyData Excel or CSV File", type=["xlsx", "xls", "csv"]
)

# Automatically load the office list from GitHub repository file
office_list_path = "OfficeList.xlsx"

try:
    office_df = pd.read_excel(office_list_path)
except Exception as e:
    st.error(f"Error loading OfficeList.xlsx from repository: {e}")
    office_df = None

if uploaded_file is not None and office_df is not None:
    try:
        # Conditional reading logic for CSV or Excel formats
        if uploaded_file.name.endswith('.csv'):
            daily_df = pd.read_csv(uploaded_file)
        else:
            daily_df = pd.read_excel(uploaded_file)
            
        st.success("Files loaded successfully! Processing your report...")
        
        # Add your data transformation and processing logic here
        # Example preview of data:
        st.write("Preview of uploaded data:", daily_df.head())

    except Exception as e:
        st.error(f"Error processing the uploaded file: {e}")
