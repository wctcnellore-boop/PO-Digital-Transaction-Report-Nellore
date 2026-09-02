import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="63 Offices Daily Report Generator", layout="centered")

st.title("📊 Daily Digital Transactions Report Generator")
st.write("Upload your raw data files to instantly generate the clean, formatted Excel report.")

# File Uploaders
uploaded_data = st.file_uploader("Upload DailyData Excel File", type=["xlsx", "xls"])
uploaded_list = st.file_uploader("Upload OfficeList Excel File", type=["xlsx", "xls"])

if uploaded_data and uploaded_list:
    if st.button("🚀 Generate Report", type="primary"):
        with st.spinner("Processing data, purging zeros, and formatting report..."):
            # Load raw data
            df_data = pd.read_excel(uploaded_data, sheet_name=0)
            df_list = pd.read_excel(uploaded_list, sheet_name=0)
            
            # Load office lookup list (Column A)
            target_offices = set(df_list.iloc[:, 0].astype(str).str.strip().str.upper())
            
            # Clean office names in main data (Column B / Index 1)
            df_data.iloc[:, 1] = df_data.iloc[:, 1].astype(str).str.strip()
            
            # Filter rows where Office Name exists in OfficeList
            filtered_df = df_data[df_data.iloc[:, 1].str.upper().isin(target_offices)].copy()
            
            # Extract required columns mapping (0-based): B=1 (Office), D=3 (Cash), P=15 (DQR), T=19 (POS-Card), V=21 (POS-QR), Z=25 (UPI)
            output_rows = []
            
            for _, row in filtered_df.iterrows():
                office_name = row.iloc[1]
                cash_cnt = pd.to_numeric(row.iloc[3], errors='coerce') or 0
                dqr_cnt = pd.to_numeric(row.iloc[15], errors='coerce') or 0
                pos_card_cnt = pd.to_numeric(row.iloc[19], errors='coerce') or 0
                pos_qr_cnt = pd.to_numeric(row.iloc[21], errors='coerce') or 0
                epay_cnt = pd.to_numeric(row.iloc[25], errors='coerce') or 0
                
                total_digital = dqr_cnt + pos_card_cnt + pos_qr_cnt + epay_cnt
                total_trans = cash_cnt + total_digital
                
                # PURGE ZERO TOTAL TRANSACTIONS
                if total_trans > 0:
                    pct_digital = (total_digital / total_trans) if total_trans > 0 else 0.0
                    output_rows.append({
                        "Office Name": office_name,
                        "Cash (Cnt)": cash_cnt,
                        "DQR Scan (Cnt)": dqr_cnt,
                        "SBI POS-CARD (Cnt)": pos_card_cnt,
                        "SBI POS BHARAT QR (Cnt)": pos_qr_cnt,
                        "SBI E PAY UPI (Cnt)": epay_cnt,
                        "Total Digital Transactions": total_digital,
                        "Total Transactions": total_trans,
                        "% of Digital Transactions": pct_digital
                    })
            
            report_df = pd.DataFrame(output_rows)
            
            # 3-LEVEL SORT:
            # 1. % Digital -> Ascending
            # 2. Cash Count -> Descending
            # 3. Total Digital -> Ascending
            report_df.sort_values(
                by=["% of Digital Transactions", "Cash (Cnt)", "Total Digital Transactions"],
                ascending=[True, False, True],
                inplace=True
            )
            report_df.reset_index(drop=True, inplace=True)
            
            # Build OpenPyXL Workbook in Memory
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DailyReport"
            
            # Title Banner
            report_date = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
            ws.merge_cells("A1:I1")
            title_cell = ws["A1"]
            title_cell.value = f"Nellore Division: Digital Transactions information dated {report_date}"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="0F2043", end_color="0F2043", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 30
            
            # Table Headers
            headers = ["Office Name", "Cash (Cnt)", "DQR Scan (Cnt)", "SBI POS-CARD (Cnt)", 
                       "SBI POS BHARAT QR (Cnt)", "SBI E PAY UPI (Cnt)", 
                       "Total Digital Transactions", "Total Transactions", "% of Digital Transactions"]
            formulas = ["a", "b", "c", "d", "e", "f", "g=c+d+e+f", "h=b+g", "i=(g/h)*100"]
            
            ws.append([]) # Row 2 Blank
            ws.append(headers) # Row 3
            ws.append(formulas) # Row 4
            
            # Header Styles
            header_fill = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
            for r in [3, 4]:
                ws.row_dimensions[r].height = 24
                for col_idx in range(1, 10):
                    cell = ws.cell(row=r, column=col_idx)
                    cell.font = Font(name="Calibri", size=11, bold=True)
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            # Populate Data Rows
            start_row = 5
            for row_idx, data in report_df.iterrows():
                r = start_row + row_idx
                ws.cell(row=r, column=1, value=data["Office Name"])
                ws.cell(row=r, column=2, value=data["Cash (Cnt)"])
                ws.cell(row=r, column=3, value=data["DQR Scan (Cnt)"])
                ws.cell(row=r, column=4, value=data["SBI POS-CARD (Cnt)"])
                ws.cell(row=r, column=5, value=data["SBI POS BHARAT QR (Cnt)"])
                ws.cell(row=r, column=6, value=data["SBI E PAY UPI (Cnt)"])
                ws.cell(row=r, column=7, value=f"=SUM(C{r}:F{r})")
                ws.cell(row=r, column=8, value=f"=B{r}+G{r}")
                ws.cell(row=r, column=9, value=f"=IFERROR(G{r}/H{r}, 0)")
            
            end_row = start_row + len(report_df) - 1
            total_row = end_row + 1
            
            # Total Row Formulas
            ws.cell(row=total_row, column=1, value="Total")
            for col_idx, col_let in enumerate(["B", "C", "D", "E", "F", "G", "H"], start=2):
                ws.cell(row=total_row, column=col_idx, value=f"=SUM({col_let}5:{col_let}{end_row})")
            ws.cell(row=total_row, column=9, value=f"=IFERROR(G{total_row}/H{total_row}, 0)")
            
            # Formatting Data Rows & Total Row
            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )
            
            red_fill = PatternFill(start_color="F8696B", end_color="F8696B", fill_type="solid")
            orange_fill = PatternFill(start_color="FCBA76", end_color="FCBA76", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFEB84", end_color="FFEB84", fill_type="solid")
            green_fill = PatternFill(start_color="8AD090", end_color="8AD090", fill_type="solid")
            
            top10_threshold = report_df["Cash (Cnt)"].nlargest(min(10, len(report_df))).min()
            top10_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            top10_font = Font(name="Calibri", size=11, bold=True, color="9C0006")
            
            for r in range(5, total_row + 1):
                ws.row_dimensions[r].height = 20
                for c in range(1, 10):
                    cell = ws.cell(row=r, column=c)
                    cell.font = Font(name="Calibri", size=11)
                    cell.border = thin_border
                    
                    if c >= 2:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    
                    if c == 9:
                        cell.number_format = "0.00%"
                        if r <= end_row:
                            val = report_df.loc[r - 5, "% of Digital Transactions"]
                            if val < 0.25:
                                cell.fill = red_fill
                            elif val < 0.50:
                                cell.fill = orange_fill
                            elif val < 0.75:
                                cell.fill = yellow_fill
                            else:
                                cell.fill = green_fill
                    elif c in range(2, 9):
                        cell.number_format = "#,##0"
                
                if r <= end_row:
                    cash_val = report_df.loc[r - 5, "Cash (Cnt)"]
                    if cash_val >= top10_threshold and cash_val > 0:
                        ws.cell(row=r, column=1).fill = top10_fill
                        ws.cell(row=r, column=1).font = top10_font
                        ws.cell(row=r, column=2).fill = top10_fill
                        ws.cell(row=r, column=2).font = top10_font

            total_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
            double_bottom_border = Border(
                top=Side(style='thin', color='000000'),
                bottom=Side(style='double', color='000000'),
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000')
            )
            for c in range(1, 10):
                cell = ws.cell(row=total_row, column=c)
                cell.font = Font(name="Calibri", size=11, bold=True)
                cell.fill = total_fill
                cell.border = double_bottom_border

            ws.column_dimensions['A'].width = 30
            for c_let in ["B", "C", "D", "E", "F", "G", "H", "I"]:
                ws.column_dimensions[c_let].width = 11

            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            out_filename = f"63_Offices_Digital_Report_{(datetime.now() - timedelta(days=1)).strftime('%Y%m%d')}.xlsx"
            
            st.success("Report generated successfully!")
            st.download_button(
                label="📥 Download Formatted Excel Report",
                data=output_buffer,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
