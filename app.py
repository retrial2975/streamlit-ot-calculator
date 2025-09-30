import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time
import numpy as np

# --- ค่าคงที่และฟังก์ชัน ---
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'Deduction', 'OT_Formatted', 'Note']

def decimal_to_hhmm(decimal_hours):
    if not isinstance(decimal_hours, (int, float)) or decimal_hours < 0: return "00:00"
    hours = int(decimal_hours)
    minutes = int(round((decimal_hours - hours) * 60))
    return f"{hours:02d}:{minutes:02d}"

def calculate_ot(row):
    try:
        time_in_str, time_out_str, day_type = row.get('TimeIn'), row.get('TimeOut'), row.get('DayType')
        if not all([time_in_str, time_out_str, day_type]): return 0.0
        
        time_in = datetime.strptime(time_in_str, '%H:%M').time()
        time_out = datetime.strptime(time_out_str, '%H:%M').time()
        
        dummy_date = datetime.now().date()
        dt_in, dt_out = datetime.combine(dummy_date, time_in), datetime.combine(dummy_date, time_out)
        if dt_out <= dt_in: dt_out += timedelta(days=1)
        
        ot_hours_decimal = 0.0
        if day_type == 'Weekday':
            standard_start_time = datetime.combine(dummy_date, time(9, 0))
            calculation_base_time = max(dt_in, standard_start_time)
            ot_start_time = calculation_base_time + timedelta(hours=9, minutes=30)
            if dt_out > ot_start_time: 
                ot_hours_decimal = (dt_out - ot_start_time).total_seconds() / 3600
        elif day_type == 'Weekend':
            total_duration = dt_out - dt_in
            breaks = timedelta(hours=0)
            if total_duration > timedelta(hours=4): breaks += timedelta(hours=1)
            if total_duration > timedelta(hours=9): breaks += timedelta(minutes=30)
            ot_hours_decimal = (total_duration - breaks).total_seconds() / 3600
        
        deduction_str = row.get('Deduction')
        deduction_decimal = 0.0
        if deduction_str:
            h, m = map(int, deduction_str.split(':'))
            deduction_decimal = h + m / 60.0
            
        return max(0.0, ot_hours_decimal - deduction_decimal)
    except (ValueError, TypeError, AttributeError): 
        return 0.0

def setup_sheet(worksheet):
    try:
        headers = worksheet.row_values(1)
    except gspread.exceptions.APIError: headers = []
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in headers]
    if missing_columns:
        start_col_index = len(headers) + 1
        cell_list = [gspread.Cell(1, start_col_index + i, value=col_name) for i, col_name in enumerate(missing_columns)]
        worksheet.update_cells(cell_list)
    return worksheet

def connect_to_gsheet(sheet_url, sheet_name):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(sheet_url)
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        return setup_sheet(worksheet)
    except Exception: return None

# --- ส่วนหน้าเว็บ Streamlit ---
st.set_page_config(layout="wide")
st.title("🚀 OT Calculator | โปรแกรมคำนวณโอที")

if 'df' not in st.session_state: st.session_state.df = None
if 'worksheet' not in st.session_state: st.session_state.worksheet = None

with st.container(border=True):
    sheet_url = st.text_input("🔗 วางลิงก์ Google Sheet ของคุณที่นี่")
    sheet_name = st.text_input("🏷️ ชื่อชีต (Sheet Name)", value="timesheet")
    if st.button("เชื่อมต่อ / รีเฟรชข้อมูล", type="primary"):
        with st.spinner("กำลังเชื่อมต่อ..."):
            st.session_state.worksheet = connect_to_gsheet(sheet_url, sheet_name)
            if st.session_state.worksheet:
                all_values = st.session_state.worksheet.get_all_values()
                if len(all_values) > 1:
                    headers, data_rows = all_values[0], all_values[1:]
                    df = pd.DataFrame(data_rows, columns=headers, dtype=str)
                else:
                    df = pd.DataFrame(columns=REQUIRED_COLUMNS, dtype=str)

                # --- [CRITICAL FIX] ---
                # 1. ทำให้แน่ใจว่าทุกคอลัมน์ที่ต้องการมีอยู่ และเป็น string
                for col in REQUIRED_COLUMNS:
                    if col not in df.columns:
                        df[col] = ''
                df = df[REQUIRED_COLUMNS].fillna('')

                # 2. แปลงคอลัมน์ 'Date' ให้เป็นชนิด datetime ที่ถูกต้อง
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                # ----------------------

                st.session_state.df = df
                st.success("เชื่อมต่อสำเร็จ!")

if st.session_state.df is not None:
    st.header("📝 ตารางเวลาทำงาน")
    st.caption("กรอกเวลาในรูปแบบ **HH:MM** (เช่น 09:30 หรือ 22:50)")

    edited_df = st.data_editor(
        st.session_state.df,
        key="main_data_editor", num_rows="dynamic",
        column_config={
            "Date": st.column_config.DateColumn("🗓️ วันที่", format="YYYY-MM-DD", required=True),
            "DayType": st.column_config.SelectboxColumn("✨ ประเภทวัน", options=["Weekday", "Weekend"], required=True),
            "TimeIn": st.column_config.TextColumn("🕘 เวลาเข้า (HH:MM)", required=True),
            "TimeOut": st.column_config.TextColumn("🕕 เวลาออก (HH:MM)", required=True),
            "Deduction": st.column_config.TextColumn("✂️ หักเวลา (HH:MM)"),
            "OT_Formatted": st.column_config.TextColumn("💰 OT (ชม.:นาที)", disabled=True),
            "Note": st.column_config.TextColumn("📝 หมายเหตุ"),
        },
        use_container_width=True, disabled=['OT_Formatted'])
    
    # ... (ส่วนสรุปผลและคำนวณเงิน ยังไม่ใส่เพื่อลดความซับซ้อน) ...

    st.markdown("---")
    st.header("⚙️ เครื่องมือจัดการ")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📅 เรียงตามวันที่", use_container_width=True):
            # ต้องจัดการกับแถวที่มีวันที่ว่าง (NaT) ก่อนเรียง
            df_to_sort = edited_df.copy()
            df_to_sort['Date'] = pd.to_datetime(df_to_sort['Date'], errors='coerce')
            df_sorted = df_to_sort.dropna(subset=['Date']).sort_values(by="Date", ascending=True)
            st.session_state.df = pd.concat([df_sorted, df_to_sort[df_to_sort['Date'].isnull()]]).reset_index(drop=True)
            st.rerun()
    with col2:
        if st.button("🮔 คำนวณ OT ทั้งหมด", use_container_width=True):
            df_to_process = edited_df.copy()
            ot_decimal_values = df_to_process.apply(calculate_ot, axis=1)
            df_to_process['OT_Formatted'] = ot_decimal_values.apply(decimal_to_hhmm)
            st.session_state.df = df_to_process
            st.rerun()
    with col3:
        if st.button("💾 บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=True):
            with st.spinner("กำลังบันทึก..."):
                df_to_save = edited_df.copy()
                # แปลง Date กลับเป็น string ก่อนบันทึก
                df_to_save['Date'] = pd.to_datetime(df_to_save['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
                df_to_save = df_to_save.reindex(columns=REQUIRED_COLUMNS).fillna('')

                st.session_state.worksheet.clear()
                set_with_dataframe(st.session_state.worksheet, df_to_save, include_index=False)
                st.success("บันทึกข้อมูลเรียบร้อย!")
