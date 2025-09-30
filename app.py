import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time

# --- ค่าคงที่และฟังก์ชัน (เหมือนเดิม แต่เสถียรขึ้น) ---
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'Deduction', 'OT_Formatted']

def decimal_to_hhmm(decimal_hours):
    if not isinstance(decimal_hours, (int, float)) or decimal_hours < 0:
        return "00:00"
    hours = int(decimal_hours)
    minutes = int(round((decimal_hours - hours) * 60))
    return f"{hours:02d}:{minutes:02d}"

def calculate_ot(time_in, time_out, day_type, deduction_time):
    try:
        if not all([isinstance(t, time) for t in [time_in, time_out]]) or not day_type:
            return 0.0

        dummy_date = datetime.now().date()
        dt_in = datetime.combine(dummy_date, time_in)
        dt_out = datetime.combine(dummy_date, time_out)
        
        if dt_out <= dt_in: dt_out += timedelta(days=1)

        total_duration = dt_out - dt_in
        ot_hours_decimal = 0.0
        
        if day_type == 'Weekday':
            ot_start_time = dt_in + timedelta(hours=9, minutes=30)
            if dt_out > ot_start_time:
                ot_hours_decimal = (dt_out - ot_start_time).total_seconds() / 3600
        
        elif day_type == 'Weekend':
            breaks = timedelta(hours=0)
            if total_duration > timedelta(hours=4): breaks += timedelta(hours=1)
            if total_duration > timedelta(hours=9): breaks += timedelta(minutes=30)
            work_duration = total_duration - breaks
            ot_hours_decimal = work_duration.total_seconds() / 3600
        
        deduction_decimal = 0.0
        if isinstance(deduction_time, time):
            deduction_decimal = deduction_time.hour + deduction_time.minute / 60.0
            
        final_ot = ot_hours_decimal - deduction_decimal
        return max(0.0, final_ot)
    except Exception:
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
    except Exception:
        return None

# --- ส่วนหน้าเว็บ Streamlit (ดีไซน์ใหม่) ---
st.set_page_config(layout="wide")
st.title("🚀 OT Calculator | โปรแกรมคำนวณโอที")

# --- 1. ส่วนเชื่อมต่อ ---
with st.container(border=True):
    sheet_url = st.text_input("🔗 วางลิงก์ Google Sheet ของคุณที่นี่")
    sheet_name = st.text_input("🏷️ ชื่อชีต (Sheet Name)", value="timesheet")
    if st.button("เชื่อมต่อ / รีเฟรชข้อมูล", type="primary"):
        with st.spinner("กำลังเชื่อมต่อ..."):
            worksheet = connect_to_gsheet(sheet_url, sheet_name)
            if worksheet:
                st.session_state.worksheet = worksheet
                all_values = worksheet.get_all_values()
                if len(all_values) > 1:
                    df = pd.DataFrame(all_values[1:], columns=all_values[0], dtype=str)
                else:
                    df = pd.DataFrame(columns=REQUIRED_COLUMNS, dtype=str)
                st.session_state.df = df
                st.success("เชื่อมต่อสำเร็จ!")
            else:
                st.error("เชื่อมต่อล้มเหลว! ตรวจสอบลิงก์และสิทธิ์การเข้าถึง")

if 'df' in st.session_state:
    # --- 2. ส่วนเพิ่มข้อมูล (ฟอร์ม) ---
    st.header("➕ เพิ่มรายการทำงานใหม่")
    with st.form("entry_form", clear_on_submit=True):
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            new_date = st.date_input("🗓️ วันที่")
        with col2:
            new_day_type = st.selectbox("✨ ประเภทวัน", ["Weekday", "Weekend"])
        with col3:
            new_time_in = st.time_input("🕘 เวลาเข้า", step=60)
        with col4:
            new_time_out = st.time_input("🕕 เวลาออก", step=60)
        with col5:
            new_deduction = st.time_input("✂️ หักเวลา", value=time(0,0), step=60)
        
        submitted = st.form_submit_button("คำนวณและเพิ่มลงตาราง")
        if submitted:
            ot_decimal = calculate_ot(new_time_in, new_time_out, new_day_type, new_deduction)
            ot_formatted = decimal_to_hhmm(ot_decimal)
            
            new_row = pd.DataFrame([{
                'Date': new_date.strftime('%Y-%m-%d'),
                'DayType': new_day_type,
                'TimeIn': new_time_in.strftime('%H:%M'),
                'TimeOut': new_time_out.strftime('%H:%M'),
                'Deduction': new_deduction.strftime('%H:%M'),
                'OT_Formatted': ot_formatted
            }])
            
            st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
            st.success(f"เพิ่มรายการสำเร็จ! คำนวณ OT ได้: {ot_formatted} ชั่วโมง")

    st.markdown("---")

    # --- 3. ส่วนแสดงผลและบันทึก ---
    st.header("📝 ตารางเวลาทำงาน")
    st.dataframe(st.session_state.df, use_container_width=True)
    
    if not st.session_state.df.empty:
        if st.button("💾 บันทึกข้อมูลทั้งหมดลง Google Sheet", type="primary"):
            with st.spinner("กำลังบันทึก..."):
                try:
                    # เรียงคอลัมน์ให้ตรงกับ REQUIRED_COLUMNS เสมอ
                    df_to_save = st.session_state.df.reindex(columns=REQUIRED_COLUMNS)
                    df_to_save.fillna('', inplace=True)
                    st.session_state.worksheet.clear()
                    # เขียน header + ข้อมูลทั้งหมด
                    set_with_dataframe(st.session_state.worksheet, df_to_save, include_index=False, allow_formulas=False)
                    st.success("บันทึกข้อมูลทั้งหมดเรียบร้อย!")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการบันทึก: {e}")
