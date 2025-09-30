import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time

# --- ค่าคงที่และฟังก์ชัน ---
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'Deduction', 'OT_Formatted']

def decimal_to_hhmm(decimal_hours):
    if not isinstance(decimal_hours, (int, float)) or decimal_hours < 0:
        return "00:00"
    hours = int(decimal_hours)
    minutes = int(round((decimal_hours - hours) * 60))
    return f"{hours:02d}:{minutes:02d}"

def calculate_ot(row):
    try:
        time_in, time_out, day_type = row.get('TimeIn'), row.get('TimeOut'), row.get('DayType')
        deduction_time = row.get('Deduction')

        if not all(isinstance(t, time) for t in [time_in, time_out]) or not day_type:
            return 0.0

        dummy_date = datetime.now().date()
        dt_in = datetime.combine(dummy_date, time_in)
        dt_out = datetime.combine(dummy_date, time_out)
        
        if dt_out <= dt_in: dt_out += timedelta(days=1)

        ot_hours_decimal = 0.0
        
        if day_type == 'Weekday':
            actual_end_shift = dt_in + timedelta(hours=9)
            ot_start_time = actual_end_shift + timedelta(minutes=30)
            if dt_out > ot_start_time:
                ot_hours_decimal = (dt_out - ot_start_time).total_seconds() / 3600
        elif day_type == 'Weekend':
            work_duration = dt_out - dt_in
            if work_duration > timedelta(hours=4): work_duration -= timedelta(hours=1)
            if work_duration > timedelta(hours=9): work_duration -= timedelta(minutes=30)
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
        st.info(f"กำลังสร้างคอลัมน์ที่ขาดไป: {', '.join(missing_columns)}")
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
    except Exception as e:
        st.error(f"การเชื่อมต่อล้มเหลว: {e}")
        return None

# --- ส่วนหน้าเว็บ Streamlit ---
st.set_page_config(layout="wide")
st.title("🚀 OT Calculator | โปรแกรมคำนวณโอที")

if 'df' not in st.session_state: st.session_state.df = None
if 'worksheet' not in st.session_state: st.session_state.worksheet = None

with st.container(border=True):
    st.subheader("เชื่อมต่อ Google Sheet")
    sheet_url = st.text_input("🔗 วางลิงก์ Google Sheet ของคุณที่นี่")
    sheet_name = st.text_input("🏷️ ชื่อชีต (Sheet Name)", value="timesheet")
    if st.button("เชื่อมต่อและดึงข้อมูล", type="primary"):
        if sheet_url and sheet_name:
            with st.spinner("กำลังเชื่อมต่อ..."):
                st.session_state.worksheet = connect_to_gsheet(sheet_url, sheet_name)
                if st.session_state.worksheet:
                    all_data = st.session_state.worksheet.get_all_records()
                    source_df = pd.DataFrame(all_data)
                    
                    # --- [ใหม่] สร้าง DataFrame ด้วยวิธีที่ปลอดภัยและตรงไปตรงมา ---
                    clean_data = {}
                    
                    # จัดการคอลัมน์ Date
                    clean_data['Date'] = pd.to_datetime(source_df.get('Date'), errors='coerce')

                    # จัดการคอลัมน์ Time ทั้งหมด
                    for col in ['TimeIn', 'TimeOut', 'Deduction']:
                        s = pd.Series(source_df.get(col), dtype=str).replace({'NaT': '', 'None': ''})
                        clean_data[col] = pd.to_datetime(s, format='%H:%M', errors='coerce').dt.time
                    
                    # จัดการคอลัมน์ String ทั้งหมด
                    for col in ['DayType', 'OT_Formatted']:
                        clean_data[col] = pd.Series(source_df.get(col), dtype=str).fillna('')
                    
                    # สร้าง DataFrame ที่สะอาดแล้วจาก Dictionary
                    st.session_state.df = pd.DataFrame(clean_data, columns=REQUIRED_COLUMNS)
                    # ----------------------------------------------------------------

                    st.success("ดึงข้อมูลสำเร็จ!")

if st.session_state.df is not None:
    st.header("📝 ตารางเวลาทำงาน")
    st.caption("✨ **คำแนะนำ:** **ดับเบิลคลิก** ที่ช่องวันที่/เวลาเพื่อเปิดตัวเลือก | หากใช้ Brave Browser ให้ปิด Shields (ไอคอนสิงโต) ก่อน")

    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic",
        column_config={
            "Date": st.column_config.DateColumn("🗓️ วันที่", format="YYYY-MM-DD", required=True),
            "DayType": st.column_config.SelectboxColumn("✨ ประเภทวัน", options=["Weekday", "Weekend"], required=True),
            "TimeIn": st.column_config.TimeColumn("🕘 เวลาเข้า", format="HH:mm", required=True, step=60),
            "TimeOut": st.column_config.TimeColumn("🕕 เวลาออก", format="HH:mm", required=True, step=60),
            "Deduction": st.column_config.TimeColumn("✂️ หักเวลา", format="HH:mm", step=60),
            "OT_Formatted": st.column_config.TextColumn("💰 OT (ชั่วโมง:นาที)", disabled=True),
        },
        use_container_width=True,
        key="data_editor"
    )

    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🧮 คำนวณ OT ทั้งหมด", use_container_width=True):
            if not edited_df.empty:
                df_to_process = edited_df.copy()
                ot_decimal_values = df_to_process.apply(calculate_ot, axis=1)
                df_to_process['OT_Formatted'] = ot_decimal_values.apply(decimal_to_hhmm)
                st.session_state.df = df_to_process
                st.rerun()

    with col2:
        if st.button("💾 บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=True):
            if st.session_state.worksheet:
                with st.spinner("กำลังบันทึก..."):
                    df_to_save = edited_df.copy()
                    
                    for col in ['TimeIn', 'TimeOut', 'Deduction']:
                        df_to_save[col] = df_to_save[col].apply(lambda t: t.strftime('%H:%M') if isinstance(t, time) else "")
                    
                    df_to_save['Date'] = pd.to_datetime(df_to_save['Date']).dt.strftime('%Y-%m-%d')
                    df_to_save.fillna('', inplace=True)
                    
                    st.session_state.worksheet.clear()
                    set_with_dataframe(st.session_state.worksheet, df_to_save, include_index=False, allow_formulas=False)
                    st.success("บันทึกข้อมูลเรียบร้อย!")
