import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time

# --- ค่าคงที่และฟังก์ชันใหม่ ---

# [ใหม่] เพิ่มคอลัมน์ Deduction
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'Deduction', 'OT_Formatted']

# [ใหม่] ฟังก์ชันแปลงเวลา HH:MM เป็นทศนิยม (เช่น "01:30" -> 1.5)
def hhmm_to_decimal(time_str):
    """แปลง string HH:MM เป็น sốทศนิยม"""
    try:
        h, m = map(int, str(time_str).split(':'))
        return h + m / 60.0
    except (ValueError, AttributeError):
        return 0

# [ใหม่] ฟังก์ชันแปลงทศนิยมเป็นเวลา HH:MM (เช่น 1.5 -> "01:30")
def decimal_to_hhmm(decimal_hours):
    """แปลง sốทศนิยมเป็น string HH:MM"""
    if decimal_hours < 0:
        decimal_hours = 0
    hours = int(decimal_hours)
    minutes = int(round((decimal_hours - hours) * 60))
    return f"{hours:02d}:{minutes:02d}"

# --- ฟังก์ชันเดิมที่ปรับปรุงแล้ว ---

def prepare_dataframe(df):
    """แปลงชนิดข้อมูลใน DataFrame ให้ถูกต้องสำหรับ st.data_editor"""
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # [ใหม่] แปลง TimeIn/TimeOut เป็น time objects
    df['TimeIn'] = pd.to_datetime(df['TimeIn'], format='%H:%M', errors='coerce').dt.time
    df['TimeOut'] = pd.to_datetime(df['TimeOut'], format='%H:%M', errors='coerce').dt.time

    # จัดการคอลัมน์ข้อความ
    str_columns = ['DayType', 'Deduction', 'OT_Formatted']
    for col in str_columns:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str)
            
    return df

def calculate_ot(row):
    """คำนวณ OT เป็นทศนิยม และจัดการเวลาหักออก"""
    try:
        time_in, time_out, day_type = row.get('TimeIn'), row.get('TimeOut'), row.get('DayType')

        # [ใหม่] ดึงค่าเวลาหักออก
        deduction_str = row.get('Deduction', '00:00')

        if not all([time_in, time_out, day_type]):
            return 0

        # แปลง time object เป็น datetime เพื่อคำนวณ
        dummy_date = datetime.now().date()
        dt_in = datetime.combine(dummy_date, time_in)
        dt_out = datetime.combine(dummy_date, time_out)
        
        if dt_out < dt_in:
            dt_out += timedelta(days=1)

        total_duration = dt_out - dt_in
        ot_hours_decimal = 0
        
        if day_type == 'Weekday':
            actual_end_shift = dt_in + timedelta(hours=9)
            ot_start_time = actual_end_shift + timedelta(minutes=30)
            
            if dt_out > ot_start_time:
                ot_duration = dt_out - ot_start_time
                ot_hours_decimal = ot_duration.total_seconds() / 3600
                
        elif day_type == 'Weekend':
            work_duration = total_duration
            if work_duration > timedelta(hours=4):
                 work_duration -= timedelta(hours=1)
            if total_duration > timedelta(hours=9):
                 work_duration -= timedelta(minutes=30)
            ot_hours_decimal = work_duration.total_seconds() / 3600
        
        # [ใหม่] หักเวลาเพิ่มเติม
        deduction_decimal = hhmm_to_decimal(deduction_str)
        final_ot = ot_hours_decimal - deduction_decimal
        
        return max(0, final_ot) # คืนค่าเป็นทศนิยม

    except (ValueError, TypeError, AttributeError):
        return 0

def setup_sheet(worksheet):
    try:
        headers = worksheet.row_values(1)
    except gspread.exceptions.APIError:
        headers = []

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in headers]

    if missing_columns:
        st.info(f"ตรวจพบคอลัมน์ที่ขาดไป กำลังสร้าง: {', '.join(missing_columns)}")
        start_col_index = len(headers) + 1
        cell_list = [gspread.Cell(1, start_col_index + i, value=col_name) for i, col_name in enumerate(missing_columns)]
        worksheet.update_cells(cell_list)
        st.success("สร้างคอลัมน์ที่จำเป็นเรียบร้อยแล้ว!")
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
                    df_from_sheet = pd.DataFrame(all_data)
                    st.session_state.df = pd.DataFrame(columns=REQUIRED_COLUMNS)
                    if not df_from_sheet.empty:
                        st.session_state.df = pd.concat([st.session_state.df, df_from_sheet], ignore_index=True)
                    st.session_state.df = st.session_state.df.reindex(columns=REQUIRED_COLUMNS)
                    st.session_state.df = prepare_dataframe(st.session_state.df)
                    st.success("ดึงข้อมูลสำเร็จ!")

if st.session_state.df is not None:
    st.header("📝 ตารางเวลาทำงาน")
    st.caption("✨ **คำแนะนำ:** **ดับเบิลคลิก** ที่ช่องวันที่/เวลาเพื่อเปิดตัวเลือก | กรอกเวลาที่ต้องการหักเพิ่มในช่อง 'หักเวลา' (รูปแบบ HH:MM)")

    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic",
        column_config={
            "Date": st.column_config.DateColumn("🗓️ วันที่", format="YYYY-MM-DD", required=True),
            "DayType": st.column_config.SelectboxColumn("✨ ประเภทวัน", options=["Weekday", "Weekend"], required=True),
            # [ใหม่] เปลี่ยนเป็น TimeColumn
            "TimeIn": st.column_config.TimeColumn("🕘 เวลาเข้า", format="HH:mm", required=True),
            "TimeOut": st.column_config.TimeColumn("🕕 เวลาออก", format="HH:mm", required=True),
            # [ใหม่] เพิ่มคอลัมน์ Deduction
            "Deduction": st.column_config.TextColumn("✂️ หักเวลา (HH:MM)"),
            # [ใหม่] เปลี่ยนชื่อและประเภทคอลัมน์ OT
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
                # [ใหม่] คำนวณค่าทศนิยมก่อน แล้วค่อยแปลงเป็น HH:MM
                ot_decimal_values = df_to_process.apply(calculate_ot, axis=1)
                df_to_process['OT_Formatted'] = ot_decimal_values.apply(decimal_to_hhmm)
                st.session_state.df = df_to_process
                st.rerun()

    with col2:
        if st.button("💾 บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=True):
            if st.session_state.worksheet:
                with st.spinner("กำลังบันทึก..."):
                    df_to_save = edited_df.copy()
                    
                    # [ใหม่] แปลง Time objects กลับเป็น string ก่อนบันทึก
                    for col in ['TimeIn', 'TimeOut']:
                        df_to_save[col] = df_to_save[col].apply(lambda t: t.strftime('%H:%M') if isinstance(t, time) else t)
                    
                    df_to_save['Date'] = pd.to_datetime(df_to_save['Date']).dt.strftime('%Y-%m-%d')
                    df_to_save.fillna('', inplace=True)
                    
                    st.session_state.worksheet.clear()
                    set_with_dataframe(st.session_state.worksheet, df_to_save, include_index=False, allow_formulas=False)
                    st.success("บันทึกข้อมูลเรียบร้อย!")
