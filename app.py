import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time

# --- ค่าคงที่และฟังก์ชันหลัก ---

# รายชื่อคอลัมน์ที่โปรแกรมต้องการ
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'OT_Hours']

def prepare_dataframe(df):
    """แปลงชนิดข้อมูลใน DataFrame ให้ถูกต้องสำหรับ st.data_editor"""
    # 1. แปลงคอลัมน์ Date ให้เป็นชนิด datetime
    # errors='coerce' จะเปลี่ยนค่าที่แปลงไม่ได้ (เช่น text ว่าง) ให้เป็น NaT (Not a Time)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # 2. แปลง OT_Hours ให้เป็นชนิดตัวเลข
    df['OT_Hours'] = pd.to_numeric(df['OT_Hours'], errors='coerce').fillna(0)

    # 3. ทำให้คอลัมน์ที่เป็นข้อความไม่มีค่า Na หรือ NaN และเป็นชนิด string
    str_columns = ['DayType', 'TimeIn', 'TimeOut']
    for col in str_columns:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str)
            
    return df

def setup_sheet(worksheet):
    """ตรวจสอบหัวคอลัมน์ในชีต ถ้าไม่มีจะสร้างให้โดยอัตโนมัติ"""
    try:
        headers = worksheet.row_values(1)
    except gspread.exceptions.APIError:
        headers = [] # กรณีชีตว่างเปล่า

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in headers]

    if missing_columns:
        st.info(f"ตรวจพบคอลัมน์ที่ขาดไป กำลังสร้าง: {', '.join(missing_columns)}")
        start_col_index = len(headers) + 1
        cell_list = [gspread.Cell(1, start_col_index + i, value=col_name) for i, col_name in enumerate(missing_columns)]
        worksheet.update_cells(cell_list)
        st.success("สร้างคอลัมน์ที่จำเป็นเรียบร้อยแล้ว!")
    return worksheet

def connect_to_gsheet(sheet_url, sheet_name):
    """เชื่อมต่อกับ Google Sheet, สร้างชีต/คอลัมน์ถ้าจำเป็น"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(sheet_url)
        
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            st.info(f"ไม่พบชีตชื่อ '{sheet_name}', กำลังสร้างชีตใหม่...")
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="10")
            st.success(f"สร้างชีต '{sheet_name}' เรียบร้อยแล้ว")

        return setup_sheet(worksheet)

    except Exception as e:
        st.error(f"การเชื่อมต่อล้มเหลว: {e}")
        return None

def calculate_ot(row):
    """คำนวณ OT จากข้อมูลแต่ละแถว"""
    try:
        time_in_str, time_out_str, day_type = row.get('TimeIn'), row.get('TimeOut'), row.get('DayType')

        if not all([time_in_str, time_out_str, day_type]):
            return 0

        time_in = datetime.strptime(str(time_in_str), '%H:%M').time()
        time_out = datetime.strptime(str(time_out_str), '%H:%M').time()
        
        dummy_date = datetime.now().date()
        dt_in = datetime.combine(dummy_date, time_in)
        dt_out = datetime.combine(dummy_date, time_out)
        
        if dt_out < dt_in:
            dt_out += timedelta(days=1)

        total_duration = dt_out - dt_in
        ot_hours = 0
        
        if day_type == 'Weekday':
            actual_end_shift = dt_in + timedelta(hours=9)
            ot_start_time = actual_end_shift + timedelta(minutes=30)
            
            if dt_out > ot_start_time:
                ot_duration = dt_out - ot_start_time
                ot_hours = ot_duration.total_seconds() / 3600
                
        elif day_type == 'Weekend':
            work_duration = total_duration
            if work_duration > timedelta(hours=4):
                 work_duration -= timedelta(hours=1)
            if total_duration > timedelta(hours=9):
                 work_duration -= timedelta(minutes=30)
            ot_hours = work_duration.total_seconds() / 3600

        return round(ot_hours, 2) if ot_hours > 0 else 0

    except (ValueError, TypeError):
        return 0

# --- ส่วนหน้าเว็บ Streamlit ---

st.set_page_config(layout="wide")
st.title("🚀 OT Calculator | โปรแกรมคำนวณโอที")

# จัดการ State ของข้อมูล
if 'df' not in st.session_state:
    st.session_state.df = None
if 'worksheet' not in st.session_state:
    st.session_state.worksheet = None

# ส่วนรับ Input และเชื่อมต่อ
with st.container(border=True):
    st.subheader("เชื่อมต่อ Google Sheet")
    sheet_url = st.text_input("🔗 วางลิงก์ Google Sheet ของคุณที่นี่", key="sheet_url")
    sheet_name = st.text_input("🏷️ ชื่อชีต (Sheet Name)", value="timesheet", key="sheet_name")

    if st.button("เชื่อมต่อและดึงข้อมูล", type="primary"):
        if sheet_url and sheet_name:
            with st.spinner("กำลังเชื่อมต่อและดึงข้อมูล..."):
                st.session_state.worksheet = connect_to_gsheet(sheet_url, sheet_name)
                if st.session_state.worksheet:
                    all_data = st.session_state.worksheet.get_all_records()
                    df_from_sheet = pd.DataFrame(all_data)
                    
                    st.session_state.df = pd.DataFrame(columns=REQUIRED_COLUMNS)
                    if not df_from_sheet.empty:
                        st.session_state.df = pd.concat([st.session_state.df, df_from_sheet], ignore_index=True)
                    
                    for col in REQUIRED_COLUMNS:
                        if col not in st.session_state.df.columns:
                            st.session_state.df[col] = ''
                    
                    st.session_state.df = st.session_state.df[REQUIRED_COLUMNS]
                    
                    # แปลงชนิดข้อมูลให้ถูกต้องก่อนแสดงผล
                    st.session_state.df = prepare_dataframe(st.session_state.df)
                    
                    st.success("ดึงข้อมูลสำเร็จ!")
                else:
                    st.session_state.df = None
        else:
            st.warning("กรุณากรอกข้อมูลให้ครบถ้วน")

# ส่วนแสดงผลและแก้ไขข้อมูล
if st.session_state.df is not None:
    st.header("📝 ตารางเวลาทำงาน")
    st.caption("ดับเบิลคลิกที่ช่องวันที่เพื่อเปิดปฏิทิน, สามารถแก้ไขข้อมูลในตาราง, เพิ่ม/ลบแถวได้โดยตรง")

    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic",
        column_config={
            "Date": st.column_config.DateColumn("🗓️ วันที่", format="YYYY-MM-DD", required=True),
            "DayType": st.column_config.SelectboxColumn("✨ ประเภทวัน", options=["Weekday", "Weekend"], required=True),
            "TimeIn": st.column_config.TextColumn("🕘 เวลาเข้า (HH:MM)"),
            "TimeOut": st.column_config.TextColumn("🕕 เวลาออก (HH:MM)"),
            "OT_Hours": st.column_config.NumberColumn("💰 ชั่วโมง OT", disabled=True, help="ค่า OT จะถูกคำนวณอัตโนมัติ"),
        },
        use_container_width=True,
        key="data_editor"
    )

    st.markdown("---")
    
    # ส่วนปุ่มคำนวณและบันทึก
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🧮 คำนวณ OT ทั้งหมด", use_container_width=True):
            if not edited_df.empty:
                df_to_process = edited_df.copy()
                st.session_state.df = df_to_process
                st.session_state.df['OT_Hours'] = st.session_state.df.apply(calculate_ot, axis=1)
                st.rerun()

    with col2:
        if st.button("💾 บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=True):
            if st.session_state.worksheet:
                with st.spinner("กำลังบันทึกข้อมูล..."):
                    df_to_save = edited_df.copy()
                    
                    df_to_save.fillna('', inplace=True)
                    if 'Date' in df_to_save.columns:
                         # แปลงวันที่กลับเป็น string ใน format ที่ต้องการก่อนบันทึก
                        df_to_save['Date'] = pd.to_datetime(df_to_save['Date']).dt.strftime('%Y-%m-%d')
                    
                    st.session_state.worksheet.clear()
                    set_with_dataframe(st.session_state.worksheet, df_to_save, include_index=False)
                    st.success("บันทึกข้อมูลเรียบร้อย!")
            else:
                st.error("ไม่สามารถบันทึกได้ กรุณาเชื่อมต่อ Google Sheet ก่อน")
