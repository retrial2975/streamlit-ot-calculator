import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time

# --- ฟังก์ชันการเชื่อมต่อและคำนวณ ---

# [ใหม่] รายชื่อคอลัมน์ที่โปรแกรมต้องการ
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'OT_Hours']

# [ใหม่] ฟังก์ชันสำหรับตรวจสอบและสร้างคอลัมน์
def setup_sheet(worksheet):
    """ตรวจสอบหัวคอลัมน์ในชีต ถ้าไม่มีจะสร้างให้โดยอัตโนมัติ"""
    try:
        # ดึงค่าแถวแรกสุด (หัวคอลัมน์)
        headers = worksheet.row_values(1)
    except gspread.exceptions.APIError:
        # กรณีชีตว่างเปล่ามากๆ ไม่มีข้อมูลเลย
        headers = []

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in headers]

    if missing_columns:
        st.info(f"ตรวจพบคอลัมน์ที่ขาดไป กำลังสร้าง: {', '.join(missing_columns)}")
        # หาคอลัมน์ว่างแรกสุดที่จะเริ่มเติม
        start_col_index = len(headers) + 1
        # สร้าง list ของ cell ที่จะอัปเดต
        cell_list = []
        for i, col_name in enumerate(missing_columns):
            cell_list.append(gspread.Cell(1, start_col_index + i, value=col_name))
        
        # ส่งคำสั่งอัปเดตทั้งหมดในครั้งเดียว
        worksheet.update_cells(cell_list)
        st.success("สร้างคอลัมน์ที่จำเป็นเรียบร้อยแล้ว!")
    return worksheet


# 1. ฟังก์ชันเชื่อมต่อกับ Google Sheets (ปรับปรุงเล็กน้อย)
def connect_to_gsheet(sheet_url, sheet_name):
    """เชื่อมต่อกับ Google Sheet และคืนค่า worksheet object"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(sheet_url)
        
        # ตรวจสอบว่ามีชีตชื่อนั้นอยู่หรือไม่ ถ้าไม่มีให้สร้างใหม่
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            st.info(f"ไม่พบชีตชื่อ '{sheet_name}', กำลังสร้างชีตใหม่...")
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="10")
            st.success(f"สร้างชีต '{sheet_name}' เรียบร้อยแล้ว")

        # [ใหม่] เรียกใช้ฟังก์ชัน setup_sheet
        return setup_sheet(worksheet)

    except Exception as e:
        st.error(f"การเชื่อมต่อล้มเหลว: {e}")
        return None

# 2. ฟังก์ชันคำนวณ OT (เหมือนเดิม)
def calculate_ot(row):
    """คำนวณ OT จากข้อมูลแถวนั้นๆ"""
    try:
        # แปลง string เวลาเป็น object datetime
        time_in_str = row['TimeIn']
        time_out_str = row['TimeOut']
        day_type = row['DayType']

        if not time_in_str or not time_out_str:
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
            std_start_time = datetime.combine(dummy_date, time(9, 0))
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

# --- ส่วนของหน้าเว็บ Streamlit (เหมือนเดิม) ---

st.set_page_config(layout="wide")
st.title("คำนวณโอที (OT Calculator) 🚀")

if 'df' not in st.session_state:
    st.session_state.df = None
if 'worksheet' not in st.session_state:
    st.session_state.worksheet = None

sheet_url = st.text_input("🔗 วางลิงก์ Google Sheet ของคุณที่นี่")
sheet_name = st.text_input("🏷️ ชื่อชีต (Sheet Name)", value="timesheet")

if st.button("เชื่อมต่อและดึงข้อมูล"):
    if sheet_url and sheet_name:
        with st.spinner("กำลังเชื่อมต่อและดึงข้อมูล..."):
            st.session_state.worksheet = connect_to_gsheet(sheet_url, sheet_name)
            if st.session_state.worksheet:
                # ดึงข้อมูลมาเฉพาะคอลัมน์ที่ต้องการ เพื่อป้องกัน error ถ้ามีคอลัมน์อื่นในชีต
                all_data = st.session_state.worksheet.get_all_records()
                df = pd.DataFrame(all_data)
                
                # สร้าง DataFrame เปล่าที่มีคอลัมน์ที่ถูกต้องก่อน
                st.session_state.df = pd.DataFrame(columns=REQUIRED_COLUMNS)
                if not df.empty:
                    # แล้วค่อยเติมข้อมูลจากชีตเข้ามา
                    st.session_state.df = pd.concat([st.session_state.df, df], ignore_index=True)
                
                # เติมค่าว่างสำหรับคอลัมน์ที่อาจจะไม่มีข้อมูล
                for col in REQUIRED_COLUMNS:
                    if col not in st.session_state.df.columns:
                        st.session_state.df[col] = ''
                
                st.session_state.df = st.session_state.df[REQUIRED_COLUMNS] # จัดลำดับคอลัมน์ให้ถูกต้อง
                st.success("ดึงข้อมูลสำเร็จ!")
            else:
                st.session_state.df = None
    else:
        st.warning("กรุณากรอกข้อมูลให้ครบถ้วน")

if st.session_state.df is not in None:
    st.header("ตารางเวลาทำงาน")
    
    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic",
        column_config={
            "Date": st.column_config.DateColumn("วันที่", format="YYYY-MM-DD"),
            "DayType": st.column_config.SelectboxColumn("ประเภทวัน", options=["Weekday", "Weekend"], required=True),
            "TimeIn": st.column_config.TextColumn("เวลาเข้า (HH:MM)"),
            "TimeOut": st.column_config.TextColumn("เวลาออก (HH:MM)"),
            "OT_Hours": st.column_config.NumberColumn("ชั่วโมง OT (คำนวณอัตโนมัติ)", disabled=True),
        },
        key="data_editor"
    )

    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("คำนวณ OT ทั้งหมด", use_container_width=True):
            if not edited_df.empty:
                edited_df['OT_Hours'] = edited_df.apply(calculate_ot, axis=1)
                st.session_state.df = edited_df
                st.rerun()

    with col2:
        if st.button("💾 บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=True):
            if st.session_state.worksheet:
                with st.spinner("กำลังบันทึก..."):
                    # แปลงคอลัมน์ Date กลับเป็น string ก่อนบันทึก
                    df_to_save = edited_df.copy()
                    # ตรวจสอบว่าคอลัมน์ Date เป็น datetime หรือไม่
                    if pd.api.types.is_datetime64_any_dtype(df_to_save['Date']):
                        df_to_save['Date'] = df_to_save['Date'].dt.strftime('%Y-%m-%d')
                    
                    # เติมค่าว่างด้วยสตริงเปล่า เพื่อป้องกัน NaN
                    df_to_save.fillna('', inplace=True)
                    
                    st.session_state.worksheet.clear()
                    set_with_dataframe(st.session_state.worksheet, df_to_save)
                    st.success("บันทึกข้อมูลเรียบร้อย!")
            else:
                st.error("ไม่สามารถบันทึกได้ กรุณาเชื่อมต่อ Google Sheet ก่อน")
