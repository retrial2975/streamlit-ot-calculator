import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta, time
import numpy as np

# --- ค่าคงที่และฟังก์ชัน ---
REQUIRED_COLUMNS = ['Date', 'DayType', 'TimeIn', 'TimeOut', 'Deduction', 'OT_Formatted', 'Note']

def prepare_dataframe(df):
    clean_df = pd.DataFrame()
    for col in REQUIRED_COLUMNS:
        source_series = df.get(col, pd.Series(dtype='object'))
        if col == 'Date':
            clean_df[col] = pd.to_datetime(source_series, errors='coerce')
        elif col in ['TimeIn', 'TimeOut', 'Deduction']:
            def to_time_obj(x):
                if isinstance(x, time): return x
                if pd.isna(x) or str(x).strip() in ['', 'None', 'NaT', 'nan']: return None
                try: return datetime.strptime(str(x), '%H:%M').time()
                except (ValueError, TypeError): return None
            clean_df[col] = source_series.apply(to_time_obj)
        else:
            clean_df[col] = pd.Series(source_series, dtype=str).fillna('')
    return clean_df

def decimal_to_hhmm(decimal_hours):
    if not isinstance(decimal_hours, (int, float)) or decimal_hours < 0: return "00:00"
    hours = int(decimal_hours)
    minutes = int(round((decimal_hours - hours) * 60))
    return f"{hours:02d}:{minutes:02d}"

def calculate_ot(row):
    try:
        time_in, time_out, day_type = row.get('TimeIn'), row.get('TimeOut'), row.get('DayType')
        if not all(isinstance(t, time) for t in [time_in, time_out]) or not day_type: return 0.0
        
        dummy_date = datetime.now().date()
        dt_in, dt_out = datetime.combine(dummy_date, time_in), datetime.combine(dummy_date, time_out)
        if dt_out <= dt_in: dt_out += timedelta(days=1)
        
        ot_hours_decimal = 0.0
        
        if day_type == 'Weekday':
            standard_start_time = datetime.combine(dummy_date, time(9, 0))
            calculation_base_time = max(dt_in, standard_start_time)
            ot_start_time = calculation_base_time + timedelta(hours=9, minutes=30)
            if dt_out > ot_start_time: ot_hours_decimal = (dt_out - ot_start_time).total_seconds() / 3600
        
        elif day_type == 'Weekend':
            total_duration = dt_out - dt_in
            breaks = timedelta(hours=0)
            if total_duration > timedelta(hours=4) and time_in < time(13, 0): breaks += timedelta(hours=1)
            if total_duration > timedelta(hours=9): breaks += timedelta(minutes=30)
            ot_hours_decimal = (total_duration - breaks).total_seconds() / 3600
        
        deduction_time = row.get('Deduction')
        deduction_decimal = 0.0
        if isinstance(deduction_time, time):
            deduction_decimal = deduction_time.hour + deduction_time.minute / 60.0
        
        return max(0.0, ot_hours_decimal - deduction_decimal)
    except Exception: 
        return 0.0

def setup_sheet(worksheet):
    headers = worksheet.row_values(1)
    if not headers:
        worksheet.update('A1', [REQUIRED_COLUMNS])
        headers = REQUIRED_COLUMNS

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
        
        worksheet = setup_sheet(worksheet)
        all_records = worksheet.get_all_records()
        source_df = pd.DataFrame(all_records)
        
        if source_df.empty:
            today = datetime.now()
            default_row = {
                'Date': today.date(),
                'DayType': 'Weekend' if today.weekday() > 4 else 'Weekday',
                'TimeIn': time(9, 0),
                'TimeOut': time(18, 0),
                'Deduction': time(18, 0),
                'OT_Formatted': None,
                'Note': 'แถวเริ่มต้นอัตโนมัติ'
            }
            st.session_state.df = pd.DataFrame([default_row])
        else:
            st.session_state.df = prepare_dataframe(source_df)
            
        return worksheet
    except Exception as e:
        st.error(f"การเชื่อมต่อล้มเหลว: {e}")
        st.session_state.df = None
        return None

# --- ส่วนหน้าเว็บ Streamlit ---
st.set_page_config(layout="wide")
st.title("🚀 OT Calculator | โปรแกรมคำนวณโอที")

if 'df' not in st.session_state: st.session_state.df = None
if 'worksheet' not in st.session_state: st.session_state.worksheet = None

with st.expander("📖 คลิกเพื่อดูวิธีการตั้งค่าและหลักการคำนวณ"):
    st.subheader("การตั้งค่า Google Sheet")
    st.write("""
        เพื่อให้โปรแกรมนี้สามารถอ่านและบันทึกข้อมูลใน Google Sheet ของคุณได้ คุณจำเป็นต้องให้สิทธิ์การเข้าถึง โดยทำตามขั้นตอนต่อไปนี้:
        1.  ไปที่ไฟล์ Google Sheet ที่คุณต้องการใช้
        2.  คลิกที่ปุ่ม **"Share"** (แชร์) ที่มุมขวาบน
        3.  ในช่อง "Add people and groups" ให้คัดลอกอีเมลนี้ไปวาง:
    """)
    st.code("streamlit-sheets-bot@streamlit-sheets-bot-473709.iam.gserviceaccount.com")
    st.write("""
        4.  ตั้งค่าสิทธิ์ให้เป็น **"Editor"** (ผู้มีสิทธิ์แก้ไข)
        5.  **ไม่ต้องติ๊ก** ช่อง "Notify people"
        6.  คลิก **"Share"**
    """)
    st.info("หลังจากให้สิทธิ์แล้ว คุณสามารถคัดลอกลิงก์ของ Google Sheet มาวางในช่องด้านล่างเพื่อเริ่มใช้งานได้ทันที")

    st.subheader("หลักการคำนวณ OT")
    st.markdown("""
    **สำหรับวันธรรมดา (Weekday):**
    - เวลาทำงานปกติคือ 8 ชั่วโมง + พักกลางวัน 1 ชั่วโมง (รวมเป็น 9 ชั่วโมงในที่ทำงาน)
    - **OT จะเริ่มนับหลังจากทำงานครบ 9 ชั่วโมง และพักอีก 30 นาที**
    - พูดง่ายๆ คือ OT จะเริ่มนับตั้งแต่ **`เวลาเข้างาน + 9 ชั่วโมง 30 นาที`** เป็นต้นไป
    - **เงื่อนไขเพิ่มเติม:**
        - หากเข้างานก่อน 9:00 น. จะเริ่มนับเวลาทำงานตอน 9:00 น.
        - หากเข้างานหลัง 9:00 น. (มาสาย) จะนับเวลาทำงานตามเวลาที่เข้างานจริง
    
    *ตัวอย่าง:*
    - เข้างาน `08:50` ออกงาน `19:45` → เวลาเริ่มนับ OT คือ `09:00 + 9:30` = `18:30` → ได้ OT `19:45 - 18:30` = **1 ชั่วโมง 15 นาที**
    - เข้างาน `09:30` ออกงาน `20:15` → เวลาเริ่มนับ OT คือ `09:30 + 9:30` = `19:00` → ได้ OT `20:15 - 19:00` = **1 ชั่วโมง 15 นาที**

    **สำหรับวันหยุด (Weekend):**
    - OT จะคิดจากระยะเวลาที่อยู่ในที่ทำงานทั้งหมด และหักเวลาพักตามเงื่อนไข
    - **เงื่อนไขการหักเวลาพัก:**
        - ทำงานเกิน 4 ชั่วโมง **และ** เข้างานก่อน 13:00 น. → **หัก 1 ชั่วโมง**
        - ทำงานเกิน 9 ชั่วโมง → **หักเพิ่มอีก 30 นาที**

    *ตัวอย่าง:*
    - เข้างาน `09:00` ออกงาน `15:00` (ทำงาน 6 ชม.) → เข้าเงื่อนไขเกิน 4 ชม. และเข้าก่อน 13:00 น. → ถูกหัก 1 ชม. → ได้ OT **5 ชั่วโมง**
    - เข้างาน `13:00` ออกงาน `19:00` (ทำงาน 6 ชม.) → **ไม่เข้าเงื่อนไข**พักกลางวัน (เพราะมาหลัง 13:00 น.) → ได้ OT **6 ชั่วโมง**
    - เข้างาน `09:00` ออกงาน `19:30` (ทำงาน 10.5 ชม.) → เข้าเงื่อนไขทั้ง 2 ข้อ → ถูกหัก 1 ชม. 30 นาที → ได้ OT **9 ชั่วโมง**
    """)

    st.subheader("หลักการคำนวณรายรับ OT")
    st.markdown("""
    โปรแกรมจะคำนวณรายรับ OT โดยประมาณจาก **"เงินเดือน"** ที่คุณกรอก โดยใช้สูตรพื้นฐานตามกฎหมายแรงงาน:
    1.  **หาค่าแรงต่อชั่วโมง:** `(เงินเดือน / 30 วัน / 8 ชั่วโมง)`
    2.  **คำนวณค่า OT ต่อชั่วโมง:** `ค่าแรงต่อชั่วโมง * 1.5` (สำหรับวันธรรมดา)
    3.  **คำนวณค่า OT ต่อนาที:** `ค่า OT ต่อชั่วโมง / 60`

    *หมายเหตุ: โปรแกรมจะทำการปัดเศษลง (Round Down) ในทุกขั้นตอนเพื่อให้ได้ค่าประมาณที่ใกล้เคียงที่สุด*
    
    **สูตรที่ใช้ในโปรแกรม:**
    - `เงินตามชั่วโมง = ROUNDDOWN(เงินเดือน / 30 / 8 * 1.5)`
    - `เงินตามนาที = ROUNDDOWN(เงินตามชั่วโมง / 60)`
    - `รายรับ OT ทั้งหมด = (OT รวม (ชม.) * เงินตามชั่วโมง) + (OT รวม (นาที) * เงินตามนาที)`
    """)

with st.container(border=True):
    sheet_url = st.text_input("🔗 วางลิงก์ Google Sheet ของคุณที่นี่")
    sheet_name = st.text_input("🏷️ ชื่อชีต (Sheet Name)", value="timesheet")
    if st.button("เชื่อมต่อ / รีเฟรชข้อมูล", type="primary"):
        with st.spinner("กำลังเชื่อมต่อ..."):
            st.session_state.worksheet = connect_to_gsheet(sheet_url, sheet_name)
            if st.session_state.df is not None:
                st.session_state.df = prepare_dataframe(st.session_state.df)
                st.success("เชื่อมต่อสำเร็จ!")

if st.session_state.df is not None:
    st.header("📝 ตารางเวลาทำงาน")
    st.caption("คุณสามารถ **แก้ไขข้อมูล** ในตารางได้โดยตรง | **ดับเบิลคลิก** ที่ช่องวันที่/เวลาเพื่อเปิดตัวเลือก")

    df_display = st.session_state.df.copy()
    df_display['Delete'] = False
    display_columns = ['Delete'] + REQUIRED_COLUMNS
    
    edited_df = st.data_editor(
        df_display[display_columns],
        key="main_data_editor", num_rows="dynamic",
        column_config={
            "Delete": st.column_config.CheckboxColumn("ลบ", default=False),
            "Date": st.column_config.DateColumn("🗓️ วันที่", format="YYYY-MM-DD", required=True),
            "DayType": st.column_config.SelectboxColumn("✨ ประเภทวัน", options=["Weekday", "Weekend"], required=True),
            "TimeIn": st.column_config.TimeColumn("🕘 เวลาเข้า", format="HH:mm", required=True, step=60),
            "TimeOut": st.column_config.TimeColumn("🕕 เวลาออก", format="HH:mm", required=True, step=60),
            "Deduction": st.column_config.TimeColumn("✂️ หักเวลา", format="HH:mm", step=60),
            "OT_Formatted": st.column_config.TextColumn("💰 OT (ชม.:นาที)", disabled=True),
            "Note": st.column_config.TextColumn("📝 หมายเหตุ"),
        },
        use_container_width=True, disabled=['OT_Formatted'])

    st.markdown("---")
    
    st.header("📊 สรุปผลและคำนวณรายรับ")
    def hhmm_to_decimal(t_str):
        try:
            h, m = map(int, t_str.split(':'))
            return h + m / 60
        except: return 0
    total_ot_decimal = edited_df['OT_Formatted'].apply(hhmm_to_decimal).sum()
    total_ot_hours, total_ot_minutes = int(total_ot_decimal), int((total_ot_decimal - int(total_ot_decimal)) * 60)
    col_summary, col_salary = st.columns(2)
    with col_summary:
        st.metric(label="ชั่วโมง OT ทั้งหมด", value=f"{total_ot_hours} ชั่วโมง {total_ot_minutes} นาที")
    with col_salary:
        salary = st.number_input("💵 กรอกเงินเดือน (บาท)", min_value=0, value=31000)
        if salary > 0:
            rate_per_hour = np.floor(salary / 30 / 8 * 1.5)
            rate_per_minute = np.floor(rate_per_hour / 60)
            ot_income = (total_ot_hours * rate_per_hour) + (total_ot_minutes * rate_per_minute)
            st.metric(label="รายรับ OT โดยประมาณ", value=f"฿ {ot_income:,.2f}")

    st.markdown("---")
    st.header("⚙️ เครื่องมือจัดการ")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("🗑️ ลบแถวที่เลือก", use_container_width=True):
            rows_to_delete = edited_df[edited_df['Delete'] == True].index
            df_after_delete = edited_df.drop(rows_to_delete)
            st.session_state.df = prepare_dataframe(df_after_delete)
            st.rerun()
    with col2:
        if st.button("📅 เรียงตามวันที่", use_container_width=True):
            df_sorted = edited_df.sort_values(by="Date", ascending=True)
            st.session_state.df = prepare_dataframe(df_sorted)
            st.rerun()
    with col3:
        if st.button("🮔 คำนวณ OT ทั้งหมด", use_container_width=True):
            df_prepared_for_calc = prepare_dataframe(edited_df)
            ot_decimal_values = df_prepared_for_calc.apply(calculate_ot, axis=1)
            df_to_process = edited_df.copy()
            df_to_process['OT_Formatted'] = ot_decimal_values.apply(decimal_to_hhmm)
            st.session_state.df = prepare_dataframe(df_to_process)
            st.rerun()
    with col4:
        if st.button("💾 บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=True):
            with st.spinner("กำลังบันทึก..."):
                df_to_save = edited_df.drop(columns=['Delete'])
                df_to_save = prepare_dataframe(df_to_save)
                
                df_to_save.dropna(subset=['Date', 'DayType', 'TimeIn', 'TimeOut'], how='any', inplace=True)

                def format_date_for_save(d):
                    if pd.isna(d): return ""
                    return d.strftime('%Y-%m-%d')
                
                df_to_save['Date'] = df_to_save['Date'].apply(format_date_for_save)
                for col in ['TimeIn', 'TimeOut', 'Deduction']:
                    df_to_save[col] = df_to_save[col].apply(lambda t: t.strftime('%H:%M') if isinstance(t, time) else "")
                
                df_to_save.fillna('', inplace=True)
                
                st.session_state.worksheet.clear()
                set_with_dataframe(st.session_state.worksheet, df_to_save, include_index=False, allow_formulas=False)
                st.success("บันทึกข้อมูลเรียบร้อย!")
