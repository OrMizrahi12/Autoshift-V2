import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import datetime

st.set_page_config(page_title="עיבוד שבועיות", page_icon="📅", layout="wide")

st.title("📅 רכז סידורי עבודה לטבלה")
st.markdown("""
כאן תוכל להעלות מספר טבלאות של סידור עבודה, והמערכת תעבד את כולן ותפיק טבלה אחת גדולה ומסודרת.
המערכת מעבדת כל קובץ בנפרד כדי להבטיח שום מידע לא הולך לאיבוד.
""")

# ============================================================
# Helper Functions
# ============================================================

SHIFT_HOURS = {
    'בוקר': '07:00 – 15:00',
    'צהריים': '15:00 – 23:00',
    'לילה': '23:00 – 07:00',
}

def clean_name_generic(name):
    if not name or pd.isna(name): return ''
    s = str(name).strip()
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\s*-\s*\d+\s*$', '', s)
    s = re.sub(r'\d{3,}', '', s)
    s = s.replace('*', '').replace('"', '').replace("'", "")
    return ' '.join(s.split()).strip()

def is_time_val(val):
    if val is None: return False
    if isinstance(val, (datetime.time, datetime.datetime)): return True
    s = str(val).strip()
    return bool(re.match(r'^\d{1,2}:\d{2}(:\d{2})?$', s))

def resolve_times(t1, t2, shift_name):
    if not t1 or not t2:
        st1 = str(t1 or "").strip()
        st2 = str(t2 or "").strip()
        if st1 == 'None': st1 = ""
        if st2 == 'None': st2 = ""
        return st1, st2
    
    try:
        def to_min(t_val):
            if isinstance(t_val, (datetime.time, datetime.datetime)):
                return t_val.hour * 60 + t_val.minute
            parts = str(t_val).strip().split(':')
            h, m = map(int, parts[:2])
            return h * 60 + m
        
        m1, m2 = to_min(t1), to_min(t2)
        
        def calc_score(start_m, end_m, shift):
            dur = (end_m - start_m) % 1440
            if dur == 0: dur = 1440
            
            expectations = {
                'בוקר': 11*60, 'צהריים': 19*60, 'לילה': 3*60, 'ערב': 20*60, 'אמצע': 14*60,
            }
            exp_mid = expectations.get(shift, 14*60)
            actual_mid = (start_m + dur/2) % 1440
            dist = min(abs(actual_mid - exp_mid), 1440 - abs(actual_mid - exp_mid))
            
            penalty = 0
            if dur > 900 or dur < 120: penalty = 2000 
            return dist + penalty

        if calc_score(m1, m2, shift_name) <= calc_score(m2, m1, shift_name):
            return format_t(t1), format_t(t2)
        else:
            return format_t(t2), format_t(t1)
    except:
        return str(t1)[:5], str(t2)[:5]

def format_t(v):
    if isinstance(v, (datetime.time, datetime.datetime)):
        return v.strftime('%H:%M')
    parts = str(v).strip().split(':')
    return f"{int(parts[0]):02d}:{int(parts[1]):02d}"

def parse_source_b(file, file_name=""):
    wb = openpyxl.load_workbook(file, data_only=True)
    file_records = []

    for sheet in wb.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        if not rows: continue

        # Identify Header
        day_cols = {}
        header_idx = -1
        for r_idx, row in enumerate(rows[:20]):
            valid = {}
            for c_idx, cell in enumerate(row):
                if cell:
                    c_str = str(cell)
                    match = re.search(r'\d{1,2}/\d{1,2}/\d{2,4}', c_str)
                    if match: valid[c_idx] = match.group()
                    elif isinstance(cell, (datetime.datetime, datetime.date)):
                        valid[c_idx] = cell.strftime('%d/%m/%Y')
            if len(valid) >= 3:
                day_cols = valid
                header_idx = r_idx
                break
        
        if header_idx == -1: continue
        
        day_indices = sorted(day_cols.keys())
        main_pos = ""
        sub_pos = ""
        shift = ""
        
        for r_idx in range(header_idx + 1, len(rows)):
            row = rows[r_idx]
            col0_val = row[0]
            col0 = str(col0_val).strip() if col0_val is not None else ""
            if col0 == 'None': col0 = ""
            
            # Shift detection
            if col0 in ['בוקר', 'צהריים', 'לילה', 'אמצע', 'ערב']:
                shift = col0
                continue
            
            # Name detection
            names_found = {}
            for c in day_indices:
                if c < len(row):
                    val = row[c]
                    if val and str(val).strip() not in ['None', ''] and not is_time_val(val):
                        names_found[c] = str(val).strip()
            
            if names_found:
                if col0 and not is_time_val(col0_val) and not re.match(r'^\d+$', col0):
                    sub_pos = col0
                
                final_pos = sub_pos if sub_pos else main_pos
                if 'מתגבר' in final_pos: continue
                
                for c, raw_names in names_found.items():
                    # Find hours (check same row and next row primarily)
                    t1, t2 = None, None
                    for offset in [0, 1, -1, 2, -2]:
                        target_r = r_idx + offset
                        if 0 <= target_r < len(rows):
                            nr = rows[target_r]
                            # Look for Col, Col+1
                            if c < len(nr) and is_time_val(nr[c]):
                                t1 = nr[c]
                                if c+1 < len(nr) and is_time_val(nr[c+1]): t2 = nr[c+1]
                                break
                    
                    entry, exit = resolve_times(t1, t2, shift)
                    
                    for n in re.split(r'[\n\r,]+', raw_names):
                        clean_n = clean_name_generic(n)
                        if clean_n:
                            file_records.append({
                                'מקור': file_name,
                                'שם העובד': clean_n,
                                'תאריך': day_cols[c],
                                'עמדה': final_pos,
                                'שעת כניסה': entry,
                                'שעת יציאה': exit
                            })
            elif col0 and not is_time_val(col0_val) and not re.match(r'^\d+$', col0):
                main_pos = col0
                sub_pos = ""

    return pd.DataFrame(file_records)

# ============================================================
# Main UI
# ============================================================

uploaded_files = st.file_uploader("📂 העלה קבצי סידור עבודה", type=['xlsx', 'xls'], accept_multiple_files=True)

if st.button("🔄 התחל עיבוד", type="primary"):
    if uploaded_files:
        all_results = []
        progress_bar = st.progress(0)
        
        for i, f in enumerate(uploaded_files):
            try:
                st.write(f"🔄 מעבד קובץ: **{f.name}**...")
                file_bytes = io.BytesIO(f.getvalue())
                df = parse_source_b(file_bytes, file_name=f.name)
                if not df.empty:
                    all_results.append(df)
                    st.success(f"✅ נמצאו {len(df)} רשומות בקובץ {f.name}")
                else:
                    st.warning(f"⚠️ לא נמצאו נתונים תקינים בקובץ {f.name}")
            except Exception as e:
                st.error(f"❌ שגיאה בקובץ {f.name}: {e}")
            progress_bar.progress((i + 1) / len(uploaded_files))
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            st.divider()
            st.subheader(f"📊 סה\"כ רשומות שרוכזו: {len(final_df)}")
            
            # Displaying the table
            st.dataframe(final_df, use_container_width=True)
            
            # Download link
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            st.download_button(
                label="📥 הורד קובץ מאוחד (Excel)",
                data=output.getvalue(),
                file_name=f"consolidated_schedule_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("נא להעלות קבצים תחילה.")
