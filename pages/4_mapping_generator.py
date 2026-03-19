import streamlit as st
import pandas as pd
import io
import re
import difflib

st.set_page_config(page_title="מחולל מיפויים", page_icon="🔗", layout="wide")

st.title("🔗 מחולל קבצי מיפוי (Master Mapping)")
st.markdown("""
כלי זה מאפשר לך להעלות את מסדי הנתונים המלאים של שתי המערכות (YLM ו-Tabit Shift),
ולזהות באופן אוטומטי את ההקשרים בין שמות העובדים והעמדות בשתיהן.
בסיום, תוכל להוריד קובץ **מיפוי קבוע** שיחסוך לך את תהליך ההתאמה הידני בדוח המשמרות!
""")

# ============================================================
# Core Functions (Imported/Copied to ensure stability)
# ============================================================
def clean_name_generic(name):
    if not name or pd.isna(name): return ''
    s = str(name).strip()
    s = re.sub(r'\s*-\s*\d+', '', s)
    s = re.sub(r'\d+', '', s)
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'[^א-תa-zA-Z\s]', '', s)
    return " ".join(s.split())

def normalize_final_letters(s):
    finals_map = {'ך': 'כ', 'ם': 'מ', 'ן': 'נ', 'ף': 'פ', 'ץ': 'צ'}
    for final, normal in finals_map.items():
        s = s.replace(final, normal)
    return s

def calc_similarity(s1, s2):
    if not s1 or not s2: return 0.0
    if s1 == s2: return 1.0
    n1 = normalize_final_letters(s1)
    n2 = normalize_final_letters(s2)
    if n1 == n2: return 1.0
    if n1 in n2 or n2 in n1: return 0.95
    t1, t2 = n1.split(), n2.split()
    set1, set2 = set(t1), set(t2)
    intersection = set1.intersection(set2)
    if intersection:
        intersect_score = len(intersection) / min(len(set1), len(set2))
        if intersect_score == 1.0: return 0.98
    else:
        intersect_score = 0.0
    st1 = " ".join(sorted(t1))
    st2 = " ".join(sorted(t2))
    sort_score = difflib.SequenceMatcher(None, st1, st2).ratio()
    orig_score = difflib.SequenceMatcher(None, s1, s2).ratio()
    return max(intersect_score * 0.9, sort_score, orig_score)

def parse_tabit_detailed_reports(files):
    import openpyxl
    names = set()
    positions = set()
    
    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True)
        for sheet in wb.worksheets:
            rows = list(sheet.iter_rows(values_only=True))
            if not rows: continue
            
            # 1. Find employee name in this sheet
            for r_idx, row in enumerate(rows):
                row_strs = [str(x).strip() if x is not None else '' for x in row]
                found_name = False
                for i, val in enumerate(row_strs):
                    if 'שם עובד' in val:
                        lines = val.split('\n')
                        for line in lines:
                            if 'שם עובד' in line:
                                parts = line.split(':')
                                if len(parts) > 1 and parts[1].strip():
                                    names.add(parts[1].strip())
                                elif i + 1 < len(row_strs) and row_strs[i+1]:
                                    names.add(str(row_strs[i+1]).split('\n')[0].strip())
                                found_name = True
                                break
                        break
                if found_name:
                    break

            # 2. Find positions from the 'תפקיד' column in this sheet
            header_r_idx = -1
            pos_c_idx = -1
            for r_idx, row in enumerate(rows):
                row_strs = [str(x).strip() if x is not None else '' for x in row]
                for i, val in enumerate(row_strs):
                    if 'תפקיד' in val:
                        header_r_idx = r_idx
                        pos_c_idx = i
                        break
                if header_r_idx != -1:
                    break
                    
            if header_r_idx != -1 and pos_c_idx != -1:
                for r_idx in range(header_r_idx + 1, len(rows)):
                    pos_val = str(rows[r_idx][pos_c_idx]).strip() if len(rows[r_idx]) > pos_c_idx and rows[r_idx][pos_c_idx] is not None else ''
                    if pos_val and pos_val.lower() not in ('nan', 'none', ''):
                        positions.add(pos_val)
                        
    return list(names), list(positions)

# ============================================================
# UI
# ============================================================
tabs = st.tabs(["1. העלאת קבצים", "2. התאמת עובדים", "3. התאמת עמדות", "4. הורדת קובץ מיפוי"])

with tabs[0]:
    st.subheader("העלאת נתוני מקור למערכת")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### מאגר YLM")
        ylm_emp_files = st.file_uploader("רשימת עובדים (YLM)", type=['xlsx', 'xls'], key='upload_ylm_emp', accept_multiple_files=True)
        ylm_sites_files = st.file_uploader("רשימת אתרים/עמדות (YLM)", type=['xlsx', 'xls'], key='upload_ylm_sites', accept_multiple_files=True)
    with col2:
        st.markdown("### מאגר Tabit Shift")
        tabit_report_files = st.file_uploader("דוחות תכנון מפורט (Tabit - מכיל עובדים ועמדות)", type=['xlsx', 'xls'], key='upload_tabit_reports', accept_multiple_files=True)

    if st.button("⬅️ עבד קבצים והמשך", type="primary"):
        if ylm_emp_files and tabit_report_files and ylm_sites_files:
            try:
                # 1. YLM Employees
                ylm_names_list = []
                for f in ylm_emp_files:
                    df_ylm_e = pd.read_excel(f, header=None)
                    ylm_header_idx, ylm_name_col = -1, -1
                    for r_idx, row in df_ylm_e.iterrows():
                        row_strs = [str(x).strip() for x in row.values]
                        # Looking for a column with 'שם' but ideally not 'משפחה' or 'משתמש' unless it's the only one
                        for i, val in enumerate(row_strs):
                            if 'שם' in val and 'משתמש' not in val:
                                ylm_header_idx = r_idx
                                ylm_name_col = i
                                break
                        if ylm_header_idx != -1: break
                                
                    if ylm_header_idx != -1:
                        ylm_names_list.extend(df_ylm_e.iloc[ylm_header_idx+1:, ylm_name_col].dropna().astype(str).unique().tolist())
                    else:
                        ylm_names_list.extend(df_ylm_e.iloc[:, 0].dropna().astype(str).unique().tolist())
                        
                st.session_state.ylm_names_clean = {clean_name_generic(n): n for n in set(ylm_names_list) if clean_name_generic(n) and n not in ('nan', 'None')}
                
                # 2. Tabit Reports (Names & Positions)
                tabit_names_raw, tabit_positions_raw = parse_tabit_detailed_reports(tabit_report_files)
                
                if not tabit_names_raw:
                    st.warning("לא נמצאו עובדים בקבצי התכנון של Tabit. אנא ודא שהפורמט נכון ('שם עובד:').")
                
                tabit_names = [n for n in tabit_names_raw if n.strip() and n != 'nan']
                st.session_state.tabit_names_clean = {clean_name_generic(n): n for n in set(tabit_names) if clean_name_generic(n)}

                # 3. YLM Sites
                ylm_sites_list = []
                for f in ylm_sites_files:
                    df_ylm_s = pd.read_excel(f)
                    site_col_ylm = next((c for c in df_ylm_s.columns if 'שם' in str(c)), df_ylm_s.columns[0])
                    ylm_sites_list.extend(df_ylm_s[site_col_ylm].dropna().astype(str).unique().tolist())
                st.session_state.ylm_sites = sorted(list(set(ylm_sites_list)))

                # 4. Tabit Positions
                st.session_state.tabit_positions = sorted(list(set(tabit_positions_raw)))

                # Auto Match Process
                # Employees
                matches_emp = []
                for t_clean, t_raw in st.session_state.tabit_names_clean.items():
                    best_y_raw = None
                    best_score = 0.0
                    for y_clean, y_raw in st.session_state.ylm_names_clean.items():
                        score = calc_similarity(t_clean, y_clean)
                        if score > best_score:
                            best_score = score
                            best_y_raw = y_raw
                            if score == 1.0: break
                    matches_emp.append({
                        "Tabit Name": t_raw,
                        "Matched YLM": best_y_raw if best_score > 0.6 else None,
                        "Score": round(best_score * 100, 1)
                    })
                # Add YLM employees that had no matching Tabit name at all
                matched_ylm_so_far = {item["Matched YLM"] for item in matches_emp if item["Matched YLM"]}
                for ylm_clean, ylm_raw in st.session_state.ylm_names_clean.items():
                    if ylm_raw not in matched_ylm_so_far:
                        # Append them as if Tabit is empty/missing
                        matches_emp.append({
                            "Tabit Name": "לא מופיע ב-Tabit",
                            "Matched YLM": ylm_raw,
                            "Score": 0.0
                        })
                st.session_state.emp_mapping_data = matches_emp
                # Positions
                matches_pos = []
                for t_pos in st.session_state.tabit_positions:
                    best_y_pos = None
                    best_score = 0.0
                    for y_site in st.session_state.ylm_sites:
                        score = calc_similarity(clean_name_generic(t_pos), clean_name_generic(y_site))
                        if score > best_score:
                            best_score = score
                            best_y_pos = y_site
                    matches_pos.append({
                        "Tabit Position": t_pos,
                        "Matched YLM (1)": best_y_pos if best_score >= 0.35 else None,
                        "Matched YLM (2)": None, # UI allows adding more
                        "Score": round(best_score * 100, 1)
                    })
                # Add YLM sites that had no matching Tabit position at all
                matched_ylm_sites_so_far = {item["Matched YLM (1)"] for item in matches_pos if item["Matched YLM (1)"]}
                for ylm_site in st.session_state.ylm_sites:
                    if ylm_site not in matched_ylm_sites_so_far:
                        matches_pos.append({
                            "Tabit Position": "לא מופיע ב-Tabit",
                            "Matched YLM (1)": ylm_site,
                            "Matched YLM (2)": None,
                            "Score": 0.0
                        })
                st.session_state.pos_mapping_data = matches_pos
                
                st.success("הקבצים עובדו בהצלחה! עבור ללשונית הבאה לאימות ההתאמות.")
            except Exception as e:
                st.error(f"שגיאה בעיבוד הקבצים: {e}")
        else:
            st.warning("יש להעלות את כל 4 הקבצים.")

with tabs[1]:
    st.subheader("אימות התאמות עובדים")
    if 'emp_mapping_data' in st.session_state:
        st.info("כאן תוכל לתקן את ההתאמות שהמערכת ביצעה. עובדים ללא התאמה יישמרו כ-'לא משוייך'.")
        
        ylm_options = ["לא משוייך", ""] + list(st.session_state.ylm_names_clean.values())
        
        mapped_emp_results = []
        # Keep track of multiple "לא מופיע ב-Tabit" to avoid duplicate keys
        missing_tabit_counter = 1
        for item in st.session_state.emp_mapping_data:
            t_name = item['Tabit Name']
            y_name = item['Matched YLM']
            
            idx = ylm_options.index(y_name) if y_name in ylm_options else 0
            
            ui_key = f"emp_{t_name}_{missing_tabit_counter}" if t_name == "לא מופיע ב-Tabit" else f"emp_{t_name}"
            if t_name == "לא מופיע ב-Tabit": missing_tabit_counter += 1
            selected_y = st.selectbox(
                f"Tabit: **{t_name}** | (ביטחון: {item['Score']}%)",
                options=ylm_options,
                index=idx,
                key=ui_key
            )
            
            final_y = selected_y if selected_y and selected_y != "לא משוייך" else "לא משוייך"
            mapped_emp_results.append({"שם מסידור העבודה (Tabit)": t_name, "שם בדוח נוכחות (YLM)": final_y})
                
        if st.button("שמור התאמות עובדים"):
            st.session_state.final_emp_map = mapped_emp_results
            st.success("המיפוי נשמר!")
    else:
        st.warning("אנא עבד את הקבצים בלשונית הראשונה.")

with tabs[2]:
    st.subheader("אימות התאמות עמדות ואתרים")
    if 'pos_mapping_data' in st.session_state:
        st.info("בחר לאילו אתרי YLM כל עמדה בסידור מקושרת (ניתן לבחור מרובים).")
        
        mapped_pos_results = []
        missing_tabit_counter = 1
        for item in st.session_state.pos_mapping_data:
            t_pos = item['Tabit Position']
            y_pos_default = item['Matched YLM (1)']
            defaults = [y_pos_default] if y_pos_default in st.session_state.ylm_sites else []
            
            ui_key = f"pos_{t_pos}_{missing_tabit_counter}" if t_pos == "לא מופיע ב-Tabit" else f"pos_{t_pos}"
            if t_pos == "לא מופיע ב-Tabit": missing_tabit_counter += 1
            
            selected_sites = st.multiselect(
                f"עמדה בTabit: **{t_pos}**",
                options=st.session_state.ylm_sites,
                default=defaults,
                key=ui_key
            )
            
            if selected_sites:
                mapped_pos_results.append({
                    "עמדת סידור (Tabit)": t_pos, 
                    "אתרי נוכחות (YLM)": ",".join(selected_sites)
                })
            else:
                mapped_pos_results.append({
                    "עמדת סידור (Tabit)": t_pos, 
                    "אתרי נוכחות (YLM)": "לא משוייך"
                })
                
        if st.button("שמור התאמות עמדות"):
            st.session_state.final_pos_map = mapped_pos_results
            st.success("המיפוי נשמר!")
    else:
        st.warning("אנא עבד את הקבצים בלשונית הראשונה.")

with tabs[3]:
    st.subheader("הורדת קובץ המסטאר")
    if 'final_emp_map' in st.session_state and 'final_pos_map' in st.session_state:
        df_emp = pd.DataFrame(st.session_state.final_emp_map)
        df_pos = pd.DataFrame(st.session_state.final_pos_map)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_emp.to_excel(writer, index=False, sheet_name='עובדים')
            df_pos.to_excel(writer, index=False, sheet_name='עמדות')
            
        excel_data = output.getvalue()
        
        st.markdown("### הכל מוכן!")
        st.markdown("הורד את קובץ המיפוי (Master_Mapping.xlsx). תוכל להעלות אותו בדף **דוח פערי משמרות** כדי לחסוך את שלב הזיהוי הידני.")
        st.download_button(
            label="💾 הורד קובץ מיפוי קבוע",
            data=excel_data,
            file_name="Master_Mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.warning("אנא השלם ושמור את ההתאמות בלשוניות הקודמות.")
