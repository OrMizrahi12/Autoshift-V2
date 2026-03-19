
import streamlit as st
import pandas as pd
from data_manager import load_data, get_shift_columns
import scheduler
import uuid  # For unique IDs
from excel_exporter import generate_styled_excel

# --- Shared Constants ---
ROW_LABELS = {
    "morning": "בוקר (07-15)",
    "afternoon": "צהריים (15-23)",
    "night": "לילה (23-07)",
    "double_m": "יכול כפולה בוקר (07-19)",
    "double_n": "יכול כפולה לילה (19-07)"
}

st.set_page_config(page_title="AutoShift - שיבוץ משמרות אוטומטי", layout="wide", initial_sidebar_state="expanded")

def load_css():
    with open("style.css", "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

import firebase_manager

st.title("AutoShift: מערכת שיבוץ משמרות חכמה")
st.markdown("---")

# --- LOGIN SYSTEM ---
if 'user_email' not in st.session_state:
    st.session_state['user_email'] = None

if not st.session_state['user_email']:
    st.markdown("## התחברות למערכת 🔐")
    st.info("אנא הזן את כתובת המייל הארגונית/אישית שלך. המערכת תשמור את נתוני השיבוץ שלך באופן פרטי ומאובטח בענן תחת כתובת זו.")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            email_input = st.text_input("כתובת אימייל (Google / אחר):", placeholder="example@gmail.com")
            submit = st.form_submit_button("התחבר", type="primary", use_container_width=True)
            
            if submit:
                if email_input and "@" in email_input and "." in email_input:
                    st.session_state['user_email'] = email_input.strip().lower()
                    st.rerun()
                else:
                    st.error("אנא הזן כתובת אימייל תקינה.")
    st.stop() # עצור את המשך טעינת האפליקציה עד להתחברות

# --- FIREBASE MANAGER SIDEBAR ---
with st.sidebar:
    st.header("👤 פרופיל משתמש")
    st.success(f"מחובר כ:\n**{st.session_state['user_email']}**")
    if st.button("🚪 התנתק", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
        
    st.markdown("---")
    st.header("💾 גיבוי בענן (Firebase)")
    st.markdown("האפליקציה שומרת את הנתונים שלך באופן פרטי ואוטומטי בענן לאחר כל שינוי.")
    
    st.markdown("#### איפוסי מערכת")
    if st.button("🗑️ מחק הכל מהענן והתחל מחדש", use_container_width=True):
        with st.spinner("מוחק נתונים..."):
            firebase_manager.delete_state_from_firebase(st.session_state['user_email'])
            # Clear local session state components but keep email
            email = st.session_state['user_email']
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state['user_email'] = email
            st.rerun()

# --- AUTO LOAD STATE FROM FIREBASE ---
if 'firebase_loaded' not in st.session_state:
    with st.spinner("טוען נתונים מהענן (אם קיימים)..."):
        firebase_manager.load_state_from_firebase(st.session_state, st.session_state['user_email'])
    st.session_state['firebase_loaded'] = True

# --- Session State Initialization ---
if 'positions' not in st.session_state:
    st.session_state['positions'] = []

if 'constraints' not in st.session_state:
    st.session_state['constraints'] = {
        "no_overlap": True,
        "no_back_to_back": True,
        "min_rest": 8,
        "allow_double": True
    }

# --- CONFIGURATION & UPLOAD SECTION ---
st.markdown("### שלב 1: העלאת נתוני עובדים וזיהוי עמדות")
uploaded_file = st.file_uploader("בחר קובץ Excel (המערכת תזהה אוטומטית את העמדות מהקובץ)", type=['xlsx', 'xls'])

if uploaded_file:
    file_id = getattr(uploaded_file, "file_id", str(uploaded_file.size) + uploaded_file.name)
    if 'current_file_id' not in st.session_state or st.session_state['current_file_id'] != file_id:
        df, header_idx = load_data(uploaded_file)
        if df is not None:
            st.session_state['employees_df'] = df
            st.session_state['current_file_id'] = file_id
            st.session_state['header_idx'] = header_idx

if st.session_state.get('employees_df') is not None:
    df = st.session_state.get('employees_df')
    header_idx = st.session_state.get('header_idx')
    
    if df is not None:
        if header_idx is not None:
            st.success(f"נתוני עובדים זמינים לפעולה (זוהתה כותרת בשורה {header_idx+1})")
        else:
            st.success("הנתונים נטענו בהצלחה מהענן!")
        
        # --- 1. Identify Columns ---
        cols = df.columns.tolist()
        potential_shifts = get_shift_columns(df)
        
        # Automatic Column Detection
        name_col_candidates = [c for c in cols if "עובדים" in str(c) or "Name" in str(c)]
        name_col = name_col_candidates[0] if name_col_candidates else cols[0]
        
        role_col_candidates = [c for c in cols if "תפקידים" in str(c) or "Position" in str(c) or "Role" in str(c)]
        role_col = role_col_candidates[0] if role_col_candidates else None
        
        # --- 2. Extract Positions from File ---
        if role_col:
            # Flatten list of roles (comma separated)
            raw_roles = df[role_col].dropna().astype(str).tolist()
            unique_roles_set = set()
            for r in raw_roles:
                # Split by comma and strip whitespace
                parts = [p.strip() for p in r.split(',') if p.strip()]
                unique_roles_set.update(parts)
            
            unique_roles = sorted(list(unique_roles_set))
            
            # Sync with session state
            # 0. Initialize deleted tracker
            if 'deleted_positions' not in st.session_state:
                st.session_state['deleted_positions'] = set()

            # 1. Initialize if empty
            if 'positions' not in st.session_state:
                st.session_state['positions'] = []
            
            current_names = [p['name'] for p in st.session_state['positions']]
            
            # 2. Add new roles found in file
            new_roles_count = 0
            for role in unique_roles:
                # Only add if NOT already participating AND NOT explicitly deleted
                if role not in current_names and role not in st.session_state['deleted_positions']:
                    # Default settings for new found role
                    st.session_state['positions'].append({
                        "id": str(uuid.uuid4()),  # Stable unique ID is crucial for deletion
                        "name": role,
                        "guards_morning": 1,
                        "guards_afternoon": 1,
                        "guards_night": 1,
                        "priority": 5,
                        "priority_morning": 1,
                        "priority_afternoon": 1,
                        "priority_night": 1,
                        "active_shifts": {d: {'M': True, 'A': True, 'N': True} for d in potential_shifts}
                    })
                    new_roles_count += 1
            
            if new_roles_count > 0:
                st.toast(f"זוהו {new_roles_count} עמדות חדשות מהקובץ והתווספו להגדרות!", icon="🏢")

        # --- 3. Position Configuration UI ---
        st.divider()
        st.subheader("הגדרת עמדות (זוהה מתוך הקובץ)")
        st.info("כאן מופיעות העמדות שנמצאו בקובץ. ניתן לשנות את דרישות האיוש או למחוק עמדות לא רלוונטיות.")

        if not st.session_state['positions']:
            st.warning("לא נמצאו עמדות בקובץ. אנא וודא שיש עמודת 'תפקידים'.")
        else:
            # Migration: Ensure all positions have IDs (in case of old state)
            for p in st.session_state['positions']:
                if 'id' not in p:
                    p['id'] = str(uuid.uuid4())

            # --- Bulk-select toolbar ---
            pos_tb_l, pos_tb_m, pos_tb_r = st.columns([2, 2, 3])
            
            # Helper to access checkboxes safely via ID
            def get_chk_key(pid): return f"pos_chk_{pid}"
            
            with pos_tb_l:
                if st.button("✅ בחר הכל", key="pos_sel_all", use_container_width=True):
                    for p in st.session_state['positions']:
                        st.session_state[get_chk_key(p['id'])] = True
            with pos_tb_m:
                if st.button("☐ בטל הכל", key="pos_desel_all", use_container_width=True):
                    for p in st.session_state['positions']:
                        st.session_state[get_chk_key(p['id'])] = False
            with pos_tb_r:
                # Count selected by checking state for each ID
                selected_ids = [
                    p['id'] for p in st.session_state['positions']
                    if st.session_state.get(get_chk_key(p['id']), False)
                ]
                n_pos_selected = len(selected_ids)
                
                if st.button(
                    f"🗑️ מחק {n_pos_selected} עמדות מסומנות" if n_pos_selected else "🗑️ מחק מסומנות",
                    key="pos_bulk_del",
                    type="primary" if n_pos_selected else "secondary",
                    disabled=(n_pos_selected == 0),
                    use_container_width=True
                ):
                    # 1. Identify names to exclude (for future loads)
                    # 2. Rebuild list excluding selected IDs
                    new_pos_list = []
                    for p in st.session_state['positions']:
                        if p['id'] in selected_ids:
                            st.session_state['deleted_positions'].add(p['name'])
                            # Clear widget state
                            st.session_state.pop(get_chk_key(p['id']), None)
                        else:
                            new_pos_list.append(p)
                    
                    st.session_state['positions'] = new_pos_list
                    st.rerun()

            # --- Per-position rows ---
            with st.expander("➕ הוסף עמדה חדשה", expanded=False):
                with st.form("add_position_form"):
                    new_pos_name = st.text_input("שם העמדה*", placeholder="לדוגמה: שער ראשי")
                    new_pos_prio = st.number_input("עדיפות כללית (1-10)", 1, 10, 5, help="1 = הכי חשוב")
                    
                    st.markdown("דרישות איוש (מס' מאבטחים):")
                    g1, g2, g3 = st.columns(3)
                    new_g_m = g1.number_input("בוקר (07-15)", 0, 10, 1)
                    new_g_a = g2.number_input("צהריים (15-23)", 0, 10, 1)
                    new_g_n = g3.number_input("לילה (23-07)", 0, 10, 1)
                    
                    st.markdown("עדיפויות משמרת (1-3):")
                    p1, p2, p3 = st.columns(3)
                    new_p_m = p1.number_input("עדיפות בוקר", 1, 3, 1)
                    new_p_a = p2.number_input("עדיפות צהריים", 1, 3, 1)
                    new_p_n = p3.number_input("עדיפות לילה", 1, 3, 1)
                    
                    submit_new_pos = st.form_submit_button("שמור עמדה חדשה", type="primary", use_container_width=True)
                    
                    if submit_new_pos:
                        if new_pos_name.strip():
                            # Clear from deleted if it was there
                            if 'deleted_positions' in st.session_state:
                                st.session_state['deleted_positions'].discard(new_pos_name.strip())
                            
                            st.session_state['positions'].append({
                                "id": str(uuid.uuid4()),
                                "name": new_pos_name.strip(),
                                "guards_morning": new_g_m,
                                "guards_afternoon": new_g_a,
                                "guards_night": new_g_n,
                                "priority": new_pos_prio,
                                "priority_morning": new_p_m,
                                "priority_afternoon": new_p_a,
                                "priority_night": new_p_n,
                                "active_shifts": {d: {'M': True, 'A': True, 'N': True} for d in potential_shifts}
                            })
                            st.success(f"העמדה {new_pos_name} נוספה בהצלחה!")
                            st.rerun()
                        else:
                            st.error("חובה להזין שם עמדה")
            
            st.markdown("---")
            # Loop by index is fine for layout, but Keys must use ID
            for idx, pos in enumerate(st.session_state['positions']):
                pid = pos['id']
                
                chk_col, name_col_ui, del_col = st.columns([0.5, 5.5, 0.8])
                with chk_col:
                    st.checkbox(
                        label=f"בחר עמדה {pos['name']}", 
                        key=f"pos_chk_{pid}", # Using ID for key
                        label_visibility="collapsed"
                    )
                with del_col:
                    if st.button("🗑️", key=f"quick_del_{pid}", help=f"מחק עמדה: {pos['name']}", use_container_width=True):
                        st.session_state['deleted_positions'].add(pos['name'])
                        st.session_state['positions'] = [p for p in st.session_state['positions'] if p['id'] != pid]
                        st.rerun()
                        
                with name_col_ui:
                    with st.expander(f"🏢 {pos['name']} — לחץ לעריכה", expanded=False):
                        # 1. Basic Info
                        c1, c2 = st.columns([3, 2])
                        new_name = c1.text_input("שם העמדה", pos['name'], key=f"p_name_{pid}")
                        pos_priority = c2.number_input(
                            "עדיפות עמדה (1-10)", 1, 10, pos.get("priority", 5),
                            key=f"prio_{pid}", help="1 = העמדה הכי חשובה למלא."
                        )

                        # 2. Activity / Schedule Matrix (Manual Grid for stability)
                        st.markdown("**📅 לו\"ז פעילות (סמן מתי העמדה פעילה)**")
                        
                        current_active = pos.get('active_shifts', {})
                        # Ensure all days/shifts exist
                        for d in potential_shifts:
                            if d not in current_active:
                                current_active[d] = {'M': True, 'A': True, 'N': True}
                        
                        new_active_shifts = {}
                        
                        # Labels for shifts
                        shift_names = {"M": "בוקר", "A": "צהריים", "N": "לילה"}
                        
                        # Header Row: Days
                        # Using small padding/columns to fit days
                        day_cols = st.columns([1.2] + [1] * len(potential_shifts))
                        day_cols[0].markdown("**משמרת**")
                        for i, d_label in enumerate(potential_shifts):
                            day_cols[i+1].markdown(f"**{d_label}**")
                        
                        # Data Rows
                        for s_key, s_label in shift_names.items():
                            row_cols = st.columns([1.2] + [1] * len(potential_shifts))
                            row_cols[0].write(s_label)
                            for i, d_label in enumerate(potential_shifts):
                                chk_key = f"act_{pid}_{d_label}_{s_key}"
                                is_active = current_active.get(d_label, {}).get(s_key, True)
                                
                                # Checkbox in each cell
                                val = row_cols[i+1].checkbox(
                                    "", 
                                    value=is_active, 
                                    key=chk_key,
                                    label_visibility="collapsed"
                                )
                                if d_label not in new_active_shifts:
                                    new_active_shifts[d_label] = {}
                                new_active_shifts[d_label][s_key] = val

                        st.markdown("דרישות איוש (מס' מאבטחים כשהעמדה פעילה):")
                        g1, g2, g3 = st.columns(3)
                        g_m = g1.number_input("בוקר (07-15)", 0, 10, pos['guards_morning'], key=f"gm_{pid}")
                        g_a = g2.number_input("צהריים (15-23)", 0, 10, pos['guards_afternoon'], key=f"ga_{pid}")
                        g_n = g3.number_input("לילה (23-07)", 0, 10, pos['guards_night'], key=f"gn_{pid}")

                        st.markdown("---")
                        st.markdown("**⭐ עדיפויות משמרת** (1 = חשוב יותר)")
                        pm_col, pa_col, pn_col = st.columns(3)
                        pm = pm_col.number_input("עדיפות בוקר", 1, 3, pos.get("priority_morning", 1), key=f"pm_{pid}")
                        pa = pa_col.number_input("עדיפות צהריים", 1, 3, pos.get("priority_afternoon", 1), key=f"pa_{pid}")
                        pn = pn_col.number_input("עדיפות לילה", 1, 3, pos.get("priority_night", 1), key=f"pn_{pid}")
                        
                        # Update state
                        pos.update({
                            "name": new_name,
                            "guards_morning": g_m, "guards_afternoon": g_a, "guards_night": g_n,
                            "priority": pos_priority, "priority_morning": pm,
                            "priority_afternoon": pa, "priority_night": pn,
                            "active_shifts": new_active_shifts
                        })



        # --- 4. Constraints ---
        st.divider()
        with st.expander("הגדרות מתקדמות ואילוצים", expanded=False):
            st.markdown("#### אילוצים וחוקים")
            c = st.session_state['constraints']
            
            no_overlap = st.checkbox("איסור חפיפת משמרות", value=c['no_overlap'])
            no_back_to_back = st.checkbox("איסור משמרות רצופות", value=c['no_back_to_back'])
            min_rest = st.number_input("שעות מנוחה מינימליות", min_value=0, value=c['min_rest'])
            allow_double = st.checkbox("אפשר כפולות (ברירת מחדל לכולם)", value=c['allow_double'])
            
            if st.button("🪄 אישור כפולות גורף לכל העובדים", use_container_width=True, help="לחיצה על הכפתור תעדכן את כל העובדים שזמינים לבוקר להיות זמינים גם לכפולת בוקר, ומי שזמין ללילה לכפולת לילה."):
                if 'employees_df' in st.session_state:
                    df = st.session_state['employees_df']
                    if 'firebase_constraints_base' not in st.session_state:
                        st.session_state['firebase_constraints_base'] = {}
                    
                    # Logic to update all employees based on current M/N availability
                    for idx, row in df.iterrows():
                        # Determine current file-based availability
                        day_cols = [c for c in df.columns if c in potential_shifts]
                        
                        # Get existing override if any, else build from scratch
                        if str(idx) in st.session_state['firebase_constraints_base']:
                            emp_df = st.session_state['firebase_constraints_base'][str(idx)].copy()
                        else:
                            # Build initial from file
                            data_dict = {}
                            for s_col in day_cols:
                                val = str(row[s_col]).lower()
                                is_m = 'בוקר' in val or 'morning' in val
                                is_a = 'צהריים' in val or 'afternoon' in val
                                is_n = 'לילה' in val or 'night' in val
                                data_dict[s_col] = [is_m, is_a, is_n, False, False]
                            emp_df = pd.DataFrame(data_dict, index=[
                                ROW_LABELS["morning"], ROW_LABELS["afternoon"], ROW_LABELS["night"],
                                ROW_LABELS["double_m"], ROW_LABELS["double_n"]
                            ]).reset_index().rename(columns={'index': 'סוג משמרת'})

                        # Apply the update
                        for col in day_cols:
                            if col in emp_df.columns:
                                is_m = emp_df.loc[emp_df['סוג משמרת'] == ROW_LABELS["morning"], col].values[0]
                                is_n = emp_df.loc[emp_df['סוג משמרת'] == ROW_LABELS["night"], col].values[0]
                                if is_m: emp_df.loc[emp_df['סוג משמרת'] == ROW_LABELS["double_m"], col] = True
                                if is_n: emp_df.loc[emp_df['סוג משמרת'] == ROW_LABELS["double_n"], col] = True
                        
                        st.session_state['firebase_constraints_base'][str(idx)] = emp_df
                    
                    st.success("הכפולות עודכנו לכל העובדים בהצלחה!")
                    st.rerun()

            st.session_state['constraints'] = {
                "no_overlap": no_overlap,
                "no_back_to_back": no_back_to_back,
                "min_rest": min_rest,
                "allow_double": allow_double,
                "auto_doubles": False # Deprecated persistent toggle
            }
            
            # Column Mapping override
            st.markdown("---")
            st.markdown("#### מיפוי עמודות (אופציונלי)")
            c1, c2 = st.columns(2)
            default_name_idx = next((i for i, c in enumerate(cols) if c == name_col), 0)
            default_pos_idx = next((i for i, c in enumerate(cols) if c == role_col), 1)
            
            selected_name_col = c1.selectbox("עמודת שם", cols, index=default_name_idx)
            selected_pos_col = c2.selectbox("עמודת תפקיד", cols, index=default_pos_idx)
            
            st.session_state['col_map'] = {"name": selected_name_col, "pos": selected_pos_col, "note": None}
            st.session_state['selected_shifts'] = st.multiselect("עמודות משמרות", cols, default=potential_shifts)

        # st.session_state['employees_df'] = df # Handled at the top of the file

        # --- 5. Data Preview ---
        with st.expander("תצוגה מקדימה של הנתונים (לחץ לפתיחה)"):
            # Display parsed table for user verification
            parsed_data = []
            for _, row in df.iterrows():
                emp_name = row[name_col]
                emp_role = str(row[role_col]) if role_col and pd.notna(row[role_col]) else "לא זוהה"
                
                # Parse availability
                avail_summary = []
                for s_col in potential_shifts:
                    val = str(row[s_col]).strip()
                    if val and val.lower() != 'nan':
                        day_label = str(s_col).split(' ')[0] # Heuristic
                        clean_val = val.replace("::", ":")
                        if ":" in clean_val: clean_val = clean_val.split(":")[-1].strip()
                        if clean_val: avail_summary.append(f"{day_label}: {clean_val}")
                
                parsed_data.append({
                    "שם עובד": emp_name,
                    "תפקיד": emp_role,
                    "זמינות": ", ".join(avail_summary)
                })
            st.dataframe(pd.DataFrame(parsed_data), use_container_width=True)

        # --- DETAILED AVAILABILITY TABLES (PER EMPLOYEE) ---
        st.divider()
        st.markdown("### פרופיל אילוצים אישי (לפי עובד)")
        st.info("כאן מוצגים האילוצים של כל עובד בנפרד, בחלוקה לימים וסוגי משמרות.")

        if 'avail_updates' not in st.session_state:
            st.session_state['avail_updates'] = {}
        if 'excluded_employees' not in st.session_state:
            st.session_state['excluded_employees'] = set()

        # Detect notes column
        note_col_candidates = [c for c in cols if "הערות" in str(c) or "Comments" in str(c) or "Note" in str(c)]
        note_col = note_col_candidates[0] if note_col_candidates else None

        # Build list of employees (excluding deleted ones)
        all_emp_rows = [
            (idx, row) for idx, row in df.iterrows()
            if str(row[name_col]) not in st.session_state['excluded_employees']
        ]
        emp_indices = [idx for idx, _ in all_emp_rows]  # original df indices

        if potential_shifts:
            # --- Bulk-select toolbar for employees ---
            emp_tb_l, emp_tb_m, emp_tb_r = st.columns([2, 2, 3])
            with emp_tb_l:
                if st.button("✅ בחר הכל", key="emp_sel_all", use_container_width=True):
                    for ei in emp_indices:
                        st.session_state[f"emp_chk_{ei}"] = True
            with emp_tb_m:
                if st.button("☐ בטל הכל", key="emp_desel_all", use_container_width=True):
                    for ei in emp_indices:
                        st.session_state[f"emp_chk_{ei}"] = False
            with emp_tb_r:
                n_emp_selected = sum(
                    1 for ei in emp_indices
                    if st.session_state.get(f"emp_chk_{ei}", False)
                )
                if st.button(
                    f"🗑️ מחק {n_emp_selected} עובדים מסומנים" if n_emp_selected else "🗑️ מחק מסומנים",
                    key="emp_bulk_del",
                    type="primary" if n_emp_selected else "secondary",
                    disabled=(n_emp_selected == 0),
                    use_container_width=True
                ):
                    for ei in emp_indices:
                        if st.session_state.get(f"emp_chk_{ei}", False):
                            emp_row = df.loc[ei]
                            st.session_state['excluded_employees'].add(str(emp_row[name_col]))
                            st.session_state.pop(f"emp_chk_{ei}", None)
                            st.session_state.pop(f"emp_chk_widget_{ei}", None) # Cleanup legacy
                    st.rerun()

            st.markdown("")

            # --- הוספת עובד חדש ---
            with st.expander("➕ הוסף עובד חדש", expanded=False):
                with st.form("add_employee_form"):
                    new_emp_name = st.text_input("שם העובד*")
                    new_emp_role = st.selectbox("תפקיד דיפולטיבי (ניתן לשינוי)", [""] + [p['name'] for p in st.session_state['positions']])
                    new_emp_note = st.text_input("הערות (אופציונלי)")
                    submit_new_emp = st.form_submit_button("שמור עובד חדש")
                    if submit_new_emp:
                        if new_emp_name.strip():
                            new_row = {name_col: new_emp_name.strip()}
                            if role_col and new_emp_role:
                                new_row[role_col] = new_emp_role
                            if note_col and new_emp_note:
                                new_row[note_col] = new_emp_note
                            for s_col in potential_shifts:
                                new_row[s_col] = "" 
                            
                            new_df = pd.DataFrame([new_row])
                            st.session_state['employees_df'] = pd.concat([st.session_state['employees_df'], new_df], ignore_index=True)
                            st.success(f"העובד {new_emp_name} נוסף בהצלחה!")
                            st.rerun()
                        else:
                            st.error("חובה להזין שם עובד")

            # Container for capturing the current state of edits
            collected_overrides = {}
            collected_role_updates = {} # Store position capability changes
            collected_pref_weights = {} # Store per-employee position preference weights (0-10)
            collected_max_shifts = {} # Store per-employee max shifts (default 6)
            collected_fixed_shifts = {} # Store per-employee fixed assignments (IRON shifts)


            for idx, row in all_emp_rows:
                emp_name = row[name_col]
                header_text = f"👤 {emp_name}"

                emp_chk_col, emp_exp_col = st.columns([0.5, 6.5])
                with emp_chk_col:
                    st.checkbox(
                        label=f"בחר את {emp_name}", 
                        key=f"emp_chk_{idx}",
                        label_visibility="collapsed"
                    )
                with emp_exp_col:
                    with st.expander(header_text, expanded=False):
                        # Show notes if they exist
                        if note_col and pd.notna(row.get(note_col, None)):
                            note_val = str(row[note_col]).strip()
                            if note_val and note_val.lower() != 'nan':
                                st.info(f"📝 **הערה:** {note_val}")

                        # --- Workload Settings ---
                        m_col1, m_col2 = st.columns([2, 5])
                        max_s = m_col1.number_input(
                            "מכסת משמרות מקסימלית:",
                            min_value=1,
                            max_value=14,
                            value=6,
                            key=f"max_s_{idx}",
                            help="כמה משמרות סה\"כ מותר לשבץ עובד זה בסידור הנוכחי."
                        )
                        collected_max_shifts[idx] = max_s
                        
                        st.markdown("---")

                        # Build initial from file (Always source of truth for structure)
                        data_dict = {}
                        for s_col in potential_shifts:
                            val = str(row[s_col]).lower()
                            day_label = str(s_col).strip()
                            is_m = 'בוקר' in val or 'morning' in val
                            is_a = 'צהריים' in val or 'afternoon' in val
                            is_n = 'לילה' in val or 'night' in val

                            data_dict[day_label] = [is_m, is_a, is_n, False, False]

                        df_emp = pd.DataFrame(data_dict, index=[
                            ROW_LABELS["morning"],
                            ROW_LABELS["afternoon"],
                            ROW_LABELS["night"],
                            ROW_LABELS["double_m"],
                            ROW_LABELS["double_n"]
                        ])

                        # Prepare for Display
                        df_display = df_emp.reset_index().rename(columns={'index': 'סוג משמרת'})

                        # 1. Identify Day Columns
                        day_cols = [c for c in df_display.columns if c != 'סוג משמרת']

                        # 2. Reverse so first day (Sun) is closest to Shift label
                        day_cols_reversed = day_cols[::-1]

                        # 3. Final order: [Late Days ... Early Days, ShiftLabel]
                        new_order = day_cols_reversed + ['סוג משמרת']
                        df_display = df_display[new_order]

                        # Ensure stable base from Firebase is applied BEFORE data_editor ONCE
                        if 'firebase_constraints_base' in st.session_state and str(idx) in st.session_state['firebase_constraints_base']:
                            loaded_df = st.session_state['firebase_constraints_base'][str(idx)]
                            # Only merge boolean columns that exist in both gracefully (for new days added)
                            for col in loaded_df.columns:
                                if col in df_display.columns and col != 'סוג משמרת':
                                    df_display[col] = loaded_df[col]

                        # Config
                        col_config = {
                            "סוג משמרת": st.column_config.TextColumn("סוג משמרת", disabled=True)
                        }
                        for col in data_dict.keys():
                            col_config[col] = st.column_config.CheckboxColumn(disabled=False)

                        # Data Editor
                        edited_display = st.data_editor(
                            df_display,
                            column_config=col_config,
                            disabled=["סוג משמרת"],
                            key=f"emp_edit_{idx}",
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        if 'current_edited_displays' not in st.session_state:
                            st.session_state['current_edited_displays'] = {}
                        st.session_state['current_edited_displays'][str(idx)] = edited_display

                        # Store back in format expected by scheduler
                        collected_overrides[idx] = edited_display.set_index("סוג משמרת")
                        
                        # --- Position Capability Management (User Request) ---
                        if role_col:
                            st.divider()
                            st.caption("🛠️ ניהול הסמכות (עמדות מורשות)")
                            
                            # Flatten active positions names
                            active_pos_names = [p['name'] for p in st.session_state['positions']]
                            
                            # Parse current employee roles
                            curr_val = str(row[role_col]) if pd.notna(row[role_col]) else ""
                            curr_roles = [r.strip() for r in curr_val.split(',') if r.strip()]
                            # Filter to only show valid/active roles as default
                            valid_defaults = [r for r in curr_roles if r in active_pos_names]
                            
                            selected_roles = st.multiselect(
                                "בחר עמדות שהעובד מוסמך אליהן:",
                                options=active_pos_names,
                                default=valid_defaults,
                                key=f"roles_sel_{idx}",
                                placeholder="בחר עמדות..."
                            )
                            # Store for processing
                            collected_role_updates[idx] = selected_roles
                            
                            # --- Position Preference Weights ---
                            if selected_roles:
                                st.markdown("")
                                st.caption("⭐ העדפות עמדה (0 = רק בחירום, 10 = עדיפות מקסימלית)")
                                emp_prefs = {}
                                n_pref_cols = min(len(selected_roles), 4)
                                pref_cols = st.columns(n_pref_cols)
                                for r_i, role_name in enumerate(selected_roles):
                                    with pref_cols[r_i % n_pref_cols]:
                                        score = st.number_input(
                                            f"🏢 {role_name}",
                                            min_value=0,
                                            max_value=10,
                                            value=5,
                                            step=1,
                                            key=f"pref_{idx}_{role_name}",
                                            help=f"העדפה של {emp_name} לעמדת {role_name}"
                                        )
                                        emp_prefs[role_name] = score
                                collected_pref_weights[idx] = emp_prefs

                            # --- Fixed Shifts (Iron Shifts) ---
                            st.divider()
                            st.caption("⚓ **משמרות ברזל (שיבוץ קבוע)**")
                            st.info("השימוש במשמרות ברזל פירושו שהאלגוריתם יחשיב את העובד כמשובץ באופן אוטומטי לעמדה וזמן אלו.")
                            
                            fs_key = f"fixed_shifts_list_{idx}"
                            if fs_key not in st.session_state:
                                st.session_state[fs_key] = []
                            
                            # Display existing fixed shifts
                            for f_idx, f_shift in enumerate(st.session_state[fs_key]):
                                fs_c1, fs_c2, fs_c3, fs_c4 = st.columns([1, 1, 1.5, 0.5])
                                fs_c1.write(f"📅 {f_shift['day']}")
                                fs_c2.write(f"⏱️ {f_shift['shift']}")
                                fs_c3.write(f"🏢 {f_shift['pos_name']}")
                                if fs_c4.button("🗑️", key=f"del_fs_{idx}_{f_idx}"):
                                    st.session_state[fs_key].pop(f_idx)
                                    st.rerun()
                            
                            # Form to add new fixed shift
                            with st.popover("➕ הוסף משמרת ברזל"):
                                afs_c1, afs_c2 = st.columns(2)
                                f_day = afs_c1.selectbox("יום", options=potential_shifts, key=f"f_d_{idx}")
                                f_shift_type = afs_c2.selectbox("סוג משמרת", options=["M", "A", "N"], format_func=lambda x: {"M":"בוקר", "A":"צהריים", "N":"לילה"}[x], key=f"f_s_{idx}")
                                f_pos = st.selectbox("עמדה", options=[p['name'] for p in st.session_state['positions']], key=f"f_p_{idx}")
                                
                                if st.button("שמור משמרת ברזל", key=f"btn_fs_{idx}"):
                                    st.session_state[fs_key].append({
                                        "day": f_day,
                                        "shift": f_shift_type,
                                        "pos_name": f_pos
                                    })
                                    st.rerun()
                            
                            collected_fixed_shifts[idx] = st.session_state[fs_key]


        # --- 6. Schedule Action ---
        st.divider()
        st.header("יצירת השיבוץ")
        
        col1, col2 = st.columns(2)
        with col1:
             st.info(f"עמדות פעילות: {len(st.session_state['positions'])}")
             st.info(f"עובדים בקובץ: {len(df)}")
        with col2:
             st.info(f"אילוצים פעילים: {len([k for k,v in st.session_state['constraints'].items() if v])}")
             
        calc_potential_ui = st.checkbox("חשב והצג מועמדים פוטנציאליים לגישור פערים (מאריך את זמן החישוב)", value=False)
        
        generate_clicked = st.button("התחל שיבוץ אוטומטי (AutoShift)", type="primary")
        if generate_clicked:
            with st.spinner("מבצע אופטימיזציה..."):
                try:
                    current_overrides = collected_overrides
                    shifts_to_use = st.session_state.get('selected_shifts', potential_shifts)
                    col_map_to_use = st.session_state.get('col_map', {"name": name_col, "pos": role_col, "note": None})
                    
                    # Filter out excluded employees (deleted) before solving
                    df_all = st.session_state['employees_df']
                    excluded_names = st.session_state.get('excluded_employees', set())
                    df_solver = df_all[~df_all[name_col].astype(str).isin(excluded_names)].copy()
                    if role_col and collected_role_updates:
                        for r_idx, r_list in collected_role_updates.items():
                            df_solver.at[r_idx, role_col] = ", ".join(r_list)
                    
                    results = scheduler.solve_roster(
                        df_solver,
                        st.session_state['positions'],
                        st.session_state['constraints'],
                        col_map=col_map_to_use,
                        shifts=shifts_to_use,
                        avail_overrides=current_overrides,
                        pref_weights=collected_pref_weights,
                        max_shifts_map=collected_max_shifts,
                        fixed_shifts_map=collected_fixed_shifts,
                        calc_potentials=calc_potential_ui
                    )
                    st.session_state['latest_roster_results'] = results
                except Exception as e:
                    st.error(f"שגיאה בתהליך השיבוץ: {e}")

        # Render saved roster if it exists, matching the inner indentation scope
        if 'latest_roster_results' in st.session_state:
            class DummyContext:
                def __enter__(self): return self
                def __exit__(self, exc_type, exc_val, exc_tb): pass
            with DummyContext():
                try:
                    results = st.session_state['latest_roster_results']
                    if results and results.get('roster') is not None:
                        st.success(f"נמצא פתרון! (סטטוס: {results['status']})")
                        
                        # Process Roster for Visualization
                        roster = results['roster']
                        unique_positions = roster['עמדה'].unique()
                        sorted_days = sorted(roster['יום'].unique()) 
                        
                        # Helper for shift time display
                        def get_time_range(raw_s):
                            if raw_s == 'M': return "07:00 - 15:00"
                            if raw_s == 'A': return "15:00 - 23:00"
                            if raw_s == 'N': return "23:00 - 07:00"
                            if raw_s == 'DM': return "07:00 - 19:00"
                            if raw_s == 'DN': return "19:00 - 07:00"
                            if raw_s == 'SHORTAGE': return ""
                            return ""

                        # --- Custom CSS ---
                        schedule_css = """
                        <style>
                        .schedule-container {
                            direction: rtl;
                            font-family: 'Segoe UI', Tahoma, sans-serif;
                            margin: 1rem 0 2rem 0;
                            overflow-x: auto;
                        }
                        .schedule-table {
                            width: 100%;
                            border-collapse: collapse;
                            border: 1px solid #d0d5dd;
                            font-size: 13px;
                            table-layout: fixed;
                        }
                        .schedule-table th.day-header {
                            background: #f8f9fa;
                            border: 1px solid #d0d5dd;
                            padding: 8px 4px;
                            text-align: center;
                            font-weight: 600;
                            color: #344054;
                            font-size: 13px;
                        }
                        .schedule-table th.day-header .day-name {
                            font-size: 14px;
                            color: #1d2939;
                        }
                        .schedule-table th.day-header .day-date {
                            font-size: 11px;
                            color: #667085;
                        }
                        .pos-header-row td {
                            background: linear-gradient(135deg, #0ea5e9, #38bdf8, #7dd3fc);
                            color: white;
                            text-align: center;
                            font-weight: 700;
                            font-size: 15px;
                            padding: 10px 6px;
                            border: 1px solid #0ea5e9;
                            letter-spacing: 0.5px;
                        }
                        .shift-badge {
                            display: inline-block;
                            padding: 3px 10px;
                            border-radius: 4px;
                            font-weight: 600;
                            font-size: 12px;
                            color: #fff;
                            min-width: 50px;
                            text-align: center;
                        }
                        .badge-morning { background: #0284c7; }
                        .badge-afternoon { background: #f59e0b; }
                        .badge-night { background: #4338ca; }
                        .schedule-table td.shift-label-cell {
                            background: #f0f9ff;
                            border: 1px solid #d0d5dd;
                            padding: 6px 4px;
                            text-align: center;
                            vertical-align: middle;
                            width: 70px;
                            min-width: 70px;
                        }
                        .schedule-table td.worker-cell {
                            border: 1px solid #e4e7ec;
                            padding: 6px 4px;
                            text-align: center;
                            vertical-align: middle;
                            min-height: 44px;
                            background: #ffffff;
                        }
                        .schedule-table td.worker-cell:hover {
                            background: #f0f9ff;
                        }
                        .worker-name {
                            font-weight: 600;
                            color: #1d2939;
                            font-size: 12.5px;
                        }
                        .worker-time {
                            font-size: 11px;
                            color: #667085;
                            margin-top: 1px;
                        }
                        .schedule-table td.shortage-cell {
                            background: #fff7ed !important;
                            border: 1px solid #fb923c;
                        }
                        .shortage-text {
                            font-weight: 700;
                            color: #c2410c;
                            font-size: 12px;
                        }
                        .shortage-hours {
                            font-size: 11px;
                            color: #ea580c;
                        }
                        .empty-cell {
                            color: #d0d5dd;
                            font-size: 11px;
                        }
                        </style>
                        """
                        st.markdown(schedule_css, unsafe_allow_html=True)

                        # --- Build HTML for each position ---
                        days_in_order = sorted_days  # Chronological: Sun, Mon, Tue...
                        num_days = len(days_in_order)

                        for pos in unique_positions:
                            pos_df = roster[roster['עמדה'] == pos]
                            
                            html = '<div class="schedule-container">'
                            html += '<table class="schedule-table">'
                            
                            # Row 1: Day Headers (משמרת first = rightmost in RTL)
                            html += '<tr>'
                            html += '<th class="day-header" style="width:70px;"><div class="day-name">משמרת</div></th>'
                            for d in days_in_order:
                                html += f'<th class="day-header"><div class="day-name">{d}</div></th>'
                            html += '</tr>'
                            
                            # Row 2: Position Header (full width)
                            html += f'<tr class="pos-header-row"><td colspan="{num_days + 1}">🛡️ {pos}</td></tr>'
                            
                            # Shift Groups
                            shift_groups = [
                                ("בוקר", "badge-morning", ['M', 'DM']),
                                ("צהריים", "badge-afternoon", ['A']),
                                ("לילה", "badge-night", ['N', 'DN']),
                            ]
                            
                            for group_label, badge_class, raw_codes in shift_groups:
                                # Collect workers per day (including shortages)
                                day_workers = {}
                                max_depth = 0
                                
                                for d in days_in_order:
                                    workers_regular = pos_df[
                                        (pos_df['יום'] == d) &
                                        (pos_df['raw_shift'].isin(raw_codes))
                                    ]
                                    # Also check for shortage rows matching this shift group
                                    shortage_rows = pos_df[
                                        (pos_df['יום'] == d) &
                                        (pos_df['raw_shift'] == 'SHORTAGE')
                                    ]
                                    # Filter shortage rows to this shift group by checking shift description
                                    relevant_shortages = []
                                    for _, sr in shortage_rows.iterrows():
                                        shift_desc = str(sr['משמרת']).lower()
                                        if group_label == "בוקר" and 'בוקר' in shift_desc:
                                            relevant_shortages.append(sr)
                                        elif group_label == "צהריים" and 'צהריים' in shift_desc:
                                            relevant_shortages.append(sr)
                                        elif group_label == "לילה" and 'לילה' in shift_desc:
                                            relevant_shortages.append(sr)
                                    
                                    entries = []
                                    for _, w in workers_regular.iterrows():
                                        time_str = get_time_range(w['raw_shift'])
                                        entries.append({
                                            'name': w['עובד'],
                                            'time': time_str,
                                            'is_shortage': False
                                        })
                                    for sr in relevant_shortages:
                                        entries.append({
                                            'name': sr['עובד'],
                                            'time': sr['משמרת'],
                                            'is_shortage': True
                                        })
                                    
                                    day_workers[d] = entries
                                    max_depth = max(max_depth, len(entries))
                                
                                if max_depth == 0:
                                    continue
                                
                                # Render rows for this shift group
                                for i in range(max_depth):
                                    html += '<tr>'
                                    # Shift label cell FIRST (= rightmost in RTL)
                                    if i == 0:
                                        html += f'<td class="shift-label-cell" rowspan="{max_depth}">'
                                        html += f'<span class="shift-badge {badge_class}">{group_label}</span>'
                                        html += '</td>'
                                    # Day cells in chronological order
                                    for d in days_in_order:
                                        entries = day_workers.get(d, [])
                                        if i < len(entries):
                                            entry = entries[i]
                                            if entry['is_shortage']:
                                                html += f'<td class="worker-cell shortage-cell">'
                                                html += f'<div class="shortage-text">{entry["name"]}</div>'
                                                html += f'<div class="shortage-hours">{entry["time"]}</div>'
                                                html += '</td>'
                                            else:
                                                html += f'<td class="worker-cell">'
                                                html += f'<div class="worker-name">{entry["name"]}</div>'
                                                html += f'<div class="worker-time">{entry["time"]}</div>'
                                                html += '</td>'
                                        else:
                                            html += '<td class="worker-cell"><div class="empty-cell"></div></td>'
                                    html += '</tr>'
                            
                            html += '</table>'
                            html += '</div>'
                            
                            st.markdown(html, unsafe_allow_html=True)

                        # --- EXPORT TO EXCEL ---
                        excel_data = generate_styled_excel(roster, sorted_days, unique_positions)
                        st.download_button(
                            label="📥 הורד סידור עבודה כמותאם (Excel)",
                            data=excel_data,
                            file_name="roster.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )

                        # --- SURPLUS REPORT (right after schedule) ---
                        surplus_data = results.get('surplus_report', {})
                        if surplus_data:
                            st.divider()
                            st.header("📊 דוח עודפים (מאבטחים זמינים שלא שובצו)")
                            st.info("המאבטחים הבאים סימנו זמינות אך לא שובצו לאף עמדה (כולן מלאות). ניתן להפנות אותם לגזרות/פרויקטים אחרים.")
                            
                            total_surplus = sum(len(emps) for emps in surplus_data.values())
                            st.metric("סה\"כ משבצות עודפות", total_surplus)
                            
                            for day, emps in surplus_data.items():
                                with st.expander(f"📅 {day} — {len(emps)} עודפים"):
                                    surplus_text = "\n".join([f"- **{emp_info['name']}** — זמין ל: {emp_info['shifts']}" for emp_info in emps])
                                    st.markdown(surplus_text)

                        # --- STATISTICS SECTION (Enhanced) ---
                        st.divider()
                        st.header("📊 דשבורד ניתוח וסטטיסטיקות")
                        
                        if not roster.empty:
                            # 1. Calculate Per-Employee Availability Count
                            # Iterate exactly as we did for the scheduler input to count marked "V"s
                            emp_availability = {}
                            
                            # Determine effective list of employees
                            active_emp_indices = [
                                idx for idx, r in st.session_state['employees_df'].iterrows() 
                                if str(r[name_col]) not in st.session_state.get('excluded_employees', set())
                            ]

                            for idx in active_emp_indices:
                                row = st.session_state['employees_df'].loc[idx]
                                ename = row[name_col]
                                
                                # Access override if exists
                                ov_data = collected_overrides.get(idx)
                                
                                # Count 'True's in availability (M, A, N, M_double, N_double)
                                total_marked = 0
                                shifts_to_use_dashboard = st.session_state.get('selected_shifts', potential_shifts)
                                
                                for s_col in shifts_to_use_dashboard:
                                    clean_d = str(s_col).strip()
                                    
                                    # Check Override
                                    if ov_data is not None and clean_d in ov_data.columns:
                                        # ov_data is a DataFrame with cols=Days, rows=Shifts
                                        # We need to sum the boolean trues for this day column
                                        try:
                                            # The override DF index is ['Morning', 'Afternoon', 'Night', 'DoubleM', 'DoubleN']
                                            # We just count how many checkmarks are True for this column
                                            day_vals = ov_data[clean_d]
                                            
                                            # LOGIC CHANGE: 
                                            # If employee marked ANY shift in a day (Morning AND/OR Night) -> Availability = 1
                                            # Because they can only be assigned once per day.
                                            if day_vals.any():
                                                total_marked += 1
                                        except:
                                            pass
                                    else:
                                        # Use File Data
                                        val = str(row[s_col]).lower()
                                        if 'nan' not in val:
                                            # Using keywords check implies availability.
                                            # If ANY keyword exists, it's a "Yes" for this day.
                                            is_avail_day = False
                                            if 'בוקר' in val or 'morning' in val: is_avail_day = True
                                            if 'צהריים' in val or 'afternoon' in val: is_avail_day = True
                                            if 'לילה' in val or 'night' in val: is_avail_day = True
                                            
                                            if is_avail_day:
                                                total_marked += 1
                                
                                emp_availability[ename] = total_marked

                            # 2. Build Analysis DataFrame
                            # Group roster by employee
                            roster_counts = roster['עובד'].value_counts().reset_index()
                            roster_counts.columns = ['עובד', 'שובץ_בפועל']
                            
                            # Merge with availability
                            analysis_df = pd.DataFrame(list(emp_availability.items()), columns=['עובד', 'זמינות_מוצהרת'])
                            analysis_df = analysis_df.merge(roster_counts, on='עובד', how='left').fillna(0)
                            analysis_df['שובץ_בפועל'] = analysis_df['שובץ_בפועל'].astype(int)
                            
                            # Calculate Utilization
                            # Avoid div by zero
                            analysis_df['אחוז_ניצול'] = analysis_df.apply(
                                lambda x: round((x['שובץ_בפועל'] / x['זמינות_מוצהרת'] * 100), 1) if x['זמינות_מוצהרת'] > 0 else 0,
                                axis=1
                            )
                            
                            # Sort by assigned shifts desc
                            analysis_df = analysis_df.sort_values('שובץ_בפועל', ascending=False)
                            
                            # --- TOP METRICS ---
                            m1, m2, m3, m4 = st.columns(4)
                            total_shifts_assigned = analysis_df['שובץ_בפועל'].sum()
                            avg_shifts = round(analysis_df['שובץ_בפועל'].mean(), 1)
                            total_shortages = sum(results.get('shortage_summary', {}).values())
                            
                            m1.metric("סה\"כ משמרות", total_shifts_assigned)
                            m2.metric("ממוצע לעובד", avg_shifts)
                            m3.metric("חוסרים", total_shortages, delta_color="inverse", delta=f"-{total_shortages}" if total_shortages > 0 else "0")
                            
                            # --- CHARTS ---
                            st.write("")
                            c_chart, c_table = st.columns([3, 2])
                            
                            with c_chart:
                                st.subheader("📉 שיבוץ מול זמינות (מי קיבל מה שביקש?)")
                                st.info("הגרף מציג לכל עובד: כמה משמרות סימן כפנוי (כחול) מול כמה קיבל בפועל (אדום).")
                                
                                # Reshape for Streamlit Bar Chart (Long format)
                                chart_data = analysis_df[['עובד', 'זמינות_מוצהרת', 'שובץ_בפועל']].set_index('עובד')
                                st.bar_chart(chart_data, color=["#e0e0e0", "#ff4b4b"], stack=False, height=520, use_container_width=True) # Gray for avail, Red for Actual
                            
                            with c_table:
                                st.subheader("📋 טבלת נתונים")
                                st.dataframe(
                                    analysis_df,
                                    use_container_width=True,
                                    column_config={
                                        "אחוז_ניצול": st.column_config.ProgressColumn(
                                            "אחוז ניצול",
                                            format="%.1f%%",
                                            min_value=0,
                                            max_value=100,
                                        ),
                                        "זמינות_מוצהרת": st.column_config.NumberColumn("זמינות (V)", help="כמות המשבצות שסומנו"),
                                        "שובץ_בפועל": st.column_config.NumberColumn("בפועל", help="משמרות ששובצו")
                                    },
                                    hide_index=True,
                                    height=520
                                )

                            # --- SHORTAGES ---
                            st.divider()
                            st.subheader("⚠️ ניתוח חוסרים")
                            shortages = results.get('shortage_summary', {})
                            
                            if shortages:
                                # Convert raw summary dict to structured DataFrame
                                s_rows = []
                                for label, count in shortages.items():
                                    # Label format: "Day|PosName|ShiftType"
                                    parts = label.split('|')
                                    if len(parts) >= 3:
                                        s_rows.append({
                                            "יום": parts[0],
                                            "עמדה": parts[1],
                                            "משמרת": parts[2],
                                            "כמות חסרה": count
                                        })
                                    else:
                                        # Fallback for unexpected format
                                        s_rows.append({"עמדה": label, "כמות חסרה": count})
                                
                                s_df = pd.DataFrame(s_rows)
                                
                                # View 1: Aggregated by Position (Bar Chart)
                                st.markdown("##### 1. סה\"כ חוסרים לפי עמדה")
                                pos_agg = s_df.groupby('עמדה')['כמות חסרה'].sum().reset_index()
                                st.bar_chart(pos_agg.set_index('עמדה'), color="#ffa600", horizontal=True)
                                
                                # View 2: Detailed Table
                                st.markdown("##### 2. פירוט חוסרים מלא")
                                st.dataframe(
                                    s_df, 
                                    use_container_width=True,
                                    hide_index=True,
                                    column_config={
                                        "כמות חסרה": st.column_config.NumberColumn("כמות חסרה", format="%d ❌")
                                    }
                                )
                            else:
                                st.success("כל העמדות מאוישות! אין חוסרים. 👏")
                            
                            # --- RECOMMENDATIONS ---
                            # --- RECOMMENDATIONS ---
                            recs_data = results.get('gap_recommendations', {})
                            
                            # Handle Legacy List Format (Safety)
                            if isinstance(recs_data, list):
                                # Fallback: Treat as one big unorganized list if something went wrong
                                if recs_data:
                                    st.divider()
                                    st.header("💡 המלצות לגישור פערים")
                                    st.warning("פורמט נתונים ישן זוהה. נסה להריץ מחדש את השיבוץ.")
                                    recs_text = "\n".join([f"👉 {r}" for r in recs_data])
                                    st.markdown(recs_text)
                            
                            # Standard Dict Format (Grouped by Shortage)
                            elif isinstance(recs_data, dict) and recs_data:
                                # Check if meaningful data exists
                                has_any_recs = any(d['available'] or d['potential'] for d in recs_data.values())
                                
                                if has_any_recs:
                                    st.divider()
                                    st.header("💡 המלצות לגישור פערים (Gap Filling)")
                                    
                                    for shortage_label, groups in recs_data.items():
                                        recs_avail = groups.get('available', [])
                                        recs_poten = groups.get('potential', [])
                                        
                                        if not recs_avail and not recs_poten:
                                            continue
                                            
                                        # Box/Container for each shortage
                                        with st.container():
                                            st.markdown(f"#### 🔎 עבור: {shortage_label}")
                                            
                                            # 1. Available Candidates
                                            if recs_avail:
                                                st.caption("✅ עובדים שסימנו זמינות:")
                                                avail_text = "\n".join([f"- {r}" for r in recs_avail])
                                                st.markdown(avail_text)
                                            
                                            # 2. Potential Candidates (Expander)
                                            if recs_poten:
                                                with st.expander(f"⚠️ הצג {len(recs_poten)} מועמדים פוטנציאליים (טכנית פנויים)"):
                                                    st.caption("לא סימנו זמינות, אך חוקית יכולים לעבוד:")
                                                    poten_text = "\n".join([f"- {r}" for r in recs_poten])
                                                    st.markdown(poten_text)
                                            
                                            st.divider()
                            

                        else:
                            st.info("אין נתונים להצגה (הרוסטר ריק)")

                    else:
                        st.error(f"לא נמצא פתרון: {results['status']}")
                        
                        # Show Diagnostics
                        if 'diagnostics' in results and results['diagnostics']:
                            st.warning("⚠️ המערכת זיהתה את הפערים הבאים (אילוצים שלא ניתן לקיים):\n\n" + "\n".join([f"- {warn}" for warn in results['diagnostics']]))
                        else:
                            st.info("טיפ: נסה להוריד את דרישות המאבטחים או לבטל את איסור החפיפות.")
                except Exception as e:
                    st.error(f"שגיאה בתהליך השיבוץ: {e}")

    else:
        st.error(f"שגיאה בקריאת הקובץ: {header_idx}")

# --- AUTOSAVE ON EVERY CHANGE ---
if st.session_state.get('firebase_loaded', False) and st.session_state.get('user_email'):
    firebase_manager.save_state_to_firebase(st.session_state, st.session_state['user_email'])
