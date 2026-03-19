import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import streamlit as st
import io
import json

def init_firebase():
    if not firebase_admin._apps:
        try:
            # Check if we are in Streamlit Cloud environment (secrets exist)
            try:
                if "firebase" in st.secrets:
                    # Convert taking the Streamlit Secrets dictionary and converting it to a standard dict
                    # because `credentials.Certificate` requires a dict, not a Secrets object
                    cert_dict = dict(st.secrets["firebase"])
                    cred = credentials.Certificate(cert_dict)
                else:
                    cred = credentials.Certificate("firebase_credentials.json")
            except FileNotFoundError:
                # streamlit raises FileNotFoundError when secrets.toml is not found
                cred = credentials.Certificate("firebase_credentials.json")
                
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.sidebar.warning(f"שגיאה בהתחברות ל-Firebase (האם הוספת את קובץ המפתח למערכת?): {e}")
            return None
    return firestore.client()

def save_state_to_firebase(session_state, user_id):
    db = init_firebase()
    if not db:
        return
    
    # Extract data to save
    data = {}
    
    if 'positions' in session_state:
        data['positions'] = session_state['positions']
    
    if 'constraints' in session_state:
        data['constraints'] = session_state['constraints']
        
    if 'deleted_positions' in session_state:
        data['deleted_positions'] = list(session_state['deleted_positions'])
        
    if 'excluded_employees' in session_state:
        data['excluded_employees'] = list(session_state['excluded_employees'])
        
    if 'col_map' in session_state:
        data['col_map'] = session_state['col_map']
        
    if 'selected_shifts' in session_state:
        data['selected_shifts'] = session_state['selected_shifts']
        
    if 'current_file_id' in session_state:
        data['current_file_id'] = session_state['current_file_id']
        
    if 'header_idx' in session_state:
        data['header_idx'] = session_state['header_idx']
        
    if 'employees_df' in session_state and session_state['employees_df'] is not None:
        data['employees_df'] = session_state['employees_df'].to_json(orient='records', force_ascii=False)

    # --- Capture Dynamic Widget States (User edits on UI) ---
    dynamic_state = {}
    
    # 1. Preserve previously loaded/accumulated edits (so hidden/skipped widgets aren't lost)
    if 'restored_edits' in session_state:
        for k, v in session_state['restored_edits'].items():
            dynamic_state[k] = v

    # 2. Merge current user interaction states
    for key in session_state.keys():
        if key.startswith("roles_sel_") or key.startswith("pref_") or key.startswith("max_s_") or key.startswith("fixed_shifts_list_"):
            dynamic_state[key] = session_state[key]
            
    data['dynamic_state_json'] = json.dumps(dynamic_state, default=str)

    if 'current_edited_displays' in session_state:
        emp_consts = {}
        for k, v in session_state['current_edited_displays'].items():
            if isinstance(v, pd.DataFrame):
                emp_consts[k] = v.to_json(orient='records', force_ascii=False)
        data['firebase_constraints_base'] = json.dumps(emp_consts, default=str)

    if 'latest_roster_results' in session_state:
        res = session_state['latest_roster_results']
        if res and isinstance(res, dict):
            res_to_save = dict(res)
            # Serialize the roster DataFrame if it exists
            if res_to_save.get('roster') is not None:
                res_to_save['roster'] = res_to_save['roster'].to_json(orient='records', force_ascii=False)
            data['latest_roster_results'] = json.dumps(res_to_save, default=str)

    try:
        db.collection('autoshift').document(user_id).set(data)
    except Exception as e:
        pass

def load_state_from_firebase(session_state, user_id):
    db = init_firebase()
    if not db:
        return False
        
    try:
        doc = db.collection('autoshift').document(user_id).get()
        if doc.exists:
            data = doc.to_dict()
            
            if 'positions' in data:
                session_state['positions'] = data['positions']
            if 'constraints' in data:
                session_state['constraints'] = data['constraints']
            if 'deleted_positions' in data:
                session_state['deleted_positions'] = set(data['deleted_positions'])
            if 'excluded_employees' in data:
                session_state['excluded_employees'] = set(data['excluded_employees'])
            if 'col_map' in data:
                session_state['col_map'] = data['col_map']
            if 'selected_shifts' in data:
                session_state['selected_shifts'] = data['selected_shifts']
            if 'current_file_id' in data:
                session_state['current_file_id'] = data['current_file_id']
            if 'header_idx' in data:
                session_state['header_idx'] = data['header_idx']
                
            if 'employees_df' in data and data['employees_df']:
                session_state['employees_df'] = pd.read_json(io.StringIO(data['employees_df']), orient='records')
                
            if 'dynamic_state_json' in data:
                dynamic_state = json.loads(data['dynamic_state_json'])
                for k, v in dynamic_state.items():
                    # Protect against legacy saves throwing ValueAssignmentNotAllowedError 
                    if not k.startswith("emp_edit_"):
                        session_state[k] = v

            if 'firebase_constraints_base' in data:
                try:
                    loaded_constraints = json.loads(data['firebase_constraints_base'])
                    parsed_constraints = {}
                    for k, v in loaded_constraints.items():
                        parsed_constraints[k] = pd.read_json(io.StringIO(v), dtype=str, orient='records')
                        # Convert string boolean-like values back to actual booleans (as JSON load natively might cast to object due to mixed types)
                        for col in parsed_constraints[k].columns:
                            if col != 'סוג משמרת':
                                parsed_constraints[k][col] = parsed_constraints[k][col].map({'True': True, 'False': False, True: True, False: False}).fillna(False)
                                
                    session_state['firebase_constraints_base'] = parsed_constraints
                except Exception as e:
                    pass

            if 'latest_roster_results' in data and data['latest_roster_results']:
                try:
                    res_loaded = json.loads(data['latest_roster_results'])
                    if res_loaded.get('roster'):
                        res_loaded['roster'] = pd.read_json(io.StringIO(res_loaded['roster']), orient='records')
                    session_state['latest_roster_results'] = res_loaded
                except Exception as e:
                    pass
                    
            return True
        return False
    except Exception as e:
        st.sidebar.error(f"שגיאה בשחזור נתונים מ-Firebase: {e}")
        return False

def delete_state_from_firebase(user_id):
    db = init_firebase()
    if not db:
        return False
        
    try:
        db.collection('autoshift').document(user_id).delete()
        st.sidebar.success("כל הנתונים אופסו ונמחקו מהענן! 🗑️")
        return True
    except Exception as e:
        st.sidebar.error(f"שגיאה במחיקת הנתונים: {e}")
        return False
