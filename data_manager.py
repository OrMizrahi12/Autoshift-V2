import pandas as pd
import io

def load_data(file_buffer):
    """
    Parses the uploaded Excel file.
    Searches for the header row containing 'עובדים' (Employees) or 'Name'.
    Returns the cleaned DataFrame and the header row index.
    """
    try:
        # Read the file without header first to locate the data
        df_raw = pd.read_excel(file_buffer, header=None)
        
        header_row_idx = None
        # Keyword to identify header row
        target_keyword = "עובדים"
        
        # Iterate through first 20 rows to find the header
        for i, row in df_raw.head(20).iterrows():
            # Check if any cell in the row contains the keyword
            row_str = row.astype(str).str.contains(target_keyword, na=False)
            if row_str.any():
                header_row_idx = i
                break
        
        if header_row_idx is None:
            # Fallback if specific Hebrew keyword not found, try English
            target_keyword = "Name"
            for i, row in df_raw.head(20).iterrows():
                row_str = row.astype(str).str.contains(target_keyword, case=False, na=False)
                if row_str.any():
                    header_row_idx = i
                    break
        
        # If still not found, default to 0 but warn
        if header_row_idx is None:
            return None, "לא נמצאה שורת כותרת (חיפשתי: 'עובדים' או 'Name')"

        # Read actual data with correct header
        file_buffer.seek(0)
        df = pd.read_excel(file_buffer, header=header_row_idx)
        
        # Clean column names (strip whitespace)
        df.columns = df.columns.astype(str).str.strip()
        
        # Locate the specific 'Name' column to drop invalid rows
        name_col = None
        for col in df.columns:
            if "עובדים" in col or "Name" in col:
                name_col = col
                break
        
        if name_col:
            # Remove rows where Name is missing (empty lines, footers)
            df = df.dropna(subset=[name_col])
        
        # Remove completely empty rows/cols
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        
        return df, header_row_idx

    except Exception as e:
        return None, str(e)

def get_shift_columns(df):
    """
    Identifies column names that represent shifts/dates.
    Heuristic: Columns that contain '/' or are not metadata (Name, Position, etc).
    """
    exclude = ["עובדים", "תפקידים", "הערות", "Name", "Position", "Comments", "Role"]
    shift_cols = []
    for col in df.columns:
        if col not in exclude and ("/" in str(col) or "-" in str(col)):
            shift_cols.append(col)
    return shift_cols
