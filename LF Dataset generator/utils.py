import json
import pandas as pd
import numpy as np

def load_csv_robust(file_path):
    """
    Loads a CSV file with automatic encoding fallback.
    """
    encodings = ['utf-8', 'latin1', 'cp1252', 'utf-8-sig']
    for enc in encodings:
        try:
            return pd.read_csv(file_path, dtype=str, encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    # Fallback with replacement of un-decodable bytes
    return pd.read_csv(file_path, dtype=str, encoding='utf-8', errors='replace')



def normalize_value(val):
    """
    Strips whitespace and normalizes empty-like values (NaN, 'NULL', 'None', 'nan', etc.) to None.
    """
    if pd.isna(val) or val is None:
        return None
    
    # Handle numpy types
    if isinstance(val, (np.integer, np.floating)):
        return val
        
    val_str = str(val).strip()
    val_lower = val_str.lower()
    
    if val_lower in ["", "null", "none", "nan", "-", "undefined"]:
        return None
        
    return val_str

def resolve_column(df, aliases):
    """
    Checks df.columns for a match against the aliases list (case-insensitive).
    Returns the exact column name in df if found, otherwise None.
    """
    # Normalize DataFrame column names for comparison
    normalized_cols = {col.strip().lower(): col for col in df.columns}
    
    for alias in aliases:
        norm_alias = alias.strip().lower()
        if norm_alias in normalized_cols:
            return normalized_cols[norm_alias]
            
    return None

def parse_and_format_json(val):
    """
    Attempts to parse a string as JSON. If successful, returns the formatted JSON string.
    If it's already a dictionary or list, dumps it formatted.
    If parsing fails, returns the cleaned string.
    """
    val_norm = normalize_value(val)
    if val_norm is None:
        return None
        
    if isinstance(val_norm, (dict, list)):
        return json.dumps(val_norm, indent=2, ensure_ascii=False)
        
    # Attempt to load string as JSON
    try:
        parsed = json.loads(val_norm)
        return json.dumps(parsed, indent=2, ensure_ascii=False)
    except (json.JSONDecodeError, TypeError):
        return val_norm

def construct_input_column(call_json, inferred_details, transcript):
    """
    Concatenates the three required input sections in the exact order with separators.
    """
    # Ensure they are formatted/parsed JSONs where applicable
    formatted_call_json = parse_and_format_json(call_json)
    formatted_inferred = parse_and_format_json(inferred_details)
    
    sections = [
        f"Call JSON\n------------------\n{formatted_call_json}",
        f"Inferred Details\n------------------\n{formatted_inferred}",
        f"Transcript\n------------------\n{normalize_value(transcript)}"
    ]
    
    return "\n\n".join(sections)


def construct_expected_output(row, mapping):
    """
    Constructs the expected output dict from the row using mapping configurations.
    Only includes fields that are actually available (not empty/None).
    Returns the JSON string representation.
    """
    expected_data = {}
    for key, aliases in mapping.items():
        col_name = resolve_column(pd.DataFrame(columns=row.index), aliases)
        if col_name:
            val = normalize_value(row[col_name])
            if val is not None:
                # Try to parse numeric or boolean values, or json values if appropriate
                # For scores, try to convert to float/int if possible
                if "score" in key:
                    try:
                        # check if it is float or integer
                        if float(val).is_integer():
                            expected_data[key] = int(float(val))
                        else:
                            expected_data[key] = float(val)
                        continue
                    except ValueError:
                        pass
                expected_data[key] = val
                
    return json.dumps(expected_data, indent=2, ensure_ascii=False)

def construct_metadata(call_id):
    """
    Constructs the metadata dict and returns the JSON string.
    """
    meta = {
        "call_id": str(call_id)
    }
    return json.dumps(meta, ensure_ascii=False)
