import os
import pandas as pd
import numpy as np

# Configurable file paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")

# Paths for calls.csv and ai_audit_data.csv
# In typical runs, you can customize these paths to point to your target directory
CALLS_CSV_PATH = os.path.join(INPUT_DIR, "calls.csv")
AI_AUDIT_CSV_PATH = os.path.join(INPUT_DIR, "ai_audit_data.csv")

def is_empty(val):
    """
    Checks if a cell value is empty, NaN, or whitespace.
    """
    if pd.isna(val) or val is None:
        return True
    val_str = str(val).strip()
    return val_str in ["", "nan", "NaN", "NULL", "None", "<NA>"]

def main():
    # Verify input files exist
    if not os.path.exists(CALLS_CSV_PATH):
        print(f"Error: Source of truth file not found at: {CALLS_CSV_PATH}")
        return
    if not os.path.exists(AI_AUDIT_CSV_PATH):
        print(f"Error: Audit data file not found at: {AI_AUDIT_CSV_PATH}")
        return

    print("Loading data...")
    # Load files. Force string datatype to preserve ID formatting
    calls_df = pd.read_csv(CALLS_CSV_PATH, dtype=str)
    ai_audit_df = pd.read_csv(AI_AUDIT_CSV_PATH, dtype=str)

    print(f"Loaded {len(calls_df)} rows from calls.csv.")
    print(f"Loaded {len(ai_audit_df)} rows from ai_audit_data.csv.")

    # Standardize and clean key columns in calls.csv
    calls_df['Seller GLID'] = calls_df['Seller GLID'].fillna("").astype(str).str.strip()
    calls_df['Caller GLID'] = calls_df['Caller GLID'].fillna("").astype(str).str.strip()
    calls_df['Call ID'] = calls_df['Call ID'].fillna("").astype(str).str.strip()

    # Filter out empty pairs in calls.csv
    valid_calls = calls_df[
        (calls_df['Seller GLID'] != "") &
        (calls_df['Caller GLID'] != "") &
        (calls_df['Call ID'] != "")
    ]

    # Find the count of matches for each (Seller GLID, Caller GLID) pair
    match_counts = valid_calls.groupby(['Seller GLID', 'Caller GLID']).size().reset_index(name='count')

    # Filter to pairs that occur EXACTLY once in calls.csv
    unique_pairs = match_counts[match_counts['count'] == 1]

    # Merge back to get the corresponding Call ID
    unique_matches = pd.merge(unique_pairs, valid_calls, on=['Seller GLID', 'Caller GLID'], how='left')

    # Create a lookup dictionary mapping (Seller GLID, Caller GLID) -> Call ID
    lookup_dict = dict(zip(
        zip(unique_matches['Seller GLID'], unique_matches['Caller GLID']),
        unique_matches['Call ID']
    ))

    print(f"Identified {len(lookup_dict)} unique (Seller GLID, Caller GLID) mapping pairs from calls.csv.")

    # Counters
    total_rows = len(ai_audit_df)
    missing_before = 0
    restored = 0
    still_missing = 0

    # Process ai_audit_data.csv row by row
    for idx, row in ai_audit_df.iterrows():
        call_id = row.get('Call ID')
        
        # Check if Call ID is missing/empty
        if is_empty(call_id):
            missing_before += 1
            
            # Normalize the lookup keys
            seller_glid = str(row.get('Seller GLID', '')).strip() if pd.notna(row.get('Seller GLID')) else ""
            caller_glid = str(row.get('Caller GLID', '')).strip() if pd.notna(row.get('Caller GLID')) else ""
            
            # Retrieve from lookup dictionary
            pair = (seller_glid, caller_glid)
            if pair in lookup_dict:
                ai_audit_df.at[idx, 'Call ID'] = lookup_dict[pair]
                restored += 1
            else:
                still_missing += 1

    # Overwrite the original ai_audit_data.csv file
    try:
        ai_audit_df.to_csv(AI_AUDIT_CSV_PATH, index=False)
        print("Updated audit file saved successfully.")
    except Exception as e:
        print(f"Error saving updated CSV file: {e}")
        return

    # Print summary of results
    print("\n=== Call ID Restoration Summary ===")
    print(f"Total rows: {total_rows}")
    print(f"Missing Call IDs before: {missing_before}")
    print(f"Successfully restored: {restored}")
    print(f"Still missing: {still_missing}")
    print("===================================\n")

if __name__ == "__main__":
    main()
