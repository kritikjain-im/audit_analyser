import os
import pandas as pd

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
os.makedirs(INPUT_DIR, exist_ok=True)

calls_path = os.path.join(INPUT_DIR, "calls.csv")
ai_audit_path = os.path.join(INPUT_DIR, "ai_audit_data.csv")

# Create calls.csv (source of truth)
calls_data = [
    # Case 1: Unique match (Seller GLID = 111, Caller GLID = 222 -> Call ID = 101)
    {"Call ID": "101", "Seller GLID": "111", "Caller GLID": "222"},
    # Case 2: Duplicated match in calls.csv (Seller GLID = 333, Caller GLID = 444 -> Ambiguous)
    {"Call ID": "102", "Seller GLID": "333", "Caller GLID": "444"},
    {"Call ID": "103", "Seller GLID": "333", "Caller GLID": "444"},
    # Case 3: Already present in ai_audit_data.csv, has a different ID in calls.csv
    {"Call ID": "104", "Seller GLID": "777", "Caller GLID": "888"},
]

# Create ai_audit_data.csv (missing IDs to be restored)
ai_audit_data = [
    # Row 1: Missing Call ID, has a unique match in calls.csv (Expectation: 101)
    {"Call ID": "", "Seller GLID": "111", "Caller GLID": "222", "Audit Column": "Audit_Result_A"},
    # Row 2: Missing Call ID, has duplicate matches in calls.csv (Expectation: remains empty)
    {"Call ID": "NaN", "Seller GLID": "333", "Caller GLID": "444", "Audit Column": "Audit_Result_B"},
    # Row 3: Missing Call ID, has no matches in calls.csv (Expectation: remains empty)
    {"Call ID": None, "Seller GLID": "555", "Caller GLID": "666", "Audit Column": "Audit_Result_C"},
    # Row 4: Call ID already present (Expectation: Stays "999")
    {"Call ID": "999", "Seller GLID": "777", "Caller GLID": "888", "Audit Column": "Audit_Result_D"},
]

# Save mock data
pd.DataFrame(calls_data).to_csv(calls_path, index=False)
pd.DataFrame(ai_audit_data).to_csv(ai_audit_path, index=False)

print("Mock retrieval data files generated successfully:")
print(f"- calls.csv: {len(calls_data)} rows written to {calls_path}")
print(f"- ai_audit_data.csv: {len(ai_audit_data)} rows written to {ai_audit_path}")
