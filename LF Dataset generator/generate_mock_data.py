import os
import pandas as pd

# Paths
INPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input")
os.makedirs(INPUT_DIR, exist_ok=True)

calls_path = os.path.join(INPUT_DIR, "calls.csv")
manual_path = os.path.join(INPUT_DIR, "manual_audit.csv")
call_jsons_path = os.path.join(INPUT_DIR, "call_jsons.csv")

# Create calls.csv data
calls_data = [
    # 1. Valid record (present in both, all fields valid)
    {
        "Call ID": "1001",
        "Call Summary": "Buyer Alice from Mumbai wants to buy Aluminium Sheets.",
        "call_json": '{"buyer_name": "Alice", "product": "Aluminium Sheets"}',
        "inferred_details": '{"buyer_location": "Mumbai", "buyer_contact": "9876543210"}',
        "other_metadata": "some_meta_1"
    },
    # 2. Duplicate of 1001 to test duplicate removal
    {
        "Call ID": "1001",
        "Call Summary": "DUPLICATE Summary.",
        "call_json": '{"buyer_name": "Alice-Dup"}',
        "inferred_details": '{"buyer_location": "Mumbai"}',
        "other_metadata": "some_meta_dup"
    },
    # 3. Missing transcript (transcript missing in manual_audit.csv)
    {
        "Call ID": "1002",
        "Call Summary": "Buyer Bob wants copper wire.",
        "call_json": '{"buyer_name": "Bob"}',
        "inferred_details": '{"buyer_location": "Delhi"}',
        "other_metadata": "some_meta_2"
    },
    # 4. Missing call_json in calls.csv (but provided in call_jsons.csv)
    {
        "Call ID": "1003",
        "Call Summary": "Buyer Charlie wants steel rods.",
        "call_json": "",  # Empty here, should fallback to call_jsons.csv
        "inferred_details": '{"buyer_location": "Pune"}',
        "other_metadata": "some_meta_3"
    },
    # 5. Missing inferred_details
    {
        "Call ID": "1004",
        "Call Summary": "Buyer Dave wants iron pipes.",
        "call_json": '{"buyer_name": "Dave"}',
        "inferred_details": "",  # Normalized to None
        "other_metadata": "some_meta_4"
    },
    # 6. Missing Call Summary (No longer required! Should succeed and be in output)
    {
        "Call ID": "1005",
        "Call Summary": "  ",
        "call_json": '{"buyer_name": "Eve"}',
        "inferred_details": '{"buyer_location": "Chennai"}',
        "other_metadata": "some_meta_5"
    },
    # 7. Missing manual audit (all manual audit fields empty in manual_audit.csv)
    {
        "Call ID": "1006",
        "Call Summary": "Buyer Frank wants glass cups.",
        "call_json": '{"buyer_name": "Frank"}',
        "inferred_details": '{"buyer_location": "Kolkata"}',
        "other_metadata": "some_meta_6"
    },
    # 8. Call ID only in calls.csv
    {
        "Call ID": "2001",
        "Call Summary": "Only in calls.",
        "call_json": '{"buyer_name": "George"}',
        "inferred_details": '{"buyer_location": "Goa"}',
        "other_metadata": "some_meta_7"
    }
]

# Create call_jsons.csv data
call_jsons_data = [
    # Call ID 1003 is missing call_json in calls.csv, so we resolve it from here:
    {"Call ID": "1003", "call_json": '{"buyer_name": "Charlie", "product": "steel rods"}'},
    # Call ID 1001 is already present in calls.csv (duplicate or match)
    {"Call ID": "1001", "call_json": '{"buyer_name": "Alice", "product": "Aluminium Sheets"}'},
]

# Create manual_audit.csv data
manual_data = [
    # 1. Valid record (matches 1001)
    {
        "Call ID": "1001",
        "Transcript": "Hello, I am Alice. I am from Mumbai. I want to buy Aluminium Sheets.",
        "Ai: Buyer Name": "Alice",
        "Ai: Buyer Location": "Mumbai",
        "Buyer Contact Details": "9876543210",
        "Product Name": "Aluminium Sheets",
        "Verdict": "Yes",
        "Accuracy Score": "100.0",
        "Completeness Score": "90",
        "Reasoning": "Matches all details discussed in call."
    },
    # 2. Missing transcript (matches 1002)
    {
        "Call ID": "1002",
        "Transcript": "NaN",  # Missing transcript
        "Ai: Buyer Name": "Bob",
        "Verdict": "Yes"
    },
    # 3. Missing call_json in calls.csv (matches 1003, but fallback resolves it)
    {
        "Call ID": "1003",
        "Transcript": "This is Charlie calling. I need steel rods.",
        "Ai: Buyer Name": "Charlie",
        "Verdict": "No"
    },
    # 4. Missing inferred_details (matches 1004)
    {
        "Call ID": "1004",
        "Transcript": "I am Dave. I need iron pipes.",
        "Ai: Buyer Name": "Dave",
        "Verdict": "Yes"
    },
    # 5. Missing Call Summary (matches 1005 - Call Summary is no longer required)
    {
        "Call ID": "1005",
        "Transcript": "I am Eve. I need PVC pipes.",
        "Ai: Buyer Name": "Eve",
        "Verdict": "Yes"
    },
    # 6. Missing manual audit (matches 1006 - but manual audit fields are empty)
    {
        "Call ID": "1006",
        "Transcript": "Hello, I am Frank. Let me tell you what I need...",
        # No audit fields populated
        "Ai: Buyer Name": "",
        "Ai: Buyer Location": None,
        "Verdict": "NaN",
        "Reasoning": "None"
    },
    # 7. Call ID only in manual_audit.csv
    {
        "Call ID": "3001",
        "Transcript": "Only in manual audit.",
        "Ai: Buyer Name": "Henry",
        "Verdict": "Yes"
    }
]

pd.DataFrame(calls_data).to_csv(calls_path, index=False)
pd.DataFrame(call_jsons_data).to_csv(call_jsons_path, index=False)
pd.DataFrame(manual_data).to_csv(manual_path, index=False)

print("Mock files created successfully:")
print(f"- calls.csv: {len(calls_data)} rows written.")
print(f"- call_jsons.csv: {len(call_jsons_data)} rows written.")
print(f"- manual_audit.csv: {len(manual_data)} rows written.")
