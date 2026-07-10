import os

# Base Directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Input File Paths
CALLS_CSV_PATH = os.path.join(INPUT_DIR, "calls.csv")
MANUAL_AUDIT_CSV_PATH = os.path.join(INPUT_DIR, "manual_audit.csv")
CALL_JSON_CSV_PATH = os.path.join(INPUT_DIR, "call_jsons.csv")

# Output File Paths
LANGFUSE_DATASET_CSV_PATH = os.path.join(OUTPUT_DIR, "langfuse_dataset.csv")
GENERATION_REPORT_TXT_PATH = os.path.join(OUTPUT_DIR, "generation_report.txt")

# Required columns in calls.csv
CALLS_REQUIRED_COLS = {
    "call_id": ["Call ID", "call_id", "CallID"],
    "inferred_details": ["inferred_details", "Inferred Details", "InferredDetails"]
}

# Required columns in manual_audit.csv
MANUAL_AUDIT_REQUIRED_COLS = {
    "call_id": ["Call ID", "call_id", "CallID"],
    "transcript": ["Transcript", "transcript", "Transcript Details"]
}

# Required columns in call_jsons.csv
CALL_JSON_REQUIRED_COLS = {
    "call_id": ["Call ID", "call_id", "CallID"],
    "call_json": ["call_json", "Call JSON", "CallJSON"]
}

# Mapping of Expected Output JSON fields to potential columns in manual_audit.csv
EXPECTED_OUTPUT_FIELDS_MAPPING = {
    "buyer_name": ["buyer_name", "Buyer Name", "BuyerName", "Ai: Buyer Name"],
    "buyer_location": ["buyer_location", "Buyer Location", "BuyerLocation", "Ai: Buyer Location"],
    "buyer_contact": ["buyer_contact", "Buyer Contact", "BuyerContact", "Buyer Contact Details", "Ai: Buyer Contact Details"],
    "product_name": ["product_name", "Product Name", "ProductName", "Ai: Product Name"],
    "specifications": ["specifications", "Specifications", "Specification", "Ai: Specifications"],
    "buyer_required_quantity": ["buyer_required_quantity", "Buyer Required Quantity", "Quantity", "Ai: Buyer Required Quantity"],
    "price_quoted": ["price_quoted", "Price Quoted", "Price Quoted by Seller", "Price", "Ai: Price Quoted by Seller"],
    "seller_actionables": ["seller_actionables", "Seller Actionables", "Actionables of the Seller", "Actionables", "Ai: Actionables of the Seller"],
    "lead_tag": ["lead_tag", "Lead Tag", "LeadTag", "Ai: Lead Tag"],
    "buyer_intent": ["buyer_intent", "Buyer Intent", "BuyerIntent", "Ai: Buyer Intent"],
    "buyer_questions": ["buyer_questions", "Buyer Questions", "BuyerQuestions", "Ai: Buyer Questions"],
    "reminder": ["reminder", "Reminder", "Ai: Reminder"],
    "verdict": ["verdict", "Verdict"],
    "category": ["category", "Category"],
    "reasoning": ["reasoning", "Reasoning"],
    "accuracy_score": ["accuracy_score", "Accuracy Score", "AccuracyScore"],
    "completeness_score": ["completeness_score", "Completeness Score", "CompletenessScore"]
}
