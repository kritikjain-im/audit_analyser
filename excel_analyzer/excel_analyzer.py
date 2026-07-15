import os
import sys
import argparse
import glob
import logging
import re
import pandas as pd

# Add the current folder to sys.path so we can import excel_writer
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from excel_writer import create_excel_report

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

FIELDS = [
    "buyer_name",
    "buyer_location",
    "buyer_contact",
    "lead_tag",
    "product_name",
    "specifications",
    "quantity",
    "price",
    "actionables",
    "buyer_intent",
    "buyer_questions",
    "reminder"
]

FIELD_TARGETS = {
    "buyer_name": {"buyername", "buyer name"},
    "buyer_location": {"buyerlocation", "buyer location"},
    "buyer_contact": {"buyercontact", "buyer contact", "buyer contact details", "contact"},
    "lead_tag": {"leadtag", "lead tag"},
    "product_name": {"productname", "product name"},
    "specifications": {"specifications", "specification"},
    "quantity": {"quantity", "buyerrequiredquantity", "buyer required quantity"},
    "price": {"price", "pricequotedbyseller", "price quoted by seller"},
    "actionables": {"actionables", "selleractionables", "seller actionables"},
    "buyer_intent": {"buyerintent", "buyer intent"},
    "buyer_questions": {"buyerquestions", "buyer questions"},
    "reminder": {"reminder"}
}

def normalize_badge(text: str) -> str:
    """Normalize input badge text based on common patterns."""
    text = text.upper().strip()
    
    # 0. Check for "inferred" first
    if any(k in text for k in ["INFERRED", "INFER", "INF"]):
        return "Inferred"
    
    # 1. Check for "not discussed" variants next
    if any(k in text for k in ["NOT DISCUSSED", "NDT DISCUSSED", "NOT DISCVSSED", "NDT DISCVSSED", "MISCVSSF"]):
        return "Correct (Not Discussed)"
        
    # 2. Check for Incorrect
    if any(k in text for k in ["INCORRECT", "INCO", "JNCO", "JNC", "INC", "NCR", "ICR"]):
        return "Incorrect"
        
    # 3. Check for Correct
    if any(k in text for k in ["CORRECT", "COAR", "COPP", "COER", "COPPECT", "COA", "COE", "CORR", "CDRRECT", "CDRR", "CDRE", "OPREIT", "OPRE", "IORRFIT", "IORR", "IRR", "ORR", "ORRFIT"]):
        return "Correct"
        
    # 4. Check for Missing
    if any(k in text for k in ["MISSING", "MISS", "MIS", "MSS", "MSI"]):
        return "Missing"
        
    # 5. Check for Hallucination
    if any(k in text for k in ["HALLUCINATION", "HALL", "HAL", "LUCI", "HAC"]):
        return "Hallucination"
        
    return None

def normalize_value(val) -> str:
    """Clean and return a normalized string for reporting."""
    if pd.isna(val) or val is None:
        return "Correct (Not Discussed)"
    
    val_str = str(val).strip()
    if val_str == "" or val_str.lower() in ["none", "null", "-", "nan"]:
        return "Correct (Not Discussed)"
        
    norm = normalize_badge(val_str)
    if norm:
        return norm
        
    val_upper = val_str.upper()
    if "NOT DISCUSSED" in val_upper:
        return "Correct (Not Discussed)"
    if "INFERRED" in val_upper:
        return "Inferred"
    if "INCORRECT" in val_upper:
        return "Incorrect"
    if "CORRECT" in val_upper:
        return "Correct"
    if "MISSING" in val_upper:
        return "Missing"
    if "HALLUCINATION" in val_upper:
        return "Hallucination"
        
    return "Correct"

def get_match_status(manual_val: str, ai_val: str) -> int:
    """Match status check mirroring main.py."""
    if manual_val is None:
        return 1
    m = str(manual_val).strip().lower()
    a = str(ai_val).strip().lower()
    if m == a:
        return 1
        
    correct_variants = {"correct", "correct (not discussed)"}
    # Case 1: Manual correct & AI correct (not discussed) or vice versa -> Match
    if m in correct_variants and a in correct_variants:
        return 1
        
    # Case 2: Manual inferred & AI correct/correct (not discussed) or vice versa -> Match
    if (m == "inferred" and a in correct_variants) or (a == "inferred" and m in correct_variants):
        return 1
        
    return 0

def map_columns(headers, fields_targets, sheet_name="Sheet"):
    """Maps columns using alphanumeric normalization with fallback to substring."""
    mapping = {}
    normalized_headers = {}
    
    for h in headers:
        if not isinstance(h, str) and not isinstance(h, (int, float)):
            continue
        h_str = str(h)
        clean_h = re.sub(r'[^a-zA-Z0-9]', '', h_str).lower()
        if clean_h:
            normalized_headers[clean_h] = h_str
            
    # Map Call ID
    call_id_col = None
    for target in {"callid", "call id"}:
        clean_t = re.sub(r'[^a-zA-Z0-9]', '', target).lower()
        if clean_t in normalized_headers:
            call_id_col = normalized_headers[clean_t]
            break
            
    if not call_id_col:
        for clean_h, orig_h in normalized_headers.items():
            if "callid" in clean_h:
                call_id_col = orig_h
                break
                
    mapping["call_id"] = call_id_col
    
    # Map Auditor Name
    auditor_col = None
    for target in {"auditor", "manualauditor", "manual auditor", "auditedby", "audited by"}:
        clean_t = re.sub(r'[^a-zA-Z0-9]', '', target).lower()
        if clean_t in normalized_headers:
            auditor_col = normalized_headers[clean_t]
            break
    if not auditor_col:
        for clean_h, orig_h in normalized_headers.items():
            if "auditor" in clean_h or "audited" in clean_h:
                auditor_col = orig_h
                break
    mapping["auditor"] = auditor_col

    # Map the 12 Fields
    for field, targets in fields_targets.items():
        matched_header = None
        for t in targets:
            clean_t = re.sub(r'[^a-zA-Z0-9]', '', t).lower()
            if clean_t in normalized_headers:
                matched_header = normalized_headers[clean_t]
                break
        
        if matched_header:
            mapping[field] = matched_header
        else:
            for clean_h, orig_h in normalized_headers.items():
                for t in targets:
                    clean_t = re.sub(r'[^a-zA-Z0-9]', '', t).lower()
                    if clean_t in clean_h and not any(sub in clean_h for sub in ["reason", "sublabel", "subfield"]):
                        matched_header = orig_h
                        break
                if matched_header:
                    break
            mapping[field] = matched_header
                
    return mapping

def get_next_filename(base_dir, base_name="Compact_Analysis"):
    version = 1
    while True:
        candidate = os.path.join(base_dir, f"{base_name}_v{version}.xlsx")
        if not os.path.exists(candidate):
            return candidate
        version += 1

def find_verdict_col(df):
    for col in df.columns:
        if str(col).lower().strip() == 'verdict':
            return col
    for col in df.columns:
        if 'verdict' in str(col).lower():
            return col
    return None

def compute_compact_summary(df, mapping):
    """Computes category counts and accuracies for the 12 fields."""
    field_display_names = {
        "buyer_name": "Buyer Name",
        "buyer_location": "Buyer Location",
        "buyer_contact": "Buyer Contact Details",
        "lead_tag": "Lead Tag",
        "product_name": "Product Name",
        "specifications": "Specifications",
        "quantity": "Buyer Required Quantity",
        "price": "Price Quoted by Seller",
        "actionables": "Seller Actionables",
        "buyer_intent": "Buyer Intent",
        "buyer_questions": "Buyer Questions",
        "reminder": "Reminder"
    }
    
    results = []
    for field in FIELDS:
        display_name = field_display_names.get(field, field)
        col_name = mapping.get(field)
        
        counts = {
            "Correct": 0,
            "Correct (ND)": 0,
            "Inferred": 0,
            "Incorrect": 0,
            "Missing": 0,
            "Hallucination": 0
        }
        
        if col_name and col_name in df.columns:
            for val in df[col_name]:
                norm_val = normalize_value(val)
                if norm_val == "Correct (Not Discussed)":
                    norm_val = "Correct (ND)"
                if norm_val in counts:
                    counts[norm_val] += 1
                    
        field_total = sum(counts.values())
        cum_correct = counts["Correct"] + counts["Correct (ND)"] + counts["Inferred"]
        average = cum_correct / field_total if field_total > 0 else 0.0
        
        results.append({
            "Field Name": display_name,
            "Cumulative Correct": cum_correct,
            "Correct": counts["Correct"],
            "Correct (ND)": counts["Correct (ND)"],
            "Inferred": counts["Inferred"],
            "Missing": counts["Missing"],
            "Incorrect": counts["Incorrect"],
            "Hallucination": counts["Hallucination"],
            "Total Calls": field_total,
            "Average": average
        })
        
    sum_cum_correct = sum(r["Cumulative Correct"] for r in results)
    sum_total_calls = sum(r["Total Calls"] for r in results)
    overall_accuracy = sum_cum_correct / sum_total_calls if sum_total_calls > 0 else 0.0
    
    overall_dict = {
        "total_correct": sum_cum_correct,
        "total_calls_fields": sum_total_calls,
        "overall_accuracy": overall_accuracy
    }
    
    return results, overall_dict

def main():
    parser = argparse.ArgumentParser(description="Direct Call Audit Comparison Analyser")
    parser.add_argument("--input-file", "-i", help="Path to input Excel workbook or Manual Audits CSV")
    parser.add_argument("--output-dir", "-o", default="output", help="Directory to save output reports")
    parser.add_argument("--manual-sheet", default=None, help="Name of manual audits sheet (if Excel)")
    parser.add_argument("--ai-sheet", default=None, help="Name of AI audits sheet or path to AI Audits CSV")
    
    args = parser.parse_args()
    
    # Setup default paths if not provided
    input_path = args.input_file
    if not input_path:
        default_manual = os.path.join("input", "Manual Audits.csv")
        default_excel = os.path.join("input", "Langfuse Datasets.xlsx")
        if os.path.exists(default_manual):
            input_path = default_manual
        elif os.path.exists(default_excel):
            input_path = default_excel
        else:
            logger.error("No input file specified or default found.")
            sys.exit(1)
            
    ai_sheet_or_file = args.ai_sheet
    if not ai_sheet_or_file and input_path.lower().endswith('.csv'):
        default_ai = os.path.join("input", "AI Audits.csv")
        if os.path.exists(default_ai):
            ai_sheet_or_file = default_ai
            
    logger.info(f"Using Manual Input File: {input_path}")
    if ai_sheet_or_file:
        logger.info(f"Using AI Input File/Sheet: {ai_sheet_or_file}")
        
    os.makedirs(args.output_dir, exist_ok=True)
    
    # 2. Load DataFrames
    df_manual = None
    df_ai = None
    
    is_csv = input_path.lower().endswith('.csv')
    
    if is_csv:
        try:
            df_manual = pd.read_csv(input_path, encoding='utf-8')
        except UnicodeDecodeError:
            df_manual = pd.read_csv(input_path, encoding='latin1')
            
        if ai_sheet_or_file and ai_sheet_or_file.lower().endswith('.csv'):
            try:
                df_ai = pd.read_csv(ai_sheet_or_file, encoding='utf-8')
            except UnicodeDecodeError:
                df_ai = pd.read_csv(ai_sheet_or_file, encoding='latin1')
        else:
            logger.error("AI CSV file must be provided via --ai-sheet when manual file is CSV.")
            sys.exit(1)
    else:
        # Load from Excel sheets
        try:
            xls = pd.ExcelFile(input_path)
            sheet_names = xls.sheet_names
            
            man_sheet = args.manual_sheet or "Manual Audits"
            if man_sheet not in sheet_names:
                man_sheet = sheet_names[0]
                
            ai_sheet = args.ai_sheet or "AI Audits"
            if ai_sheet not in sheet_names:
                ai_sheet = sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]
                
            df_manual = pd.read_excel(input_path, sheet_name=man_sheet)
            df_ai = pd.read_excel(input_path, sheet_name=ai_sheet)
        except Exception as e:
            logger.error(f"Failed to read Excel workbook: {str(e)}")
            sys.exit(1)
            
    # Clean Verdict columns and filter to 'yes' or 'no'
    verdict_man_col = find_verdict_col(df_manual)
    if verdict_man_col:
        df_manual = df_manual[df_manual[verdict_man_col].astype(str).str.strip().str.lower().isin(['yes', 'no'])]
    else:
        logger.warning("Could not find Verdict column in Manual dataset. Skipping filtering.")
        
    verdict_ai_col = find_verdict_col(df_ai)
    if verdict_ai_col:
        df_ai = df_ai[df_ai[verdict_ai_col].astype(str).str.strip().str.lower().isin(['yes', 'no'])]
    else:
        logger.warning("Could not find Verdict column in AI dataset. Skipping filtering.")
        
    # Map columns
    manual_mapping = map_columns(df_manual.columns, FIELD_TARGETS, "Manual")
    ai_mapping = map_columns(df_ai.columns, FIELD_TARGETS, "AI")
    
    manual_call_id_col = manual_mapping["call_id"]
    ai_call_id_col = ai_mapping["call_id"]
    
    if not manual_call_id_col or not ai_call_id_col:
        logger.error("Could not map Call ID columns in one or both datasets.")
        sys.exit(1)
        
    df_manual = df_manual.dropna(subset=[manual_call_id_col])
    df_ai = df_ai.dropna(subset=[ai_call_id_col])
    
    # Calculate Sheet 1: AI Audits Summary
    summary_ai, overall_ai = compute_compact_summary(df_ai, ai_mapping)
    
    # Calculate Sheet 2: Manual Audits Summary
    summary_manual, overall_manual = compute_compact_summary(df_manual, manual_mapping)
    
    # Calculate Sheet 3: AI Matched Summary (AI audits for manual call IDs)
    manual_cids = set(df_manual[manual_call_id_col].astype(str).str.split('.').str[0].str.strip())
    df_ai_matched = df_ai[df_ai[ai_call_id_col].astype(str).str.split('.').str[0].str.strip().isin(manual_cids)]
    summary_ai_matched, overall_ai_matched = compute_compact_summary(df_ai_matched, ai_mapping)
    
    # Extract AI/Manual calls dictionary for comparison (similar to old comparison code)
    manual_calls = {}
    manual_auditor_col = manual_mapping.get("auditor")
    for _, row in df_manual.iterrows():
        call_id = str(row[manual_call_id_col]).split('.')[0].strip()
        auditor = str(row[manual_auditor_col]).strip() if (manual_auditor_col and not pd.isna(row[manual_auditor_col])) else "Manual Auditor"
        
        field_vals = {}
        for f in FIELDS:
            col_name = manual_mapping.get(f)
            val = row[col_name] if col_name else None
            field_vals[f] = None if (pd.isna(val) or val is None or str(val).strip().lower() in ["", "none", "null", "-", "nan"]) else str(val).strip()
            
        manual_calls[call_id] = {"auditor": auditor, "fields": field_vals}
        
    ai_calls = {}
    for _, row in df_ai.iterrows():
        call_id = str(row[ai_call_id_col]).split('.')[0].strip()
        field_vals = {}
        for f in FIELDS:
            col_name = ai_mapping.get(f)
            val = row[col_name] if col_name else None
            field_vals[f] = None if (pd.isna(val) or val is None or str(val).strip().lower() in ["", "none", "null", "-", "nan"]) else str(val).strip()
            
        ai_calls[call_id] = {"fields": field_vals}
        
    # Intersect call IDs
    common_cids = sorted(list(set(manual_calls.keys()) & set(ai_calls.keys())))
    logger.info(f"Overlap Call IDs compared: {len(common_cids)}")
    
    total_matched_fields = 0
    data_calls = []
    
    for call_id in common_cids:
        m_call = manual_calls[call_id]
        a_call = ai_calls[call_id]
        
        matches_by_field = {}
        manual_vals = {}
        ai_vals = {}
        
        for field in FIELDS:
            m_val_raw = m_call["fields"].get(field)
            a_val_raw = a_call["fields"].get(field)
            
            is_blank = (m_val_raw is None)
            if is_blank:
                norm_a_val = normalize_value(a_val_raw)
                matches_by_field[field] = 1
                manual_vals[field] = norm_a_val
                ai_vals[field] = norm_a_val
            else:
                norm_m_val = normalize_value(m_val_raw)
                norm_a_val = normalize_value(a_val_raw)
                matches_by_field[field] = get_match_status(norm_m_val, norm_a_val)
                manual_vals[field] = norm_m_val
                ai_vals[field] = norm_a_val
                
        matches_count = sum(matches_by_field.values())
        total_matched_fields += matches_count
        accuracy = (matches_count / len(FIELDS)) * 100
        
        data_calls.append({
            "call_id": call_id,
            "manual_auditor": m_call["auditor"],
            "matches_by_field": matches_by_field,
            "matches": matches_count,
            "total_fields": len(FIELDS),
            "accuracy": accuracy,
            "manual_values": manual_vals,
            "ai_values": ai_vals
        })
        
    # Calculate statistics for Sheet 7 (comparative stats)
    field_accuracies = {}
    for field in FIELDS:
        matches = sum(item["matches_by_field"][field] for item in data_calls)
        field_accuracies[field] = (matches / len(common_cids)) * 100 if common_cids else 0.0
        
    overall_accuracy = (total_matched_fields / (len(common_cids) * len(FIELDS))) * 100 if common_cids else 0.0
    
    # Calculate Sheet 4: Overview Data
    yes_ai = sum(df_ai[verdict_ai_col].astype(str).str.strip().str.lower() == 'yes') if verdict_ai_col else 0
    no_ai = sum(df_ai[verdict_ai_col].astype(str).str.strip().str.lower() == 'no') if verdict_ai_col else 0
    yes_manual = sum(df_manual[verdict_man_col].astype(str).str.strip().str.lower() == 'yes') if verdict_man_col else 0
    no_manual = sum(df_manual[verdict_man_col].astype(str).str.strip().str.lower() == 'no') if verdict_man_col else 0
    
    overview_data = {
        "ai_summary": {
            "total": len(df_ai),
            "yes": yes_ai,
            "no": no_ai
        },
        "manual_summary": {
            "total": len(df_manual),
            "yes": yes_manual,
            "no": no_manual
        },
        "comparison_summary": {
            "matched_calls": len(common_cids),
            "total_fields": len(common_cids) * len(FIELDS),
            "matched_fields": total_matched_fields
        }
    }
    
    output_path = get_next_filename(args.output_dir, "Compact_Analysis")
    logger.info(f"Generating 7-sheet workbook at: {output_path}")
    
    create_excel_report(
        summary_ai, overall_ai,
        summary_manual, overall_manual,
        summary_ai_matched, overall_ai_matched,
        overview_data,
        data_calls, field_accuracies, overall_accuracy,
        output_path
    )
    
    logger.info("Excel-to-Excel Call Audit Analyser completed successfully!")

if __name__ == "__main__":
    main()
