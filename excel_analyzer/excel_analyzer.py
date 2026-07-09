import os
import sys
import argparse
import glob
import logging
import re
import pandas as pd
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
    """Clean and return a normalized string for openpyxl reporting."""
    if pd.isna(val) or val is None:
        return "Correct (Not Discussed)"
    
    val_str = str(val).strip()
    if val_str == "" or val_str.lower() in ["none", "null", "-", "nan"]:
        return "Correct (Not Discussed)"
        
    norm = normalize_badge(val_str)
    if norm:
        return norm
        
    # String fallbacks
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
        # Substring fallback for Call ID
        for clean_h, orig_h in normalized_headers.items():
            if "callid" in clean_h or "callid" in clean_h:
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
        # Substring fallback for Auditor Name
        for clean_h, orig_h in normalized_headers.items():
            if "auditor" in clean_h or "audited" in clean_h:
                auditor_col = orig_h
                break
    mapping["auditor"] = auditor_col

    # Map the 12 Fields
    for field, targets in fields_targets.items():
        matched_header = None
        # Exact match
        for t in targets:
            clean_t = re.sub(r'[^a-zA-Z0-9]', '', t).lower()
            if clean_t in normalized_headers:
                matched_header = normalized_headers[clean_t]
                break
        
        if matched_header:
            mapping[field] = matched_header
        else:
            # Substring match fallback
            for clean_h, orig_h in normalized_headers.items():
                for t in targets:
                    clean_t = re.sub(r'[^a-zA-Z0-9]', '', t).lower()
                    # Filter out subfields like 'reason' or 'sublabel' from matching main field
                    if clean_t in clean_h and not any(sub in clean_h for sub in ["reason", "sublabel", "subfield"]):
                        matched_header = orig_h
                        break
                if matched_header:
                    break
            if matched_header:
                logger.info(f"[{sheet_name}] Substring match: mapped field '{field}' to column '{matched_header}'")
                mapping[field] = matched_header
            else:
                mapping[field] = None
                
    return mapping

def get_next_filename(base_dir, base_name="AI_Auditor_Comparison_Sheet"):
    version = 1
    while True:
        candidate = os.path.join(base_dir, f"{base_name}_v{version}.xlsx")
        if not os.path.exists(candidate):
            return candidate
        version += 1

def main():
    parser = argparse.ArgumentParser(description="Direct Excel-to-Excel Call Audit Comparison Analyser")
    parser.add_argument("--input-file", "-i", help="Path to input Excel workbook")
    parser.add_argument("--output-dir", "-o", default="output", help="Directory to save output reports")
    parser.add_argument("--manual-sheet", default=None, help="Name of manual audits sheet")
    parser.add_argument("--ai-sheet", default=None, help="Name of AI audits sheet")
    
    args = parser.parse_args()
    
    # 1. Locate Input File
    input_path = args.input_file
    if not input_path:
        # Search output or input directories for xlsx files
        xlsx_candidates = []
        scan_dirs = ["input", "input2", "."]
        for s_dir in scan_dirs:
            if os.path.exists(s_dir):
                xlsx_candidates.extend(glob.glob(os.path.join(s_dir, "*.xlsx")))
        
        # Exclude generated output files if they are in the scan path
        xlsx_candidates = [f for f in xlsx_candidates if "AI_Audit_Analysis" not in f and "AI_Auditor_Comparison" not in f]
        
        if xlsx_candidates:
            input_path = sorted(xlsx_candidates)[0]
            logger.info(f"Auto-detected input Excel file: {input_path}")
        else:
            logger.error("No input Excel workbook specified or auto-detected.")
            sys.exit(1)
            
    if not os.path.exists(input_path):
        logger.error(f"Input file '{input_path}' does not exist.")
        sys.exit(1)
        
    os.makedirs(args.output_dir, exist_ok=True)
    
    # 2. Load DataFrames (Support CSV and Excel)
    df_manual = None
    df_ai = None
    manual_sheet = args.manual_sheet or "Manual Audits"
    ai_sheet = args.ai_sheet or "AI Audits"
    
    is_csv = input_path.lower().endswith('.csv')
    
    if is_csv:
        logger.info(f"Processing input path as a CSV file: {input_path}")
        try:
            df = pd.read_csv(input_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(input_path, encoding='latin1')
        except Exception as e:
            logger.error(f"Failed to read CSV file '{input_path}': {str(e)}")
            sys.exit(1)
            
        # Check if user specified a separate AI CSV via --ai-sheet
        if args.ai_sheet and args.ai_sheet.lower().endswith('.csv'):
            ai_csv_path = args.ai_sheet
            if not os.path.exists(ai_csv_path):
                logger.error(f"AI CSV file '{ai_csv_path}' specified via --ai-sheet does not exist.")
                sys.exit(1)
            logger.info(f"Using separate AI CSV file: {ai_csv_path}")
            df_manual = df
            try:
                df_ai = pd.read_csv(ai_csv_path, encoding='utf-8')
            except UnicodeDecodeError:
                df_ai = pd.read_csv(ai_csv_path, encoding='latin1')
            except Exception as e:
                logger.error(f"Failed to read AI CSV file '{ai_csv_path}': {str(e)}")
                sys.exit(1)
        else:
            # Split single CSV file containing both AI and Manual audits
            headers = list(df.columns)
            mapping_temp = map_columns(headers, FIELD_TARGETS, "CSV")
            call_id_col = mapping_temp.get("call_id")
            if not call_id_col:
                logger.error(f"Could not locate 'Call ID' column in CSV. Headers: {headers}")
                sys.exit(1)
                
            # Drop empty Call ID rows
            df = df.dropna(subset=[call_id_col])
            
            auditor_col = mapping_temp.get("auditor")
            type_col = None
            for col in df.columns:
                if str(col).lower().strip() == 'type':
                    type_col = col
                    break
                    
            ai_mask = pd.Series(False, index=df.index)
            if type_col is not None:
                ai_mask = ai_mask | (df[type_col].astype(str).str.lower().str.strip() == 'ai')
            if auditor_col is not None:
                ai_mask = ai_mask | (df[auditor_col].astype(str).str.contains('AI', case=False, na=False))
                
            df_ai = df[ai_mask]
            df_manual = df[~ai_mask]
            
            logger.info(f"Split single CSV: {len(df_manual)} Manual rows, {len(df_ai)} AI rows.")
    else:
        # Load Workbook sheets (Excel)
        try:
            xls = pd.ExcelFile(input_path)
        except Exception as e:
            logger.error(f"Failed to read Excel file '{input_path}': {str(e)}")
            sys.exit(1)
            
        sheet_names = xls.sheet_names
        logger.info(f"Sheets found in workbook: {sheet_names}")
        
        # Determine Manual sheet name
        manual_sheet = args.manual_sheet
        if not manual_sheet:
            for name in sheet_names:
                if "manual" in name.lower():
                    manual_sheet = name
                    break
            if not manual_sheet and len(sheet_names) > 0:
                manual_sheet = sheet_names[0]
                
        # Determine AI sheet name
        ai_sheet = args.ai_sheet
        if not ai_sheet:
            for name in sheet_names:
                if "ai" in name.lower() and name != manual_sheet:
                    ai_sheet = name
                    break
            if not ai_sheet and len(sheet_names) > 1:
                ai_sheet = sheet_names[1]
                
        if not manual_sheet or not ai_sheet:
            logger.error("Could not determine Manual and AI sheets in the workbook.")
            sys.exit(1)
            
        logger.info(f"Using Manual Sheet: '{manual_sheet}'")
        logger.info(f"Using AI Sheet: '{ai_sheet}'")
        
        try:
            df_manual = pd.read_excel(input_path, sheet_name=manual_sheet)
            df_ai = pd.read_excel(input_path, sheet_name=ai_sheet)
        except Exception as e:
            logger.error(f"Failed to load sheets: {str(e)}")
            sys.exit(1)
        
    # 3. Map Columns
    manual_mapping = map_columns(df_manual.columns, FIELD_TARGETS, manual_sheet)
    ai_mapping = map_columns(df_ai.columns, FIELD_TARGETS, ai_sheet)
    
    if not manual_mapping.get("call_id"):
        logger.error(f"Could not locate 'Call ID' column in Manual sheet. Headers: {list(df_manual.columns)}")
        sys.exit(1)
    if not ai_mapping.get("call_id"):
        logger.error(f"Could not locate 'Call ID' column in AI sheet. Headers: {list(df_ai.columns)}")
        sys.exit(1)
        
    # Check mapped fields
    missing_manual_fields = [f for f in FIELDS if not manual_mapping.get(f)]
    missing_ai_fields = [f for f in FIELDS if not ai_mapping.get(f)]
    
    if missing_manual_fields:
        logger.warning(f"Fields missing from Manual sheet mapping: {missing_manual_fields}. These fields will default to empty.")
    if missing_ai_fields:
        logger.warning(f"Fields missing from AI sheet mapping: {missing_ai_fields}. These fields will default to 'Correct'.")
        
    # 4. Extract Manual Data (De-duplicating Call IDs)
    manual_calls = {}
    manual_call_id_col = manual_mapping["call_id"]
    manual_auditor_col = manual_mapping.get("auditor")
    
    for idx, row in df_manual.iterrows():
        raw_cid = row[manual_call_id_col]
        if pd.isna(raw_cid):
            continue
        # Clean Call ID
        call_id = str(raw_cid).split('.')[0].strip()
        if not call_id:
            continue
            
        if call_id in manual_calls:
            logger.warning(f"Duplicate Call ID '{call_id}' in Manual sheet. Discarding duplicate row.")
            continue
            
        # Get Auditor Name
        auditor = "Manual Auditor"
        if manual_auditor_col and not pd.isna(row[manual_auditor_col]):
            auditor = str(row[manual_auditor_col]).strip()
            
        # Build field map (keeping raw values for blank check)
        field_vals = {}
        for f in FIELDS:
            col_name = manual_mapping.get(f)
            val = row[col_name] if col_name else None
            # Keep raw/blank check representation
            if pd.isna(val) or val is None:
                val = None
            else:
                val_str = str(val).strip()
                if val_str == "" or val_str.lower() in ["none", "null", "-", "nan"]:
                    val = None
                else:
                    val = val_str
            field_vals[f] = val
            
        manual_calls[call_id] = {
            "auditor": auditor,
            "fields": field_vals
        }
        
    # 5. Extract AI Data (De-duplicating Call IDs)
    ai_calls = {}
    ai_call_id_col = ai_mapping["call_id"]
    
    for idx, row in df_ai.iterrows():
        raw_cid = row[ai_call_id_col]
        if pd.isna(raw_cid):
            continue
        call_id = str(raw_cid).split('.')[0].strip()
        if not call_id:
            continue
            
        if call_id in ai_calls:
            logger.warning(f"Duplicate Call ID '{call_id}' in AI sheet. Discarding duplicate row.")
            continue
            
        # Build field map (keeping raw values for comparison)
        field_vals = {}
        for f in FIELDS:
            col_name = ai_mapping.get(f)
            val = row[col_name] if col_name else None
            if pd.isna(val) or val is None:
                val = None
            else:
                val_str = str(val).strip()
                if val_str == "" or val_str.lower() in ["none", "null", "-", "nan"]:
                    val = None
                else:
                    val = val_str
            field_vals[f] = val
            
        ai_calls[call_id] = {
            "fields": field_vals
        }
        
    # 6. Intersect to common Call IDs
    manual_cids = set(manual_calls.keys())
    ai_cids = set(ai_calls.keys())
    common_cids = sorted(list(manual_cids & ai_cids))
    
    excluded_manual = manual_cids - ai_cids
    excluded_ai = ai_cids - manual_cids
    
    logger.info(f"Manual sheet total calls: {len(manual_cids)}")
    logger.info(f"AI sheet total calls: {len(ai_cids)}")
    logger.info(f"Common Call IDs to process: {len(common_cids)}")
    
    if excluded_manual:
        logger.info(f"Manual calls missing in AI sheet (excluded): {sorted(list(excluded_manual))}")
    if excluded_ai:
        logger.info(f"AI calls missing in Manual sheet (excluded): {sorted(list(excluded_ai))}")
        
    if not common_cids:
        logger.error("No common Call IDs found between Manual and AI sheets.")
        sys.exit(1)
        
    # 7. Perform Comparisons
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
            
            # Check for blank values
            is_blank = (m_val_raw is None)
                    
            if is_blank:
                # Blank matches AI value automatically
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
        
    # 8. Calculate Statistics
    total_calls = len(data_calls)
    field_accuracies = {}
    for field in FIELDS:
        matches = sum(item["matches_by_field"][field] for item in data_calls)
        field_accuracies[field] = (matches / total_calls) * 100
        
    overall_accuracy = sum(item["matches"] for item in data_calls) / (total_calls * len(FIELDS)) * 100
    
    # 9. Output report
    output_path = get_next_filename(args.output_dir)
    logger.info(f"Generating direct sheet analysis report at: {output_path}")
    
    create_excel_report(data_calls, field_accuracies, overall_accuracy, output_path)
    logger.info("Excel-to-Excel Call Audit Analyser completed successfully!")

if __name__ == "__main__":
    main()
