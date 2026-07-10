import os
import sys
import pandas as pd
import json
import logging
import ast

# Ensure parent directory is in sys.path to allow imports if needed
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import config
import utils

# Set up logging to stdout
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def extract_inferred_details_from_call_json(call_json_str):
    """
    Dynamically extracts buyer and seller location/name details from the call_json structure
    and compiles them into a standardized inferred_details JSON string.
    """
    if pd.isna(call_json_str) or not isinstance(call_json_str, str):
        return None
    try:
        # Safely parse JSON or Python dict representation
        try:
            data = json.loads(call_json_str)
        except json.JSONDecodeError:
            data = ast.literal_eval(call_json_str.strip())
            
        if not isinstance(data, dict):
            return None
            
        buyer_details = data.get("buyer_details", {})
        seller_details = data.get("seller_details", {})
        
        buyer_name = buyer_details.get("buyer_name") or buyer_details.get("name") or ""
        buyer_location = buyer_details.get("buyer_location", {})
        buyer_city = ""
        if isinstance(buyer_location, dict):
            buyer_city = buyer_location.get("buyer_city") or buyer_location.get("city") or ""
        elif isinstance(buyer_location, str):
            buyer_city = buyer_location
            
        buyer_number = buyer_details.get("buyer_number") or buyer_details.get("buyer_phone") or data.get("buyer_Number") or ""
        
        seller_name = seller_details.get("seller_name") or seller_details.get("name") or ""
        seller_location = seller_details.get("seller_location", {})
        seller_city = ""
        if isinstance(seller_location, dict):
            seller_city = seller_location.get("seller_city") or seller_location.get("city") or ""
        elif isinstance(seller_location, str):
            seller_city = seller_location
            
        inferred = {
            "buyer_city": buyer_city,
            "buyer_name": buyer_name,
            "buyer_number": buyer_number,
            "seller_city": seller_city,
            "seller_name": seller_name
        }
        if any(inferred.values()):
            return json.dumps(inferred, ensure_ascii=False)
    except Exception as e:
        logger.debug(f"Failed to extract inferred_details from call_json: {e}")
        
    return None

def load_and_preprocess_csv(file_path, description):
    """
    Loads CSV, normalizes empty values, and handles encoding robustly.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found at: {file_path}")
        
    logger.info(f"Loading {description} from {file_path}...")
    try:
        df = utils.load_csv_robust(file_path)
    except Exception as e:
        logger.error(f"Error reading CSV file {file_path}: {e}")
        raise
        
    logger.info(f"Loaded {len(df)} rows from {description}.")
    return df

def main():
    # Make sure output directory exists
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    
    # Verify input files exist
    if (not os.path.exists(config.CALLS_CSV_PATH) or 
        not os.path.exists(config.MANUAL_AUDIT_CSV_PATH) or
        not os.path.exists(config.CALL_JSON_CSV_PATH)):
        logger.error("Required input CSV files are missing. Please place calls.csv, manual_audit.csv, and call_jsons.csv in input/ directory.")
        sys.exit(1)
        
    # Step 1: Load three CSVs
    try:
        calls_df = load_and_preprocess_csv(config.CALLS_CSV_PATH, "calls.csv")
        manual_df = load_and_preprocess_csv(config.MANUAL_AUDIT_CSV_PATH, "manual_audit.csv")
        call_jsons_df = load_and_preprocess_csv(config.CALL_JSON_CSV_PATH, "call_jsons.csv")
    except Exception as e:
        logger.error(f"Error during loading and preprocessing: {e}")
        sys.exit(1)
        
    # Step 2: Resolve column names dynamically
    # Resolve Call ID column in all files
    calls_id_col = utils.resolve_column(calls_df, ["Call ID", "call_id", "CallID"])
    manual_id_col = utils.resolve_column(manual_df, ["Call ID", "call_id", "CallID"])
    call_jsons_id_col = utils.resolve_column(call_jsons_df, ["Call ID", "call_id", "CallID", "PNS_Call_ID"])
    
    if not calls_id_col:
        raise KeyError("Could not find Call ID column in calls.csv")
    if not manual_id_col:
        raise KeyError("Could not find Call ID column in manual_audit.csv")
    if not call_jsons_id_col:
        raise KeyError("Could not find Call ID column in call_jsons.csv")
        
    # Resolve key columns in manual_audit.csv
    transcript_col = utils.resolve_column(manual_df, ["Transcript", "transcript"])
    type_col = utils.resolve_column(manual_df, ["Type", "type"])
    
    if not transcript_col:
        raise KeyError("Could not find Transcript column in manual_audit.csv")
    if not type_col:
        raise KeyError("Could not find Type column in manual_audit.csv")
        
    # Resolve optional call_json and inferred_details column names
    calls_call_json_col = utils.resolve_column(calls_df, ["call_json", "Call JSON", "CallJSON"])
    calls_inferred_col = utils.resolve_column(calls_df, ["inferred_details", "Inferred Details", "InferredDetails"])
    
    call_jsons_call_json_col = utils.resolve_column(call_jsons_df, ["call_json", "Call JSON", "CallJSON"])
    call_jsons_inferred_col = utils.resolve_column(call_jsons_df, ["inferred_details", "Inferred Details", "InferredDetails"])
    
    # Step 3: Handle duplicates and standard indexing
    # Normalize ID series to strings, stripping whitespace
    calls_df[calls_id_col] = calls_df[calls_id_col].fillna("").astype(str).str.strip()
    calls_df = calls_df[calls_df[calls_id_col] != ""]
    calls_df = calls_df[calls_df[calls_id_col].str.lower() != "nan"]
    calls_dups = calls_df.duplicated(subset=[calls_id_col]).sum()
    calls_clean = calls_df.drop_duplicates(subset=[calls_id_col], keep="first").set_index(calls_id_col)
    
    call_jsons_df[call_jsons_id_col] = call_jsons_df[call_jsons_id_col].fillna("").astype(str).str.strip()
    call_jsons_df = call_jsons_df[call_jsons_df[call_jsons_id_col] != ""]
    call_jsons_df = call_jsons_df[call_jsons_df[call_jsons_id_col].str.lower() != "nan"]
    call_jsons_dups = call_jsons_df.duplicated(subset=[call_jsons_id_col]).sum()
    call_jsons_clean = call_jsons_df.drop_duplicates(subset=[call_jsons_id_col], keep="first").set_index(call_jsons_id_col)

    # Normalize manual_audit.csv Call ID and Type columns
    manual_df[manual_id_col] = manual_df[manual_id_col].fillna("").astype(str).str.strip()
    manual_df = manual_df[manual_df[manual_id_col] != ""]
    manual_df = manual_df[manual_df[manual_id_col].str.lower() != "nan"]
    
    # 3a. Build Transcript Lookup map from ALL rows in manual_audit.csv (so we can retrieve transcripts from 'ai' rows if 'manual' rows are missing them)
    valid_transcripts = manual_df[manual_df[transcript_col].notna()]
    valid_transcripts = valid_transcripts[valid_transcripts[transcript_col].astype(str).str.strip() != ""]
    transcript_lookup = dict(zip(valid_transcripts[manual_id_col], valid_transcripts[transcript_col]))
    logger.info(f"Built transcript lookup table with {len(transcript_lookup)} transcripts.")

    # 3b. Filter manual_audit.csv to include only rows of Type == 'manual'
    manual_df['Type_clean'] = manual_df[type_col].fillna("").astype(str).str.strip().str.lower()
    manual_only_df = manual_df[manual_df['Type_clean'] == 'manual']
    manual_dups = manual_only_df.duplicated(subset=[manual_id_col]).sum()
    manual_only_clean = manual_only_df.drop_duplicates(subset=[manual_id_col], keep="first").set_index(manual_id_col)
    logger.info(f"Filtered manual_audit.csv to {len(manual_only_clean)} manual audit entries.")

    # Step 4: Find the eligible Call IDs
    # Since call_jsons.csv contains the PNS call records (most call JSONs), and manual audits represent our ground truth,
    # the eligible Call IDs are the manual audit entries that exist in call_jsons.csv OR calls.csv.
    manual_only_ids = set(manual_only_clean.index)
    calls_ids = set(calls_clean.index)
    pns_ids = set(call_jsons_clean.index)
    
    eligible_ids = sorted(list(manual_only_ids.intersection(pns_ids.union(calls_ids))))
    logger.info(f"Identified {len(eligible_ids)} manual audits with matching Call JSON sources (intersection of manual and PNS/calls).")

    # Step 5: Validate and generate rows
    final_rows = []
    
    # Counters for missing data log
    missing_transcript_count = 0
    missing_call_json_count = 0
    missing_inferred_details_count = 0
    missing_manual_audit_count = 0
    
    for call_id in eligible_ids:
        manual_row = manual_only_clean.loc[call_id]
        
        # 1. Retrieve Transcript (using cross-type lookup)
        transcript = utils.normalize_value(transcript_lookup.get(call_id))
        
        # 2. Retrieve Call JSON (checks call_jsons.csv first, calls.csv second)
        call_json = None
        if call_id in call_jsons_clean.index and call_jsons_call_json_col:
            call_json = utils.normalize_value(call_jsons_clean.loc[call_id][call_jsons_call_json_col])
            
        if (call_json is None or call_json == "") and call_id in calls_clean.index and calls_call_json_col:
            call_json = utils.normalize_value(calls_clean.loc[call_id][calls_call_json_col])
            
        # 3. Retrieve Inferred Details (checks calls.csv first, call_jsons.csv second, manual_audit.csv third)
        inferred_details = None
        if call_id in calls_clean.index and calls_inferred_col:
            inferred_details = utils.normalize_value(calls_clean.loc[call_id][calls_inferred_col])
            
        if (inferred_details is None or inferred_details == "") and call_id in call_jsons_clean.index and call_jsons_inferred_col:
            inferred_details = utils.normalize_value(call_jsons_clean.loc[call_id][call_jsons_inferred_col])
            
        if (inferred_details is None or inferred_details == ""):
            inferred_details = utils.normalize_value(manual_row.get('Inferred Details') or manual_row.get('inferred_details'))
            
        # 3a. Reconstruct inferred_details from call_json as a robust fallback
        if (inferred_details is None or inferred_details == "") and call_json:
            inferred_details = extract_inferred_details_from_call_json(call_json)
            
        # 4. Construct Expected Output from manual row
        expected_output_json = utils.construct_expected_output(manual_row, config.EXPECTED_OUTPUT_FIELDS_MAPPING)
        expected_output_dict = json.loads(expected_output_json)
        
        # Validation checks (mutually exclusive counting)
        if transcript is None or transcript == "":
            missing_transcript_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing Transcript.")
            continue
            
        if call_json is None or call_json == "":
            missing_call_json_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing call_json.")
            continue
            
        if inferred_details is None or inferred_details == "":
            missing_inferred_details_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing inferred_details.")
            continue
            
        if not expected_output_dict:
            missing_manual_audit_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing manual audit fields.")
            continue
            
        # Format JSON columns for Langfuse output
        formatted_call_json = utils.parse_and_format_json(call_json)
        formatted_inferred = utils.parse_and_format_json(inferred_details)
        metadata_text = utils.construct_metadata(call_id)
        
        final_rows.append({
            "Call JSON": formatted_call_json,
            "Transcript": transcript,
            "Inferred details": formatted_inferred,
            "expected_output": expected_output_json,
            "metadata": metadata_text
        })
        
    # Step 6: Create final DataFrame
    output_df = pd.DataFrame(final_rows)
    
    # Step 7: Export langfuse_dataset.csv
    try:
        output_df.to_csv(config.LANGFUSE_DATASET_CSV_PATH, index=False)
        logger.info(f"Dataset generated successfully at {config.LANGFUSE_DATASET_CSV_PATH}.")
    except Exception as e:
        logger.error(f"Error saving output dataset: {e}")
        sys.exit(1)
        
    # Statistics Logging
    report_lines = [
        f"Rows in calls.csv : {len(calls_df)}",
        f"Rows in manual audit : {len(manual_df)}",
        f"Manual entries (Type == 'manual') : {len(manual_only_df)}",
        f"Rows in call_jsons.csv : {len(call_jsons_df)}",
        f"Eligible Call IDs for matching : {len(eligible_ids)}",
        f"Missing transcript : {missing_transcript_count}",
        f"Missing call_json : {missing_call_json_count}",
        f"Missing inferred_details : {missing_inferred_details_count}",
        f"Missing manual audit : {missing_manual_audit_count}",
        f"Final dataset size : {len(output_df)}"
    ]
    
    # Log duplicates information if any
    duplicates_lines = []
    if calls_dups > 0:
        duplicates_lines.append(f"Duplicate Call IDs removed from calls.csv: {calls_dups}")
    if manual_dups > 0:
        duplicates_lines.append(f"Duplicate Call IDs removed from manual_audit.csv (manual type): {manual_dups}")
    if call_jsons_dups > 0:
        duplicates_lines.append(f"Duplicate Call IDs removed from call_jsons.csv: {call_jsons_dups}")
        
    full_report = "\n".join(report_lines)
    if duplicates_lines:
        full_report += "\n\nDuplicate Records:\n" + "\n".join(duplicates_lines)
        
    # Print statistics
    print("\n=== Generation Summary ===")
    print(full_report)
    print("==========================\n")
    
    # Save generation_report.txt
    try:
        with open(config.GENERATION_REPORT_TXT_PATH, "w", encoding="utf-8") as f:
            f.write(full_report)
        logger.info(f"Report saved successfully at {config.GENERATION_REPORT_TXT_PATH}.")
    except Exception as e:
        logger.error(f"Error saving output report: {e}")

if __name__ == "__main__":
    main()
