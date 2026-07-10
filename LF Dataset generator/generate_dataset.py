import os
import sys
import pandas as pd
import json
import logging

# Ensure parent directory is in sys.path to allow imports if needed
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import config
import utils

# Set up logging to stdout
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def load_and_preprocess_csv(file_path, required_cols_mapping, description):
    """
    Loads CSV, validates that necessary columns (or their aliases) exist,
    normalizes Call ID, handles duplicates, and returns cleaned DataFrame and resolved column names.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found at: {file_path}")
        
    logger.info(f"Loading {description} from {file_path}...")
    # Load the CSV
    try:
        df = pd.read_csv(file_path, dtype=str)
    except Exception as e:
        logger.error(f"Error reading CSV file {file_path}: {e}")
        raise
        
    logger.info(f"Loaded {len(df)} rows from {description}.")
    
    # Resolve required columns
    resolved_cols = {}
    for key, aliases in required_cols_mapping.items():
        col_name = utils.resolve_column(df, aliases)
        if not col_name:
            raise KeyError(f"Required column for '{key}' (tried aliases: {aliases}) not found in {description} columns: {list(df.columns)}")
        resolved_cols[key] = col_name
        logger.debug(f"Resolved key '{key}' to column '{col_name}' in {description}.")
        
    # Clean and normalize Call ID
    call_id_col = resolved_cols["call_id"]
    
    # Normalize empty values in Call ID
    df = df[df[call_id_col].notna()]
    df[call_id_col] = df[call_id_col].astype(str).str.strip()
    
    # Drop rows where Call ID is empty or 'nan'
    df = df[df[call_id_col] != ""]
    df = df[df[call_id_col].str.lower() != "nan"]
    
    # Handle duplicates
    initial_len = len(df)
    duplicates_count = df.duplicated(subset=[call_id_col]).sum()
    if duplicates_count > 0:
        logger.warning(f"Found {duplicates_count} duplicate Call IDs in {description}. Keeping first occurrence.")
        df = df.drop_duplicates(subset=[call_id_col], keep="first")
        
    return df, resolved_cols, duplicates_count

def main():
    # Make sure output directory exists
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    
    # Check if input files exist
    if not os.path.exists(config.CALLS_CSV_PATH) or not os.path.exists(config.MANUAL_AUDIT_CSV_PATH):
        logger.error("Input CSV files are missing. Please place them in input/ directory.")
        print(f"Error: Make sure input files exist at:\n- {config.CALLS_CSV_PATH}\n- {config.MANUAL_AUDIT_CSV_PATH}")
        sys.exit(1)
        
    # Step 1: Load both CSVs
    try:
        calls_df, calls_cols, calls_dups = load_and_preprocess_csv(
            config.CALLS_CSV_PATH, config.CALLS_REQUIRED_COLS, "calls.csv"
        )
        manual_df, manual_cols, manual_dups = load_and_preprocess_csv(
            config.MANUAL_AUDIT_CSV_PATH, config.MANUAL_AUDIT_REQUIRED_COLS, "manual_audit.csv"
        )
    except Exception as e:
        logger.error(f"Error during loading and preprocessing: {e}")
        sys.exit(1)
        
    # Step 2: Find intersection of Call IDs present in BOTH files
    calls_id_col = calls_cols["call_id"]
    manual_id_col = manual_cols["call_id"]
    
    calls_ids = set(calls_df[calls_id_col])
    manual_ids = set(manual_df[manual_id_col])
    common_ids = sorted(list(calls_ids.intersection(manual_ids)))
    
    logger.info(f"Intersection of Call IDs: {len(common_ids)} common IDs found.")
    
    # Set indices to Call ID for easy lookup
    calls_indexed = calls_df.set_index(calls_id_col)
    manual_indexed = manual_df.set_index(manual_id_col)
    
    # Step 3, 4 & 5: Validate and generate rows
    final_rows = []
    
    # Counters for missing data log
    missing_transcript_count = 0
    missing_call_json_count = 0
    missing_inferred_details_count = 0
    missing_summary_count = 0
    missing_manual_audit_count = 0
    
    for call_id in common_ids:
        calls_row = calls_indexed.loc[call_id]
        manual_row = manual_indexed.loc[call_id]
        
        # Extract fields from calls.csv
        call_json = utils.normalize_value(calls_row[calls_cols["call_json"]])
        inferred_details = utils.normalize_value(calls_row[calls_cols["inferred_details"]])
        call_summary = utils.normalize_value(calls_row[calls_cols["call_summary"]])
        
        # Extract fields from manual_audit.csv
        transcript = utils.normalize_value(manual_row[manual_cols["transcript"]])
        
        # Validate manual audit expected output fields
        expected_output_json = utils.construct_expected_output(manual_row, config.EXPECTED_OUTPUT_FIELDS_MAPPING)
        expected_output_dict = json.loads(expected_output_json)
        
        # Step 4: Validate. Reject row if ANY required field is empty.
        # Track counts in order to make statistics mutually exclusive and easy to read.
        if transcript is None:
            missing_transcript_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing Transcript.")
            continue
            
        if call_json is None:
            missing_call_json_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing call_json.")
            continue
            
        if inferred_details is None:
            missing_inferred_details_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing inferred_details.")
            continue
            
        if call_summary is None:
            missing_summary_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing Call Summary.")
            continue
            
        if not expected_output_dict:
            missing_manual_audit_count += 1
            logger.debug(f"Call ID {call_id}: Discarded due to missing manual audit fields.")
            continue
            
        # Step 5: Create columns
        formatted_call_json = utils.parse_and_format_json(call_json)
        formatted_inferred = utils.parse_and_format_json(inferred_details)
        metadata_text = utils.construct_metadata(call_id)
        
        final_rows.append({
            "Call JSON": formatted_call_json,
            "Inferred details": formatted_inferred,
            "Call Summary": call_summary,
            "Transcript": transcript,
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
        f"Rows in calls.csv : {len(calls_df) + calls_dups}",
        f"Rows in manual audit : {len(manual_df) + manual_dups}",
        f"Common Call IDs : {len(common_ids)}",
        f"Missing transcript : {missing_transcript_count}",
        f"Missing call_json : {missing_call_json_count}",
        f"Missing inferred_details : {missing_inferred_details_count}",
        f"Missing summary : {missing_summary_count}",
        f"Missing manual audit : {missing_manual_audit_count}",
        f"Final dataset size : {len(output_df)}"
    ]
    
    # Log duplicates information if any
    duplicates_lines = []
    if calls_dups > 0:
        duplicates_lines.append(f"Duplicate Call IDs removed from calls.csv: {calls_dups}")
    if manual_dups > 0:
        duplicates_lines.append(f"Duplicate Call IDs removed from manual_audit.csv: {manual_dups}")
        
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
