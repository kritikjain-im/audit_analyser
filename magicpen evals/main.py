import os
import sys
import time
import argparse
import logging
import pandas as pd
import config
import utils
import prompt_loader
import benchmark
import excel_writer

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

def load_csv_robust(file_path: str) -> pd.DataFrame:
    """
    Attempts to load a CSV file using multiple encodings to handle special characters safely.
    """
    for encoding in ["utf-8", "latin-1", "cp1252", "utf-8-sig"]:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            logger.info(f"Successfully loaded CSV with encoding: {encoding}")
            return df
        except Exception:
            continue
    # Fallback to standard read_csv
    return pd.read_csv(file_path)

def main() -> None:
    parser = argparse.ArgumentParser(description="MagicPen LLM Benchmarking Utility")
    parser.add_argument(
        "--input-csv", 
        default=config.INPUT_CSV_PATH,
        help="Path to the input CSV file containing magicpen data"
    )
    parser.add_argument(
        "--workers", 
        type=int,
        default=config.MAX_WORKERS,
        help="Maximum concurrent worker threads"
    )
    parser.add_argument(
        "--output-excel", 
        default=config.OUTPUT_EXCEL_PATH,
        help="Path to save the generated Excel workbook"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=15,
        help="Limit the number of rows to evaluate (set to 0 for unlimited)"
    )
    
    args = parser.parse_args()
    
    # Override configuration with CLI arguments
    config.INPUT_CSV_PATH = args.input_csv
    config.MAX_WORKERS = args.workers
    config.OUTPUT_EXCEL_PATH = args.output_excel
    
    # Start tracking total runtime
    script_start_time = time.perf_counter()
    
    # 1. Verify files exist
    if not os.path.exists(config.INPUT_CSV_PATH):
        logger.error(f"Input CSV not found at: {config.INPUT_CSV_PATH}")
        sys.exit(1)
    if not os.path.exists(config.PROMPT_TEMPLATE_PATH):
        logger.error(f"Prompt template file not found at: {config.PROMPT_TEMPLATE_PATH}")
        sys.exit(1)
        
    # 2. Load CSV rows
    logger.info("Loading input CSV...")
    df = load_csv_robust(config.INPUT_CSV_PATH)
    if args.limit > 0:
        logger.info(f"Limiting evaluation to the first {args.limit} rows.")
        df = df.head(args.limit)
    rows = df.to_dict(orient="records")
    total_rows = len(rows)
    logger.info(f"Loaded {total_rows} rows from input sheet.")
    
    # 3. Load template
    logger.info("Loading prompt template...")
    system_template, user_template = prompt_loader.load_template(config.PROMPT_TEMPLATE_PATH)
    
    # 4. Run benchmark orchestrator
    logger.info("Starting benchmarking runs...")
    results = benchmark.run_benchmark(rows, system_template, user_template)
    
    # 5. Generate Excel spreadsheet
    logger.info(f"Writing Excel comparison workbook to: {config.OUTPUT_EXCEL_PATH}...")
    try:
        excel_writer.write_magicpen_report(rows, results, config.OUTPUT_EXCEL_PATH)
        logger.info("Excel report written successfully.")
    except PermissionError:
        fallback_path = config.OUTPUT_EXCEL_PATH.replace(".xlsx", f"_{int(time.time())}.xlsx")
        logger.warning(f"Permission denied on output Excel file. It might be open in Excel. Saving to fallback path: {fallback_path}")
        try:
            excel_writer.write_magicpen_report(rows, results, fallback_path)
            logger.info(f"Excel report written successfully to: {fallback_path}")
            config.OUTPUT_EXCEL_PATH = fallback_path
        except Exception as e:
            logger.error(f"Failed to write Excel report to fallback path: {str(e)}")
    except Exception as e:
        logger.error(f"Failed to write Excel report: {str(e)}")
        
    # 6. Calculate statistics and summaries
    total_runtime = time.perf_counter() - script_start_time
    
    success_requests = 0
    failed_requests = 0
    
    model_stats = {}
    for model, res_list in results.items():
        latencies = []
        prompt_tokens_list = []
        completion_tokens_list = []
        total_tokens_list = []
        costs = []
        
        for r in res_list:
            if r["status"] == "Success":
                success_requests += 1
                latencies.append(r["latency"])
                if isinstance(r["prompt_tokens"], int):
                    prompt_tokens_list.append(r["prompt_tokens"])
                if isinstance(r["completion_tokens"], int):
                    completion_tokens_list.append(r["completion_tokens"])
                if isinstance(r["total_tokens"], int):
                    total_tokens_list.append(r["total_tokens"])
                if isinstance(r["cost"], (int, float)):
                    costs.append(r["cost"])
            else:
                failed_requests += 1
                latencies.append(r["latency"]) # latency of failed runs is still recorded
                
        model_stats[model] = {
            "avg_latency": utils.calculate_mean(latencies),
            "p95_latency": utils.calculate_percentile(latencies, 0.95),
            "avg_prompt_tokens": utils.calculate_mean(prompt_tokens_list),
            "avg_completion_tokens": utils.calculate_mean(completion_tokens_list),
            "avg_total_tokens": utils.calculate_mean(total_tokens_list),
            "avg_cost": utils.calculate_mean(costs)
        }
        
    # Print console summary dashboard
    print("\n" + "=" * 40)
    print("        MAGICPEN BENCHMARK SUMMARY")
    print("=" * 40)
    print(f"Rows Processed      : {total_rows}")
    print(f"Successful Requests : {success_requests}")
    print(f"Failed Requests     : {failed_requests}")
    print(f"Total Runtime       : {total_runtime:.2f} seconds")
    print("-" * 40)
    
    for model, stats in model_stats.items():
        display_name = config.MODELS_CONFIG.get(model, {}).get("display_name", model)
        print(f"\nModel: {display_name} ({model})")
        print(f"  Average Latency          : {stats['avg_latency']:.3f} sec")
        print(f"  P95 Latency              : {stats['p95_latency']:.3f} sec")
        print(f"  Average Prompt Tokens    : {stats['avg_prompt_tokens']:.1f}")
        print(f"  Average Completion Tokens: {stats['avg_completion_tokens']:.1f}")
        print(f"  Average Total Tokens     : {stats['avg_total_tokens']:.1f}")
        print(f"  Average Estimated Cost   : ${stats['avg_cost']:.8f}")
        
    print("=" * 40 + "\n")
    logger.info("MagicPen Benchmarker finished successfully!")

if __name__ == "__main__":
    main()
