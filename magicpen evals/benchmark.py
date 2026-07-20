import concurrent.futures
import threading
import logging
from collections import defaultdict
import config
import prompt_loader
import api_client

logger = logging.getLogger(__name__)

# Thread lock for clean console printing
print_lock = threading.Lock()

def run_single_benchmark_task(
    row_idx: int, 
    row: dict, 
    model: str, 
    system_template: str,
    user_template: str,
    total_rows: int, 
    results_dict: dict
) -> None:
    """
    Submits a single request to the API, measures latency/cost, and stores the results.
    """
    model_conf = config.MODELS_CONFIG.get(model, {})
    display_name = model_conf.get("display_name", model)
    
    # 1. Format prompt
    system_prompt, user_prompt = prompt_loader.format_prompts(system_template, user_template, row)
    
    # 2. Query model
    res = api_client.query_model_with_retry(model, system_prompt, user_prompt)
    
    # 3. Store result (thread-safe by writing to pre-allocated indices)
    results_dict[model][row_idx] = res
    
    # 4. Print console logs in a thread-safe manner
    with print_lock:
        tokens_str = "NA"
        if isinstance(res["prompt_tokens"], int) and isinstance(res["completion_tokens"], int):
            tokens_str = f"{res['prompt_tokens']} + {res['completion_tokens']}"
            
        cost_str = "NA"
        if isinstance(res["cost"], (int, float)):
            cost_str = f"${res['cost']:.8f}"
            
        print(f"\nRow {row_idx + 1}/{total_rows}")
        print(f"Testing {display_name}...")
        print(f"Latency: {res['latency']:.3f} sec")
        print(f"Tokens: {tokens_str}")
        print(f"Cost: {cost_str}")

def run_benchmark(rows: list, system_template: str, user_template: str) -> dict:
    """
    Runs the concurrent benchmark for all models across all input rows.
    """
    models = list(config.MODELS_CONFIG.keys())
    total_rows = len(rows)
    
    # Pre-allocate result lists to preserve input row ordering
    results = {model: [None] * total_rows for model in models}
    
    print(f"Starting benchmark: {total_rows} rows x {len(models)} models = {total_rows * len(models)} requests.")
    print(f"Parallelism (max_workers) set to: {config.MAX_WORKERS}\n")
    
    # Build task list: list of (row_idx, row, model)
    tasks = []
    for model in models:
        for row_idx, row in enumerate(rows):
            tasks.append((row_idx, row, model))
            
    with concurrent.futures.ThreadPoolExecutor(max_workers=config.MAX_WORKERS) as executor:
        futures = []
        for row_idx, row, model in tasks:
            f = executor.submit(
                run_single_benchmark_task,
                row_idx, row, model, system_template, user_template, total_rows, results
            )
            futures.append(f)
            
        # Wait for all threads to complete
        concurrent.futures.wait(futures)
        
    print("\nBenchmark runs completed.")
    return results
