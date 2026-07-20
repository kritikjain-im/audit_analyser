import time
import logging
import openai
from openai import OpenAI
import config

logger = logging.getLogger(__name__)

# Initialize the OpenAI-compatible client
client = OpenAI(
    api_key=config.API_KEY,
    base_url=config.BASE_URL
)

def query_model_with_retry(model: str, system_prompt: str, user_prompt: str) -> dict:
    """
    Queries an LLM model with the given prompt.
    Implements retry mechanism with exponential backoff on specific error codes/timeouts.
    
    Returns a dictionary containing:
        - response: generated text or "ERROR"
        - latency: duration of the request in seconds
        - prompt_tokens: count or "NA"
        - completion_tokens: count or "NA"
        - total_tokens: count or "NA"
        - cost: estimated cost in USD or "NA"
    """
    attempts = config.RETRY_ATTEMPTS
    backoff = config.RETRY_BACKOFF_FACTOR
    
    # Retrieve pricing configs
    model_pricing = config.MODELS_CONFIG.get(model, {})
    input_price = model_pricing.get("input_price_per_million", 0.0)
    output_price = model_pricing.get("output_price_per_million", 0.0)
    
    last_err_msg = ""
    
    for attempt in range(1, attempts + 2):  # e.g., attempt 1, 2, 3, 4 (1 initial + 3 retries)
        start_time = time.perf_counter()
        try:
            # Query the chat completions endpoint
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                max_tokens=500,
                temperature=0.7,
                timeout=config.TIMEOUT
            )
            latency = time.perf_counter() - start_time
            
            # Extract content
            content = response.choices[0].message.content.strip()
            
            # Safely extract token usage
            try:
                usage = response.usage
                prompt_tokens = usage.prompt_tokens
                completion_tokens = usage.completion_tokens
                total_tokens = usage.total_tokens
            except AttributeError:
                prompt_tokens = "NA"
                completion_tokens = "NA"
                total_tokens = "NA"
                
            # Compute cost
            if isinstance(prompt_tokens, int) and isinstance(completion_tokens, int):
                cost = (prompt_tokens / 1_000_000 * input_price) + (completion_tokens / 1_000_000 * output_price)
            else:
                cost = "NA"
                
            return {
                "response": content,
                "latency": latency,
                "prompt_tokens": prompt_tokens,
                "completion_tokens": completion_tokens,
                "total_tokens": total_tokens,
                "cost": cost,
                "status": "Success"
            }
            
        except (openai.APITimeoutError, openai.APIConnectionError) as e:
            latency = time.perf_counter() - start_time
            last_err_msg = f"{type(e).__name__}: {str(e)}"
            if attempt > attempts:
                break
            logger.warning(f"Attempt {attempt} failed for model {model} (Connection/Timeout). Retrying in {backoff}s...")
            time.sleep(backoff)
            backoff *= 2
            
        except openai.APIStatusError as e:
            latency = time.perf_counter() - start_time
            last_err_msg = f"APIStatusError ({e.status_code}): {e.message}"
            if e.status_code not in config.RETRY_STATUS_CODES or attempt > attempts:
                break
            logger.warning(f"Attempt {attempt} failed for model {model} (HTTP {e.status_code}). Retrying in {backoff}s...")
            time.sleep(backoff)
            backoff *= 2
            
        except Exception as e:
            latency = time.perf_counter() - start_time
            last_err_msg = f"Unexpected Error: {str(e)}"
            break  # Don't retry on unrecognized/fatal code exceptions
            
    # If all retries failed
    logger.error(f"All {attempts + 1} attempts failed for model {model}. Last Error: {last_err_msg}")
    return {
        "response": "ERROR",
        "latency": latency,
        "prompt_tokens": "NA",
        "completion_tokens": "NA",
        "total_tokens": "NA",
        "cost": "NA",
        "status": "Failed"
    }
