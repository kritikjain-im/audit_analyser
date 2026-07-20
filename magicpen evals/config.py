import os

# Base paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# API Credentials (configurable via environment variables, with defaults)
BASE_URL = os.getenv("MAGICPEN_BASE_URL", "https://imllm.intermesh.net/v1")
API_KEY = os.getenv("MAGICPEN_API_KEY", "sk-NGqP-r2T7W8MLLPDMwqThQ")

# File Configurations
INPUT_CSV_PATH = os.path.join(SCRIPT_DIR, "Magicpen Improv 2.0 - Sheet1.csv")
PROMPT_TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "magicpen_prompt_v2.txt")
OUTPUT_EXCEL_PATH = os.path.join(SCRIPT_DIR, "MagicPen_Model_Comparison.xlsx")

# Parallelism
MAX_WORKERS = 8

# Request Retries
RETRY_ATTEMPTS = 3
RETRY_BACKOFF_FACTOR = 2  # Exponential backoff base in seconds
RETRY_STATUS_CODES = {429, 500, 502, 503, 504}
TIMEOUT = 25  # Timeout in seconds

# Models Configuration
# To add, remove, or modify a model, edit this dictionary.
MODELS_CONFIG = {
    "google/gemini-2.5-flash-lite": {
        "display_name": "Gemini 2.5 Flash Lite",
        "excel_sheet_name": "Gemini Metrics",
        "excel_comparison_header": "Existing MagicPen Response (Gemini 2.5 Flash Lite)",
        "input_price_per_million": 0.075,
        "output_price_per_million": 0.30,
    },
    "google/gemini-3.5-flash": {
        "display_name": "Gemini 3.5 Flash",
        "excel_sheet_name": "Gemini 3.5 Flash Metrics",
        "excel_comparison_header": "Prompt Version 2 (Gemini 3.5 Flash)",
        "input_price_per_million": 0.075,
        "output_price_per_million": 0.30,
    },
    "openai/gpt-5.4-nano": {
        "display_name": "GPT-5.4 Nano",
        "excel_sheet_name": "GPT5 Nano Metrics",
        "excel_comparison_header": "Prompt Version 2 (GPT-5.4 Nano)",
        "input_price_per_million": 0.15,
        "output_price_per_million": 0.60,
    },
    "openai/gpt-5.4-mini": {
        "display_name": "GPT-5.4 Mini",
        "excel_sheet_name": "GPT5 Mini Metrics",
        "excel_comparison_header": "Prompt Version 2 (GPT-5.4 Mini)",
        "input_price_per_million": 0.30,
        "output_price_per_million": 1.20,
    },
    "google/gemma-4-26b-a4b-it": {
        "display_name": "Gemma 4 26B",
        "excel_sheet_name": "Gemma Metrics",
        "excel_comparison_header": "New Prompt Version 2 (Gemma 4 26B)",
        "input_price_per_million": 0.20,
        "output_price_per_million": 0.80,
    }
}
