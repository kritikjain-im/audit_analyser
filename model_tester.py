import csv
import concurrent.futures
import requests
from collections import Counter

# =====================================================
# CONFIG
# =====================================================

BASE_URL = "https://imllm.intermesh.net/v1"
API_KEY = "sk-NGqP-r2T7W8MLLPDMwqThQ"

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

MODELS_URL = (
    f"{BASE_URL}/models"
    "?return_wildcard_routes=false"
    "&include_model_access_groups=false"
    "&only_model_access_groups=false"
    "&include_metadata=false"
)

CHAT_URL = f"{BASE_URL}/chat/completions"

MAX_WORKERS = 10
TIMEOUT = 25

# =====================================================
# FILTERS
# =====================================================

SKIP_KEYWORDS = [
    "embedding",
    "embed",
    "rerank",
    "image",
    "vision",
    "tts",
    "stt",
    "speech",
    "audio",
    "transcribe",
    "moderation",
    "whisper",
    "dall",
    "flux",
]

# =====================================================
# FETCH MODELS
# =====================================================

print("Fetching models...")

resp = requests.get(MODELS_URL, headers=HEADERS)
resp.raise_for_status()

data = resp.json()["data"]

models = []

for m in data:

    model = m["id"].lower()

    if any(x in model for x in SKIP_KEYWORDS):
        continue

    models.append(m["id"])

models = sorted(set(models))

print(f"Testing {len(models)} candidate chat models.\n")

# =====================================================
# ERROR CLASSIFIER
# =====================================================

def classify(msg):

    s = msg.lower()

    if "authentication" in s or "api_key" in s:
        return "Authentication"

    if "permission" in s or "forbidden" in s:
        return "Permission"

    if "rate limit" in s or "429" in s:
        return "RateLimit"

    if "timeout" in s:
        return "Timeout"

    if "invalid" in s:
        return "InvalidRequest"

    return "Unknown"

# =====================================================
# TEST ONE MODEL
# =====================================================

def test_model(model):

    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": "Reply ONLY with OK."
            }
        ],
        "max_tokens": 2,
        "temperature": 0
    }

    try:

        r = requests.post(
            CHAT_URL,
            headers=HEADERS,
            json=payload,
            timeout=TIMEOUT,
        )

        if r.status_code == 200:
            return {
                "model": model,
                "status": "Healthy",
                "error_type": "",
                "error": "",
            }

        try:
            err = r.json()["error"]["message"]
        except Exception:
            err = r.text

        return {
            "model": model,
            "status": "Broken",
            "error_type": classify(err),
            "error": err[:500],
        }

    except Exception as e:

        err = str(e)

        return {
            "model": model,
            "status": "Broken",
            "error_type": classify(err),
            "error": err[:500],
        }

# =====================================================
# RUN
# =====================================================

healthy = []
broken = []

with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

    futures = {
        executor.submit(test_model, m): m
        for m in models
    }

    total = len(futures)

    for idx, future in enumerate(
        concurrent.futures.as_completed(futures), 1
    ):

        result = future.result()

        print(
            f"[{idx}/{total}]",
            result["model"],
            "->",
            result["status"],
        )

        if result["status"] == "Healthy":
            healthy.append(result)
        else:
            broken.append(result)

# =====================================================
# SAVE CSVs
# =====================================================

healthy.sort(key=lambda x: x["model"])
broken.sort(key=lambda x: x["model"])

with open(
    "healthy_models.csv",
    "w",
    newline="",
    encoding="utf-8",
) as f:

    writer = csv.DictWriter(
        f,
        fieldnames=["model"]
    )

    writer.writeheader()

    for row in healthy:
        writer.writerow({"model": row["model"]})

with open(
    "broken_models.csv",
    "w",
    newline="",
    encoding="utf-8",
) as f:

    writer = csv.DictWriter(
        f,
        fieldnames=[
            "model",
            "error_type",
            "error",
        ],
        extrasaction="ignore",
    )

    writer.writeheader()

    writer.writerows(broken)

# =====================================================
# SUMMARY
# =====================================================

counter = Counter([x["error_type"] for x in broken])

with open("summary.txt", "w") as f:

    f.write(f"Healthy models : {len(healthy)}\n")
    f.write(f"Broken models  : {len(broken)}\n\n")

    f.write("Breakdown:\n")

    for k, v in counter.items():
        f.write(f"{k}: {v}\n")

print("\n==============================")
print(f"Healthy : {len(healthy)}")
print(f"Broken  : {len(broken)}")
print("==============================")
print("Generated:")
print("  healthy_models.csv")
print("  broken_models.csv")
print("  summary.txt")
