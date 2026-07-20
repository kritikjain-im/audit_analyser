import math
import statistics
import pandas as pd

def normalize_value(val) -> str:
    """
    Strips whitespace, handles NaN, null, and formats as clean text.
    """
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.lower() in {"nan", "null", "none", "<na>"}:
        return ""
    return val_str

def calculate_mean(values) -> float:
    """
    Calculates the arithmetic mean of a list of numbers.
    """
    valid_vals = [v for v in values if isinstance(v, (int, float)) and not pd.isna(v)]
    if not valid_vals:
        return 0.0
    return sum(valid_vals) / len(valid_vals)

def calculate_percentile(values, percentile: float) -> float:
    """
    Calculates the percentile using linear interpolation.
    percentile is a float between 0.0 and 1.0 (e.g. 0.95 for P95).
    """
    valid_vals = sorted([v for v in values if isinstance(v, (int, float)) and not pd.isna(v)])
    if not valid_vals:
        return 0.0
    k = (len(valid_vals) - 1) * percentile
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return valid_vals[int(k)]
    d0 = valid_vals[int(f)] * (c - k)
    d1 = valid_vals[int(c)] * (k - f)
    return d0 + d1

def calculate_stddev(values) -> float:
    """
    Calculates the sample standard deviation.
    """
    valid_vals = [v for v in values if isinstance(v, (int, float)) and not pd.isna(v)]
    if len(valid_vals) < 2:
        return 0.0
    return statistics.stdev(valid_vals)

def calculate_min(values) -> float:
    """
    Calculates the minimum value.
    """
    valid_vals = [v for v in values if isinstance(v, (int, float)) and not pd.isna(v)]
    if not valid_vals:
        return 0.0
    return min(valid_vals)

def calculate_max(values) -> float:
    """
    Calculates the maximum value.
    """
    valid_vals = [v for v in values if isinstance(v, (int, float)) and not pd.isna(v)]
    if not valid_vals:
        return 0.0
    return max(valid_vals)
