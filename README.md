# Quick Start Commands

## Setup
```bash
pip install -r requirements.txt
```

---

## 1. Consolidated Call Audit Analyzer (`excel_analyzer.py`)

### How It Works
Parses and aggregates AI and Manual audit sheets or CSV datasets, filters them based on call `Verdict`, compares matching records by `Call ID`, and outputs a unified Excel workbook containing **7 detailed sheets**.

- **Verdict Filter:** Only processes calls where the `Verdict` is case-insensitively `'yes'` or `'no'` (omitting `Skipped` and empty records).
- **Default File Detection:** Automatically defaults to checking the `excel_analyzer/input` directory for `Manual Audits.csv` and `AI Audits.csv` if no custom inputs are provided.
- **Auto-Increment Output File:** Outputs to `excel_analyzer/output/Compact_Analysis_vX.xlsx` (auto-incrementing the version version code to prevent overwriting existing reports).

### Output Workbook Structure (7 Sheets)
1. **AI Audits Summary**: Field-by-field occurrences (Correct, ND, Inferred, Missing, Incorrect, Hallucination) and accuracies computed across all valid AI audits. Overall accuracy summary at the end.
2. **Manual Audits Summary**: Same compact summary layout but computed across all manual audits.
3. **AI Matched Summary**: AI audit summaries computed only on the overlapping calls that were manual-audited.
4. **Compact Overview**: Four summary tables (PNS call AI summary, Manual summary, field comparison with fraction/percentage formats, and comparison summary stats).
5. **Binary Comparison**: Call-by-call `1` (match) or `0` (mismatch) verification table.
6. **Detailed Value Comparison**: Call-by-call side-by-side values comparing Manual against AI audits.
7. **Summary Statistics**: Overall comparative metrics for categories (Correct, Inferred, ND, etc.) and matching scores.

### Commands

Run using default input paths (`excel_analyzer/input/Manual Audits.csv` and `excel_analyzer/input/AI Audits.csv`):
```bash
python excel_analyzer/excel_analyzer.py
```

Run specifying custom files or sheets:
```bash
python excel_analyzer/excel_analyzer.py -i "excel_analyzer/input/Manual Audits.csv" --ai-sheet "excel_analyzer/input/AI Audits.csv"
```

Run specifying custom Excel Workbook input:
```bash
python excel_analyzer/excel_analyzer.py -i "input/Langfuse Datasets.xlsx" --manual-sheet "Manual Audits" --ai-sheet "AI Audits"
```

---

## 2. OCR Screenshot Analyzer (`ocr_analyzer/main.py`)

### How It Works
Performs optical character recognition (OCR) on screenshots of seller calls to extract field values. It matches extracted fields with ground truth data to output accuracy reports.

### Commands
Run OCR comparison on a folder of screenshots:
```bash
python ocr_analyzer/main.py --input-dir ocr_analyzer/input --output-dir ocr_analyzer/output
```

Bypass cache and force-run OCR on all images:
```bash
python ocr_analyzer/main.py --force-refresh
```

---

## 3. Git Helpers
Stage, commit, and push changes:
```bash
git add -A
git commit -m "Your commit message"
git push
```
