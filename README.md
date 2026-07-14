# Quick Start Commands

## Setup
```bash
pip install -r requirements.txt
```

---

## 1. Direct Excel-to-Excel Analyzer (`excel_analyzer.py`)

### How It Works
Compares AI audits against Manual audits by matching common `Call ID` values. It dynamically maps 12 key auditing fields from Excel worksheets or CSV sheets (filtering out sublabels and reason columns). It then:
- Validates the manual audit values against AI values using custom normalization badges.
- Assigns a match status (1 or 0) for each field.
- Generates a comparison sheet containing field-by-field accuracies and overall comparative score metrics.

### Commands
Run comparison on Excel workbook sheets:
```bash
python excel_analyzer/excel_analyzer.py --input-file "Langfuse Datasets.xlsx" --output-dir excel_analyzer/output
```

Run on custom sheets/ranges:
```bash
python excel_analyzer/excel_analyzer.py --input-file "Langfuse Datasets.xlsx" --manual-sheet "Manual Audits" --ai-sheet "AI Audits" --output-dir excel_analyzer/output
```

Run on a combined CSV file (splits manual/AI rows automatically):
```bash
python excel_analyzer/excel_analyzer.py --input-file "path/to/combined_audits.csv" --output-dir excel_analyzer/output
```

Run on separate CSV files:
```bash
python excel_analyzer/excel_analyzer.py --input-file "excel_analyzer/Manual Audits.csv" --ai-sheet "excel_analyzer/AI Audits.csv" --output-dir excel_analyzer/output
```

---

## 2. Category & Summary Analyzer 2.0 (`excel_analyzer_2.0.py`)

### How It Works
Processes a single audit CSV (either AI Audits or Manual Audits) to count category-specific occurrences (Correct, Correct (Not Discussed), Inferred, Incorrect, Missing, Hallucination) and compute field-level and overall accuracies.
- **Filtering:** Automatically filters the dataset to keep only calls where `Verdict` is `yes` or `no` (case-insensitive), skipping any rows where the verdict is `Skipped` or `NaN`.
- **Accuracy Definition:** Field accuracy is computed as `(Correct + Correct (Not Discussed) + Inferred) / Total calls`.
- **Overall Accuracy:** Appends a summary row containing the overall accuracy defined as `(Number of 'yes' verdicts) / Total calls`.
- **Output:** Outputs a styled Excel report containing 12 field rows plus the Overall Accuracy row with professional column formatting and highlight colors.

### Commands
Run summary on Manual Audits CSV:
```bash
python excel_analyzer/excel_analyzer_2.0.py -i "excel_analyzer/Manual Audits.csv"
```

Run summary on AI Audits CSV:
```bash
python excel_analyzer/excel_analyzer_2.0.py -i "excel_analyzer/AI Audits.csv"
```

Specifying a custom output path:
```bash
python excel_analyzer/excel_analyzer_2.0.py -i "excel_analyzer/AI Audits.csv" -o "excel_analyzer/output/custom_ai_summary.xlsx"
```

---

## 3. OCR Screenshot Analyzer (`ocr_analyzer/main.py`)

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

## 4. Git Helpers
Stage, commit, and push changes:
```bash
git add -A
git commit -m "Your commit message"
git push
```
