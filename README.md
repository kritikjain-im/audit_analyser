# Quick Start Commands

## Setup
```bash
pip install -r requirements.txt
```

## 1. Direct Excel-to-Excel Analyzer
Run comparison on Excel workbook sheets:
```bash
python excel_analyzer/excel_analyzer.py --input-file "Langfuse Datasets.xlsx" --output-dir excel_analyzer/output
```

Run on custom sheets/ranges:
```bash
python excel_analyzer/excel_analyzer.py --input-file "Langfuse Datasets.xlsx" --manual-sheet "Manual Audits" --ai-sheet "AI Audits" --output-dir excel_analyzer/output
```

Run on a combined CSV file (splits manually/AI rows automatically):
```bash
python excel_analyzer/excel_analyzer.py --input-file "path/to/combined_audits.csv" --output-dir excel_analyzer/output
```

Run on separate CSV files:
```bash
python excel_analyzer/excel_analyzer.py --input-file "path/to/manual_audits.csv" --ai-sheet "path/to/ai_audits.csv" --output-dir excel_analyzer/output
```

---

## 2. OCR Screenshot Analyzer
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
