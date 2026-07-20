# Excel Call Audit Analyser

A python utility to compare manual and AI call audits.

## Setup

Ensure dependencies are installed:
```bash
pip install -r requirements.txt
```

## How to Run

Run the comparison on the default Excel workbook:
```bash
python excel_analyzer/excel_analyzer.py --input-file "Langfuse Datasets.xlsx" --output-dir excel_analyzer/output
```

Or run using custom manual and AI sheet names:
```bash
python excel_analyzer/excel_analyzer.py --input-file "Langfuse Datasets.xlsx" --manual-sheet "Manual Audits" --ai-sheet "AI Audits" --output-dir excel_analyzer/output
```
