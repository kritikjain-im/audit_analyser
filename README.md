# AI vs Manual Call Audit Analyzer

This Python application parses screenshots of call audits (PNG/JPG), compares the AI Audit against the Manual Audit across **12 distinct fields**, and automatically exports a beautifully formatted versioned Excel workbook.

It runs completely offline using a local OCR pipeline.

## Features

- **Local Offline OCR**: Uses `easyocr` to detect and extract evaluation badges from screenshots without requiring external API keys.
- **12-Field Comparison**: Support for all 12 audit fields (2-column, 6-row layout).
- **Match-based Binary Evaluation**: Calculates matching statuses as binary values (`1` for match, `0` for mismatch).
- **Auto-Incrementing File Versioning**: Prevents overwriting previous runs by outputting files like `AI_Audit_Analysis_v1.xlsx`, `AI_Audit_Analysis_v2.xlsx`, etc.
- **Three-Sheet Output Layout**:
  - **Comparative Analysis (Sheet 1)**: Color-coded binary matches, matches count, accuracy per call, and field-wise accuracies.
  - **Detailed Value Comparison (Sheet 2)**: Side-by-side comparison of raw text values filled by Manual and AI.
  - **Summary Statistics (Sheet 3)**: A compact summary showing counts of each parameter (`Correct`, `Correct (Not Discussed)`, `Incorrect`, `Missing`, `Hallucination`), matched calls, total calls, and average accuracies.
- **Professional Formatting**: Highlights matches/mismatches with soft colors, auto-fits column widths, freezes the header pane, and includes bold summary rows.

## Project Structure

```
auditor-summary-analyzer/
├── input/                  # Folder containing call audit screenshots
├── output/                 # Destination folder for generated versioned Excel reports
├── .cache/                 # Cached OCR JSON outputs to avoid re-processing identical files
├── .gitignore              # Git ignore configuration
├── main.py                 # Core runner, OCR processing, and comparison logic
├── excel_writer.py         # Styling and generating the Excel workbook using openpyxl
└── requirements.txt        # Python dependencies
```

## Setup Instructions

### 1. Prerequisites
Ensure you have **Python 3.11+** installed.

### 2. Install Dependencies
Install all required libraries:
```bash
pip install -r requirements.txt
```

---

## How to Run

Place your screenshots in the `input/` folder, then run the script:

```bash
python main.py
```

### Optional Arguments

- `--input-dir`: Custom path to folder of screenshots (default: `input`)
- `--output-dir`: Custom path for saving Excel files (default: `output`)
- `--force-refresh`: Ignore cache and run local OCR on all screenshots again

Example:
```bash
# Run with force refresh
python main.py --force-refresh
```
