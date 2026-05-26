# Budget Reader

A local Python tool that parses state budget bill PDFs and produces a clean, accurate appropriations summary as an Excel workbook.

## What It Does

- Extracts **authoritative appropriation totals** per cabinet, department, and agency from a budget bill PDF
- Reports two fiscal year columns side by side
- Includes both the **named-entity appropriations total** (sum of all discrete department allocations) and the **bill's stated all-funds total** (including bonds, transfers, and other funds), with a note explaining the difference
- Outputs a formatted `.xlsx` Excel workbook to a `Reports/` subfolder automatically

## How It Works

The parser targets `TOTAL - [ENTITY NAME]` summary lines in the bill — the lines the legislature itself uses to summarize each cabinet or department's appropriation. It reads only those lines, ignoring all line-item and sub-unit figures, to prevent double-counting. Prior-year columns are detected and dropped automatically so figures always align to the correct fiscal years.

No AI or API calls are made — extraction is done entirely by regex pattern matching against the raw PDF text, so nothing is chunked, truncated, or misinterpreted.

## Requirements

- Python 3.9+
- `pdfplumber` Python package
- `openpyxl` Python package

### Install dependencies

```bash
pip3 install pdfplumber openpyxl
```

## Usage

### Double-click launcher (recommended)

Run `Run Budget Analyzer.command` from the project folder. Drag your PDF into the Terminal window when prompted and press Enter. The report saves to `Reports/`.

### Command line

```bash
python3 budget_parser.py path/to/budget.pdf
```

Output saves to `Reports/<filename>_summary.xlsx` by default. To specify a custom output path:

```bash
python3 budget_parser.py path/to/budget.pdf path/to/output.xlsx
```

## Output

The generated `.xlsx` contains:

- A table of all cabinets and departments ranked by appropriation, with one column per fiscal year
- Dollar amounts formatted as `$1,234,567` — native Excel number format, no text conversion
- A **Named Entity Appropriations** subtotal row
- The **Bill's Stated All-Funds Total** row (from the bill's own grand summary, if present)
- A methodology note explaining both figures
- Frozen header row so column labels stay visible while scrolling

## Project Structure

```
Budget Reader/
├── budget_parser.py           # Main script
├── Run Budget Analyzer.command  # Double-click launcher for macOS
├── README.md
├── .gitignore
└── Reports/                   # Generated reports (git-ignored)
```

## Notes on Accuracy

This tool is designed for bills that use `TOTAL - [ENTITY]` summary lines — the standard format used by most state legislatures. Bills with non-standard formatting may require adjustments to the extraction logic. Always verify figures against the enrolled bill before citing.

## Using with Claude Code

From the project folder:

```bash
claude
```

Claude Code reads the project files directly and can modify, debug, or extend the script based on your instructions.
