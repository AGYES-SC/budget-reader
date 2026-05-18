# Budget Reader

A local Python tool that parses state budget bill PDFs and produces a clean, accurate appropriations summary as a Word document.

## What It Does

- Extracts **authoritative appropriation totals** per cabinet, department, and agency from a budget bill PDF
- Reports two fiscal year columns side by side
- Includes both the **named-entity appropriations total** (sum of all discrete department allocations) and the **bill's stated all-funds total** (including bonds, transfers, and other funds), with a note explaining the difference
- Outputs a formatted `.docx` report to a `Reports/` subfolder automatically

## How It Works

The parser targets `TOTAL - [ENTITY NAME]` summary lines in the bill — the lines the legislature itself uses to summarize each cabinet or department's appropriation. It reads only those lines, ignoring all line-item and sub-unit figures, to prevent double-counting. Prior-year columns are detected and dropped automatically so figures always align to the correct fiscal years.

## Requirements

- Python 3.9+
- Node.js 16+
- `pdfplumber` Python package
- `docx` Node package (installed globally)

### Install dependencies

```bash
pip3 install pdfplumber
sudo npm install -g docx
```

## Usage

### Double-click launcher (recommended)

Run `Run Budget Analyzer.command` from the project folder. Drag your PDF into the Terminal window when prompted and press Enter. The report saves to `Reports/`.

### Command line

```bash
python3 budget_parser.py path/to/budget.pdf
```

Output saves to `Reports/<filename>_summary.docx` by default. To specify a custom output path:

```bash
python3 budget_parser.py path/to/budget.pdf path/to/output.docx
```

## Output

The generated `.docx` contains:

- A table of all cabinets and departments ranked by appropriation, with one column per fiscal year
- A **Named Entity Appropriations** subtotal row
- The **Bill's Stated All-Funds Total** row (from the bill's own grand summary)
- A methodology note explaining both figures

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

Claude Code reads the project files directly and can modify, debug, or extend the script based on your instructions.# budget-reader
