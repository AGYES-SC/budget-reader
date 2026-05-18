#!/bin/bash

# Change to the folder this script lives in
cd "$(dirname "$0")"

# Install dependency if missing
python3 -c "import pdfplumber" 2>/dev/null || pip3 install pdfplumber

echo ""
echo "================================"
echo "  State Budget Bill Analyzer"
echo "================================"
echo ""
echo "Drag your PDF into this window and press Enter:"
read -r pdf_path

# Strip surrounding quotes (in case user drags file in)
pdf_path="${pdf_path%\'}"
pdf_path="${pdf_path#\'}"
pdf_path="${pdf_path%\"}"
pdf_path="${pdf_path#\"}"
# Trim whitespace
pdf_path="$(echo "$pdf_path" | xargs)"

if [ ! -f "$pdf_path" ]; then
  echo ""
  echo "Error: file not found — $pdf_path"
  echo "Press Enter to exit."
  read -r
  exit 1
fi

echo ""
echo "Running analysis..."
echo ""
python3 "$(dirname "$0")/budget_parser.py" "$pdf_path"

echo ""
echo "Press Enter to close."
read -r
