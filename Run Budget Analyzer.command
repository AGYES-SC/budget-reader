#!/bin/bash

# Change to the folder this script lives in
cd "$(dirname "$0")"

# Load shell profile so OPENAI_API_KEY set via ~/.zshrc or ~/.bash_profile is available
source ~/.zshrc 2>/dev/null || source ~/.bash_profile 2>/dev/null || true

# Fall back to a .env file in the same folder if the key still isn't set
if [ -z "$OPENAI_API_KEY" ] && [ -f "$(dirname "$0")/.env" ]; then
  export $(grep -v '^#' "$(dirname "$0")/.env" | xargs) 2>/dev/null
fi

# Install dependencies if missing
python3 -c "import pdfplumber" 2>/dev/null || pip3 install pdfplumber
python3 -c "import openai" 2>/dev/null || pip3 install openai
python3 -c "import openpyxl" 2>/dev/null || pip3 install openpyxl

echo ""
echo "================================"
echo "  State Budget Bill Analyzer"
echo "================================"
echo ""

if [ -z "$OPENAI_API_KEY" ]; then
  echo "ERROR: OPENAI_API_KEY is not set."
  echo "Add it to ~/.zshrc:  export OPENAI_API_KEY=\"sk-...\""
  echo ""
  echo "Press Enter to exit."
  read -r
  exit 1
fi

echo "Opening file picker — select your PDF..."
echo ""

pdf_path=$(osascript << 'APPLESCRIPT'
choose file with prompt "Select a Budget PDF"
POSIX path of result
APPLESCRIPT
)

if [ -z "$pdf_path" ]; then
  echo "No file selected. Press Enter to exit."
  read -r
  exit 1
fi

if [ ! -f "$pdf_path" ]; then
  echo "Error: file not found — $pdf_path"
  echo "Press Enter to exit."
  read -r
  exit 1
fi

echo "Selected: $pdf_path"
echo ""
echo "Running analysis..."
echo ""
python3 "$(dirname "$0")/budget_parser.py" "$pdf_path"

echo ""
echo "Press Enter to close."
read -r
