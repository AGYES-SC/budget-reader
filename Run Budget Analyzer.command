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
echo "Opening file picker — select your PDF..."
echo ""

pdf_path=$(osascript -e 'tell application "System Events"
  set chosenFile to (choose file with prompt "Select a Budget PDF" of type {"com.adobe.pdf"})
  return POSIX path of chosenFile
end tell' 2>/dev/null)

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
