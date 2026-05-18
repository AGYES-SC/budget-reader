#!/bin/bash

# Change to the folder this script lives in
cd "$(dirname "$0")"

# Load shell profile so ANTHROPIC_API_KEY set via ~/.zshrc or ~/.bash_profile is available
source ~/.zshrc 2>/dev/null || source ~/.bash_profile 2>/dev/null || true

# Fall back to a .env file in the same folder if the key still isn't set
if [ -z "$ANTHROPIC_API_KEY" ] && [ -f "$(dirname "$0")/.env" ]; then
  export $(grep -v '^#' "$(dirname "$0")/.env" | xargs) 2>/dev/null
fi

# Install dependencies if missing
python3 -c "import pdfplumber" 2>/dev/null || pip3 install pdfplumber
python3 -c "import anthropic" 2>/dev/null || pip3 install anthropic

echo ""
echo "================================"
echo "  State Budget Bill Analyzer"
echo "================================"
echo ""

# Warn (don't block) if the API key is missing — pattern extraction still works without it
if [ -z "$ANTHROPIC_API_KEY" ]; then
  echo "⚠️  ANTHROPIC_API_KEY is not set."
  echo "   Pattern-based extraction will still run."
  echo "   To enable the AI fallback, see setup instructions in README.md."
  echo ""
fi
echo "Opening file picker — select your PDF file..."
pdf_path=$(osascript -e 'tell application "Finder"
  activate
end tell
tell application "System Events"
  set chosenFile to (choose file with prompt "Select a Budget PDF" of type {"pdf", "com.adobe.pdf"})
  return POSIX path of chosenFile
end tell' 2>/dev/null)

if [ -z "$pdf_path" ]; then
  echo ""
  echo "No file selected. Press Enter to exit."
  read -r
  exit 1
fi

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
