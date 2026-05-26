#!/usr/bin/env python3
"""
State Budget Bill Analyzer
Extracts agency/department appropriations from a PDF using OpenAI GPT-4o,
then produces a formatted Excel summary report.

Requires:
  - OPENAI_API_KEY environment variable
  - pip install openai pdfplumber openpyxl
"""

import sys
import os
import json
import re
from pathlib import Path
from datetime import datetime
import pdfplumber
import openai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Extraction (GPT-4o)
# ---------------------------------------------------------------------------

def extract_budget_data(pdf_path: str) -> dict:
    full_text_lines = []
    page_count = 0
    warnings = []

    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if not text.strip():
                warnings.append(f"Page {page_num}: no extractable text (may be scanned).")
                continue
            for line in text.split('\n'):
                full_text_lines.append(line)

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Error: OPENAI_API_KEY is not set.")
        print("Add it to ~/.zshrc:  export OPENAI_API_KEY=\"sk-...\"")
        sys.exit(1)

    client = openai.OpenAI(api_key=api_key)

    # Build page-aware chunks: group complete pages up to ~40K chars each.
    # This prevents an entity's header and dollar row from landing in different
    # chunks, which caused silent parse failures and missing totals.
    PAGE_CHUNK_LIMIT = 40_000
    chunks = []          # list of (label, text)
    current_lines = []
    current_len = 0
    chunk_start_page = 1

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, 1):
            page_text = page.extract_text() or ""
            if not page_text.strip():
                continue
            page_len = len(page_text)
            # If adding this page would exceed the limit, flush the current chunk first
            if current_lines and current_len + page_len > PAGE_CHUNK_LIMIT:
                label = f"pages {chunk_start_page}–{page_num - 1} of {total_pages}"
                chunks.append((label, '\n'.join(current_lines)))
                current_lines = []
                current_len = 0
                chunk_start_page = page_num
            current_lines.append(page_text)
            current_len += page_len
        if current_lines:
            label = f"pages {chunk_start_page}–{total_pages} of {total_pages}"
            chunks.append((label, '\n'.join(current_lines)))

    system_prompt = (
        "You are a state budget bill analyst. Extract top-level government appropriations "
        "from the provided budget bill text.\n\n"
        "Return ONLY valid JSON in exactly this shape:\n"
        "{\n"
        '  "fiscal_years": ["2026-27"],\n'
        '  "entities": {\n'
        '    "Department Of Education": {"2026-27": 123456789.0},\n'
        '    "Department Of Transportation": {"2026-27": 98765432.0}\n'
        "  }\n"
        "}\n\n"
        "INCLUDE: government departments, cabinets, agencies, boards, commissions, and bureaus "
        "that receive a direct legislative appropriation — i.e. entities with their own "
        "TOTAL or appropriation summary line in the bill.\n\n"
        "EXCLUDE — do not list these as separate rows:\n"
        "- Private organizations, nonprofits, clubs, foundations, ranches, or associations "
        "that receive grants or contracts through a department (e.g. 'Boys and Girls Clubs', "
        "'Alabama Sheriff's Youth Ranch', 'Heart Gallery Alabama').\n"
        "- Individual programs, projects, or line items that are sub-components of a "
        "department's total (e.g. 'Transportation Pilot Program' listed under Human Resources).\n"
        "- Grand-total or all-funds rollup lines that aggregate multiple departments.\n\n"
        "If a section lists a department total AND then itemizes grants or programs beneath it, "
        "include ONLY the department total row — not the individual grants.\n\n"
        "Other rules:\n"
        "- fiscal_years: fiscal-year labels in this chunk, formatted YYYY-YY (e.g. 2026-27).\n"
        "- Entity names in Title Case.\n"
        "- Dollar amounts as plain floats — no $ signs, no commas. "
        "If amounts are in thousands, multiply by 1000.\n"
        "- If no appropriations are found, return "
        '{"fiscal_years": [], "entities": {}}.'
    )

    # canonical_name maps lowercase-normalized name -> display name (first seen)
    # all_entities maps lowercase-normalized name -> {fy: amount}
    canonical_name: dict = {}
    all_entities: dict = {}
    all_fiscal_years: list = []

    print(f"  Extracting with GPT-4o — {len(chunks)} chunk(s)...")
    for idx, (label, chunk) in enumerate(chunks, 1):
        print(f"    Chunk {idx}/{len(chunks)}: {label}")
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                max_tokens=16384,
                temperature=0,
                seed=42,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Extract appropriations ({label}):\n\n{chunk}"},
                ],
            )
            finish_reason = response.choices[0].finish_reason
            raw = response.choices[0].message.content or ""

            if finish_reason == "length":
                warnings.append(
                    f"Chunk {idx} ({label}): response hit token limit — some entities may be missing. "
                    "Delete the cache file and re-run to retry."
                )
                print(f"    WARNING: chunk {idx} hit token limit — partial results only")

            data = json.loads(raw)

            for fy in data.get("fiscal_years", []):
                if fy not in all_fiscal_years:
                    all_fiscal_years.append(fy)

            for name, amounts in data.get("entities", {}).items():
                if not isinstance(amounts, dict):
                    continue
                # Normalize for deduplication so "Dept Of X" and "Dept of X"
                # (capitalization differences across chunks) merge into one row.
                key = name.strip().lower()
                if key not in canonical_name:
                    canonical_name[key] = name.strip()
                    all_entities[key] = {}
                for fy, amt in amounts.items():
                    try:
                        amt = float(amt)
                    except (TypeError, ValueError):
                        continue
                    # Keep the higher figure when the same entity appears in multiple chunks
                    if fy not in all_entities[key] or amt > all_entities[key][fy]:
                        all_entities[key][fy] = amt

        except Exception as exc:
            warnings.append(f"Chunk {idx} ({label}) error: {exc}")
            print(f"    ERROR on chunk {idx}: {exc}")

    fiscal_years = all_fiscal_years or ['Total']
    primary_fy = fiscal_years[0]

    # Rebuild with display names and drop zero-amount entities
    ordered_entities = {
        canonical_name[key]: amts
        for key, amts in all_entities.items()
        if any(v > 0 for v in amts.values())
    }

    grand_totals = {fy: sum(e.get(fy, 0) for e in ordered_entities.values()) for fy in fiscal_years}

    return {
        "entities": ordered_entities,
        "fiscal_years": fiscal_years,
        "grand_totals": grand_totals,
        "bill_grand_totals": {},
        "page_count": page_count,
        "warnings": warnings,
    }


# ---------------------------------------------------------------------------
# Excel Report
# ---------------------------------------------------------------------------

COLOR_HEADER_BG = "1F3864"
COLOR_HEADER_FG = "FFFFFF"
COLOR_TOTAL_BG  = "D9E1F2"
COLOR_META_FG   = "666666"
COLOR_ROW_EVEN  = "F5F7FA"
COLOR_ROW_ODD   = "FFFFFF"
COLOR_NOTES_FG  = "888888"

DOLLAR_FMT = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _border() -> Border:
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def generate_excel_report(data: dict, source_file: str, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Appropriations Summary"

    fys      = data["fiscal_years"]
    entities = data["entities"]
    grand_totals = data["grand_totals"]
    bill_totals  = data["bill_grand_totals"]
    num_cols = 1 + len(fys)

    # Row 1: Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    c = ws.cell(row=1, column=1, value="State Budget Bill — Appropriations Summary")
    c.font      = Font(name="Arial", size=16, bold=True, color=COLOR_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 2: Metadata
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    meta = (f"Source: {Path(source_file).name}   |   "
            f"Pages: {data['page_count']}   |   "
            f"Figures: all-funds TOTAL per entity")
    c = ws.cell(row=2, column=1, value=meta)
    c.font      = Font(name="Arial", size=10, italic=True, color=COLOR_META_FG)
    c.alignment = Alignment(horizontal="left")
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 6  # spacer

    # Row 4: Column headers
    ws.row_dimensions[4].height = 22
    c = ws.cell(row=4, column=1, value="Agency / Cabinet / Department")
    c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
    c.fill      = _fill(COLOR_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border    = _border()

    for col, fy in enumerate(fys, start=2):
        c = ws.cell(row=4, column=col, value=fy)
        c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
        c.fill      = _fill(COLOR_HEADER_BG)
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border    = _border()

    # Rows 5+: Entity data
    for offset, (name, amounts) in enumerate(entities.items()):
        r    = 5 + offset
        fill = COLOR_ROW_EVEN if offset % 2 == 0 else COLOR_ROW_ODD

        c = ws.cell(row=r, column=1, value=name)
        c.font      = Font(name="Arial", size=10)
        c.fill      = _fill(fill)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = _border()

        for col, fy in enumerate(fys, start=2):
            amt = amounts.get(fy, 0)
            c = ws.cell(row=r, column=col, value=amt if amt else None)
            c.font          = Font(name="Arial", size=10)
            c.fill          = _fill(fill)
            c.alignment     = Alignment(horizontal="right", vertical="center")
            c.border        = _border()
            c.number_format = DOLLAR_FMT

    next_row = 5 + len(entities)

    # Named-entity totals row
    c = ws.cell(row=next_row, column=1, value="Total — Named Entity Appropriations")
    c.font      = Font(name="Arial", size=11, bold=True)
    c.fill      = _fill(COLOR_TOTAL_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border    = _border()

    for col, fy in enumerate(fys, start=2):
        amt = grand_totals.get(fy, 0)
        c = ws.cell(row=next_row, column=col, value=amt if amt else None)
        c.font          = Font(name="Arial", size=11, bold=True)
        c.fill          = _fill(COLOR_TOTAL_BG)
        c.alignment     = Alignment(horizontal="right", vertical="center")
        c.border        = _border()
        c.number_format = DOLLAR_FMT

    next_row += 1

    # Bill's stated all-funds total (if present)
    if bill_totals:
        c = ws.cell(row=next_row, column=1,
                    value="Bill's Stated All-Funds Total (incl. bonds, transfers, other funds)")
        c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
        c.fill      = _fill(COLOR_HEADER_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = _border()

        for col, fy in enumerate(fys, start=2):
            amt = bill_totals.get(fy, 0)
            c = ws.cell(row=next_row, column=col, value=amt if amt else None)
            c.font          = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
            c.fill          = _fill(COLOR_HEADER_BG)
            c.alignment     = Alignment(horizontal="right", vertical="center")
            c.border        = _border()
            c.number_format = DOLLAR_FMT

        next_row += 1

    # Warnings
    if data["warnings"]:
        next_row += 1
        ws.merge_cells(start_row=next_row, start_column=1,
                       end_row=next_row, end_column=num_cols)
        c = ws.cell(row=next_row, column=1, value="Notes")
        c.font = Font(name="Arial", size=10, bold=True)
        next_row += 1
        for w in data["warnings"]:
            ws.merge_cells(start_row=next_row, start_column=1,
                           end_row=next_row, end_column=num_cols)
            c = ws.cell(row=next_row, column=1, value=f"• {w}")
            c.font      = Font(name="Arial", size=9, color=COLOR_NOTES_FG)
            c.alignment = Alignment(wrap_text=True)
            next_row += 1

    # Methodology note
    next_row += 1
    ws.merge_cells(start_row=next_row, start_column=1,
                   end_row=next_row, end_column=num_cols)
    c = ws.cell(row=next_row, column=1,
                value=(
                    "Methodology: Appropriations were extracted from the source PDF using AI (GPT-4o). "
                    "Each named agency, department, cabinet, or bureau is listed with its all-funds appropriation "
                    "for the identified fiscal year(s). Grand-total and rollup lines are excluded to prevent "
                    "double-counting. Verify all figures against the enrolled bill before citing."
                ))
    c.font      = Font(name="Arial", size=8, italic=True, color=COLOR_NOTES_FG)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[next_row].height = 54

    # Column widths and freeze
    ws.column_dimensions["A"].width = 52
    for col in range(2, num_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20
    ws.freeze_panes = "A5"

    wb.save(output_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python budget_parser.py <budget.pdf> [output.xlsx]")
        sys.exit(1)

    pdf_path = sys.argv[1]

    reports_dir = Path(__file__).parent / "Reports"
    reports_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = (sys.argv[2] if len(sys.argv) > 2
                else str(reports_dir / (Path(pdf_path).stem + f"_summary_{timestamp}.xlsx")))

    if not Path(pdf_path).exists():
        print(f"Error: file not found — {pdf_path}")
        sys.exit(1)

    cache_path = reports_dir / (Path(pdf_path).stem + "_extraction.json")

    print(f"Analyzing: {pdf_path}")
    if cache_path.exists():
        print(f"  Using cached extraction — delete {cache_path.name} to re-run AI extraction")
        with open(cache_path) as f:
            data = json.load(f)
    else:
        data = extract_budget_data(pdf_path)
        with open(cache_path, 'w') as f:
            json.dump(data, f, indent=2)

    print(f"\nFound {len(data['entities'])} budget entities")
    print(f"Fiscal years: {', '.join(data['fiscal_years'])}")
    for fy in data['fiscal_years']:
        print(f"Total {fy}: ${data['grand_totals'].get(fy, 0):,.0f}")
    if data['warnings']:
        print(f"\nWarnings ({len(data['warnings'])}):")
        for w in data['warnings']:
            print(f"  - {w}")

    print(f"\nGenerating report -> {out_path}")
    generate_excel_report(data, pdf_path, out_path)
    print("Done.")


if __name__ == '__main__':
    main()
