#!/usr/bin/env python3
"""
State Budget Bill Analyzer
Extracts agency/department appropriations from a PDF and produces a formatted Excel report.

Requires:
  - pip install pdfplumber openpyxl
"""

import sys
import os
import re
from pathlib import Path
import pdfplumber
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

# "TOTAL - DEPARTMENT OF EDUCATION" — requires a letter after the dash to
# exclude "TOTAL -0-" capital project lines.
TOTAL_HEADER_RE = re.compile(
    r'TOTAL\s*[-–]\s*([A-Z].+)',
    re.IGNORECASE
)

# A TOTAL row that carries dollar figures, with optional leading line number.
TOTAL_ROW_RE = re.compile(
    r'^(?:\d+\s+)?TOTAL\s',
    re.IGNORECASE
)

# Comma-grouped numbers (e.g. 1,234,567 or 1,234,567.89)
NUMBER_RE = re.compile(r'[\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?')

# Fiscal year column header e.g. "2026-27"
YEAR_COL_RE = re.compile(r'\b(20\d\d-\d\d)\b')


def parse_num(s: str) -> float:
    return float(s.replace(',', '').strip())


# ---------------------------------------------------------------------------
# Extraction
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

    # Detect fiscal year columns from the first 300 lines of the document.
    fiscal_years = []
    for line in full_text_lines[:300]:
        for y in YEAR_COL_RE.findall(line):
            if y not in fiscal_years:
                fiscal_years.append(y)
        if len(fiscal_years) >= 2:
            break

    # Grand-summary labels that are NOT discrete entity appropriations.
    skip_terms = ['STATE/EXECUTIVE BUDGET', 'PHASE I TOBACCO', 'FUNDS TRANSFER']

    entities: dict = {}

    i = 0
    while i < len(full_text_lines):
        line = full_text_lines[i]
        m = TOTAL_HEADER_RE.search(line)

        if m:
            raw_name = re.sub(r'^\d+\s+', '', m.group(1).strip())

            if any(t in raw_name.upper() for t in skip_terms):
                i += 1
                continue

            # Scan the next 20 lines for the all-funds TOTAL row.
            # Take the rightmost N numbers (N = fiscal year count) to safely
            # drop any prior-year columns regardless of how many are present.
            found: dict = {}
            for j in range(i + 1, min(i + 20, len(full_text_lines))):
                scan = full_text_lines[j]
                if TOTAL_ROW_RE.match(scan):
                    nums = [n for n in NUMBER_RE.findall(scan) if parse_num(n) > 10_000]
                    if nums and fiscal_years:
                        nums = nums[-len(fiscal_years):]
                        for k, fy in enumerate(fiscal_years):
                            if k < len(nums):
                                found[fy] = parse_num(nums[k])
                    elif nums:
                        found['Total'] = parse_num(nums[-1])
                    break

            if found:
                name = raw_name.title()
                if name in entities:
                    # Keep the entry with the higher sum (handles duplicate appearances)
                    if sum(found.values()) > sum(entities[name].values()):
                        entities[name] = found
                else:
                    entities[name] = found

        i += 1

    primary_fy = fiscal_years[0] if fiscal_years else 'Total'
    sorted_entities = dict(
        sorted(entities.items(), key=lambda x: x[1].get(primary_fy, 0), reverse=True)
    )

    fy_list = fiscal_years or ['Total']
    grand_totals = {fy: sum(e.get(fy, 0) for e in sorted_entities.values()) for fy in fy_list}

    # Extract the bill's own stated all-funds grand total from "TOTAL FUNDS".
    bill_grand_totals: dict = {}
    in_budget_total = False
    for line in full_text_lines:
        if re.search(r'TOTAL\s*-\s*STATE/EXECUTIVE BUDGET', line, re.IGNORECASE):
            in_budget_total = True
        if in_budget_total and re.search(r'TOTAL FUNDS', line, re.IGNORECASE):
            nums = [n for n in NUMBER_RE.findall(line) if parse_num(n) > 1_000_000]
            if nums and fiscal_years:
                offset = len(nums) - len(fiscal_years)
                for k, fy in enumerate(fiscal_years):
                    idx = offset + k
                    if 0 <= idx < len(nums):
                        bill_grand_totals[fy] = parse_num(nums[idx])
            in_budget_total = False

    return {
        "entities": sorted_entities,
        "fiscal_years": fy_list,
        "grand_totals": grand_totals,
        "bill_grand_totals": bill_grand_totals,
        "page_count": page_count,
        "warnings": warnings,
    }


# ---------------------------------------------------------------------------
# Excel Report
# ---------------------------------------------------------------------------

# Color palette (hex, no #)
COLOR_HEADER_BG  = "1F3864"   # dark navy — column headers
COLOR_HEADER_FG  = "FFFFFF"
COLOR_TOTAL_BG   = "D9E1F2"   # soft blue — named-entity totals row
COLOR_BILL_BG    = "1F3864"   # same dark navy — bill's stated total row
COLOR_BILL_FG    = "FFFFFF"
COLOR_META_FG    = "666666"
COLOR_ROW_EVEN   = "F5F7FA"   # alternating row shading
COLOR_ROW_ODD    = "FFFFFF"
COLOR_NOTES_FG   = "888888"


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _border() -> Border:
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def generate_excel_report(data: dict, source_file: str, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Appropriations Summary"

    fys = data["fiscal_years"]
    entities = data["entities"]
    grand_totals = data["grand_totals"]
    bill_totals = data["bill_grand_totals"]

    num_cols = 1 + len(fys)          # name column + one per fiscal year
    name_col_width = 52
    amt_col_width = 20

    # -- Row 1: Title ----------------------------------------------------------
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    title_cell = ws.cell(row=1, column=1,
                         value="State Budget Bill — Appropriations Summary")
    title_cell.font = Font(name="Arial", size=16, bold=True, color=COLOR_HEADER_BG)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    # -- Row 2: Metadata -------------------------------------------------------
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    meta = (f"Source: {Path(source_file).name}   |   "
            f"Pages: {data['page_count']}   |   "
            f"Figures: all-funds TOTAL per entity")
    meta_cell = ws.cell(row=2, column=1, value=meta)
    meta_cell.font = Font(name="Arial", size=10, italic=True, color=COLOR_META_FG)
    meta_cell.alignment = Alignment(horizontal="left")
    ws.row_dimensions[2].height = 16

    # -- Row 3: blank ----------------------------------------------------------
    ws.row_dimensions[3].height = 6

    # -- Row 4: Column headers -------------------------------------------------
    hdr_row = 4
    ws.row_dimensions[hdr_row].height = 22

    name_hdr = ws.cell(row=hdr_row, column=1, value="Agency / Cabinet / Department")
    name_hdr.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
    name_hdr.fill      = _fill(COLOR_HEADER_BG)
    name_hdr.alignment = Alignment(horizontal="left", vertical="center")
    name_hdr.border    = _border()

    for col_idx, fy in enumerate(fys, start=2):
        cell = ws.cell(row=hdr_row, column=col_idx, value=fy)
        cell.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
        cell.fill      = _fill(COLOR_HEADER_BG)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border    = _border()

    # -- Rows 5+: Entity data rows ---------------------------------------------
    data_start_row = 5
    for row_offset, (name, amounts) in enumerate(entities.items()):
        r = data_start_row + row_offset
        fill_color = COLOR_ROW_EVEN if row_offset % 2 == 0 else COLOR_ROW_ODD

        name_cell = ws.cell(row=r, column=1, value=name)
        name_cell.font      = Font(name="Arial", size=10)
        name_cell.fill      = _fill(fill_color)
        name_cell.alignment = Alignment(horizontal="left", vertical="center")
        name_cell.border    = _border()

        for col_idx, fy in enumerate(fys, start=2):
            amt = amounts.get(fy, 0)
            cell = ws.cell(row=r, column=col_idx, value=amt if amt else None)
            cell.font         = Font(name="Arial", size=10)
            cell.fill         = _fill(fill_color)
            cell.alignment    = Alignment(horizontal="right", vertical="center")
            cell.border       = _border()
            cell.number_format = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'

    next_row = data_start_row + len(entities)

    # -- Named-entity totals row -----------------------------------------------
    total_name = ws.cell(row=next_row, column=1,
                         value="Total — Named Entity Appropriations")
    total_name.font      = Font(name="Arial", size=11, bold=True)
    total_name.fill      = _fill(COLOR_TOTAL_BG)
    total_name.alignment = Alignment(horizontal="left", vertical="center")
    total_name.border    = _border()

    for col_idx, fy in enumerate(fys, start=2):
        amt = grand_totals.get(fy, 0)
        cell = ws.cell(row=next_row, column=col_idx, value=amt if amt else None)
        cell.font          = Font(name="Arial", size=11, bold=True)
        cell.fill          = _fill(COLOR_TOTAL_BG)
        cell.alignment     = Alignment(horizontal="right", vertical="center")
        cell.border        = _border()
        cell.number_format = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'

    next_row += 1

    # -- Bill's stated all-funds total row (if available) ----------------------
    if bill_totals:
        bill_name = ws.cell(row=next_row, column=1,
                            value="Bill’s Stated All-Funds Total (incl. bonds, transfers, other funds)")
        bill_name.font      = Font(name="Arial", size=11, bold=True, color=COLOR_BILL_FG)
        bill_name.fill      = _fill(COLOR_BILL_BG)
        bill_name.alignment = Alignment(horizontal="left", vertical="center")
        bill_name.border    = _border()

        for col_idx, fy in enumerate(fys, start=2):
            amt = bill_totals.get(fy, 0)
            cell = ws.cell(row=next_row, column=col_idx, value=amt if amt else None)
            cell.font          = Font(name="Arial", size=11, bold=True, color=COLOR_BILL_FG)
            cell.fill          = _fill(COLOR_BILL_BG)
            cell.alignment     = Alignment(horizontal="right", vertical="center")
            cell.border        = _border()
            cell.number_format = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'

        next_row += 1

    # -- Warnings / Notes section ----------------------------------------------
    if data["warnings"]:
        next_row += 1
        ws.merge_cells(start_row=next_row, start_column=1,
                       end_row=next_row, end_column=num_cols)
        notes_hdr = ws.cell(row=next_row, column=1, value="Notes")
        notes_hdr.font = Font(name="Arial", size=10, bold=True)
        next_row += 1
        for w in data["warnings"]:
            ws.merge_cells(start_row=next_row, start_column=1,
                           end_row=next_row, end_column=num_cols)
            note_cell = ws.cell(row=next_row, column=1, value=f"• {w}")
            note_cell.font      = Font(name="Arial", size=9, color=COLOR_NOTES_FG)
            note_cell.alignment = Alignment(wrap_text=True)
            next_row += 1

    # -- Methodology note ------------------------------------------------------
    next_row += 1
    ws.merge_cells(start_row=next_row, start_column=1,
                   end_row=next_row, end_column=num_cols)
    methodology = (
        "Methodology: Appropriations are drawn exclusively from “TOTAL – [Entity]” summary "
        "lines in the PDF, representing discrete all-funds appropriations to named cabinets, departments, "
        "and agencies. Line-item and sub-unit figures are excluded to prevent double-counting. "
        "The Bill’s Stated All-Funds Total is taken from the “TOTAL FUNDS” line in the "
        "bill’s own grand summary and includes bond proceeds, intergovernmental transfers, "
        "investment income, and other funds not attributed to named entities. "
        "Verify all figures against the enrolled bill before citing."
    )
    meth_cell = ws.cell(row=next_row, column=1, value=methodology)
    meth_cell.font      = Font(name="Arial", size=8, italic=True, color=COLOR_NOTES_FG)
    meth_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[next_row].height = 54

    # -- Column widths ---------------------------------------------------------
    ws.column_dimensions["A"].width = name_col_width
    for col_idx in range(2, num_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = amt_col_width

    # -- Freeze panes (keep header row visible while scrolling) ----------------
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

    if not Path(pdf_path).exists():
        print(f"Error: file not found — {pdf_path}")
        sys.exit(1)

    reports_dir = Path(__file__).parent / "Reports"
    reports_dir.mkdir(exist_ok=True)
    out_path = (sys.argv[2] if len(sys.argv) > 2
                else str(reports_dir / (Path(pdf_path).stem + "_summary.xlsx")))

    print(f"Analyzing: {pdf_path}")
    data = extract_budget_data(pdf_path)

    print(f"\nFound {len(data['entities'])} budget entities")
    print(f"Fiscal years: {', '.join(data['fiscal_years'])}")
    for fy in data['fiscal_years']:
        print(f"  Named entity total {fy}: ${data['grand_totals'].get(fy, 0):,.0f}")
    if data['bill_grand_totals']:
        for fy in data['fiscal_years']:
            print(f"  Bill all-funds total {fy}: ${data['bill_grand_totals'].get(fy, 0):,.0f}")
    if data['warnings']:
        print(f"\nWarnings ({len(data['warnings'])}):")
        for w in data['warnings']:
            print(f"  - {w}")

    print(f"\nGenerating report -> {out_path}")
    generate_excel_report(data, pdf_path, out_path)
    print("Done.")


if __name__ == '__main__':
    main()
