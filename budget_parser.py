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
from collections import deque
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
    warnings = []

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Error: OPENAI_API_KEY is not set.")
        print("Add it to ~/.zshrc:  export OPENAI_API_KEY=\"sk-...\"")
        sys.exit(1)

    client = openai.OpenAI(api_key=api_key)

    # Build page-aware chunks (~40K chars each).
    # Chunks are stored as lists of (page_num, tagged_text) so they can be
    # split in half and retried automatically when GPT-4o hits the token limit.
    PAGE_CHUNK_LIMIT = 40_000
    initial_chunks = []   # list of [(page_num, tagged_text), ...]
    current_pages  = []
    current_len    = 0
    page_count     = 0

    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, 1):
            page_text = page.extract_text() or ""
            if not page_text.strip():
                warnings.append(f"Page {page_num}: no extractable text (may be scanned).")
                continue
            tagged = f"--- PAGE {page_num} ---\n{page_text}"
            tagged_len = len(tagged)
            if current_pages and current_len + tagged_len > PAGE_CHUNK_LIMIT:
                initial_chunks.append(current_pages)
                current_pages = []
                current_len   = 0
            current_pages.append((page_num, tagged))
            current_len += tagged_len
        if current_pages:
            initial_chunks.append(current_pages)

    system_prompt = (
        "You are a state budget bill analyst. Extract every named appropriation from the "
        "provided state budget bill text.\n\n"
        "Return ONLY valid JSON in exactly this shape:\n"
        "{\n"
        '  "fiscal_years": ["2026-27"],\n'
        '  "entities": {\n'
        '    "Department Of Education": {"2026-27": 123456789.0},\n'
        '    "Department Of Transportation": {"2026-27": 98765432.0}\n'
        "  }\n"
        "}\n\n"
        "CORE RULE — one row per entity, never double-count:\n"
        "A budget section for a single entity may show a TOTAL appropriation and then "
        "list how those funds are divided among programs, initiatives, or purposes. "
        "Record ONLY the entity's TOTAL. Do NOT separately list the program lines — "
        "they are sub-allocations of the total already counted above.\n"
        "Example: 'State Department of Education — $3.1B' followed by "
        "'Alabama Reading Initiative — $151M', 'Alabama Numeracy Act — $114M', etc. "
        "→ return ONLY 'State Department of Education: $3.1B'. "
        "The reading and numeracy programs are spending directions within that $3.1B, not separate appropriations.\n\n"
        "HOW TO TELL A SEPARATE APPROPRIATION FROM A SUB-ALLOCATION:\n"
        "A SEPARATE APPROPRIATION has its own section title and its own appropriation figure "
        "in the bill. Include it as its own row regardless of size or whether it shares a "
        "name prefix with another entity.\n"
        "  INCLUDE examples (each is its own section in the bill):\n"
        "  • 'Governor's Office — Custodial Fund   $653,000'\n"
        "  • 'Department of Agriculture — General Administration   $14,719,072'\n"
        "  • 'Department of Agriculture — Meat Inspection   $3,286,000'\n"
        "  The Agriculture examples above are PARALLEL divisions, each independently funded. "
        "Many state bills break a department into multiple separately-appropriated divisions or "
        "programs — each one is its own row, even though they share the department name.\n"
        "A SUB-ALLOCATION is a line item INSIDE another entity's section that shows how that "
        "entity's already-stated total is divided among internal purposes. It is NOT its own section.\n"
        "  EXCLUDE examples (these appear under a parent total, not as their own sections):\n"
        "  • Within 'State Dept of Education — Total $3.1B': sub-lines for "
        "'Reading Initiative $151M', 'Numeracy Act $114M' → breakdowns of the $3.1B, skip them.\n"
        "RULE OF THUMB: Own section title + own dollar figure = separate row. "
        "Line item inside another section's breakdown = sub-allocation, skip it.\n\n"
        "SEPARATE SECTIONS RULE — do not combine across sections:\n"
        "Some bills fund the same agency in multiple independent sections "
        "(e.g. a General Revenue Fund article AND a Special Revenue Fund article). "
        "If the same entity name carries its own separate appropriation in two or more "
        "distinct sections, return each as its own independent entry with its exact stated amount. "
        "Never add or merge figures from separate sections — return them as separate rows. "
        "This applies across sections only; within a single section still return just the entity's TOTAL.\n\n"
        "INCLUDE one row for each of these:\n"
        "- Government departments, agencies, boards, commissions, bureaus — at their TOTAL level\n"
        "- Universities and colleges at their TOTAL appropriation level\n"
        "- Named funds or authorities that receive a direct top-level appropriation\n"
        "- Miscellaneous/special sections — include each named item that has its own "
        "independent appropriation not already captured in a parent entity's total\n"
        "- Law enforcement, public safety, and corrections entities and their funds — "
        "sheriffs' programs, police funds, jail commissions, corrections facilities "
        "are government appropriations and must always be included\n\n"
        "EXCLUDE:\n"
        "- Individual persons' names (e.g. 'Smith, John')\n"
        "- Program lines, initiatives, or designated uses that appear WITHIN a named "
        "entity's budget section and are sub-allocations of that entity's total\n"
        "- Private nonprofit organizations that receive grants routed THROUGH a "
        "department which already has its own larger total line\n"
        "- Grand-total rollup lines that sum multiple departments "
        "(e.g. 'Total General Fund', 'All Funds Total')\n\n"
        "Other rules:\n"
        "- fiscal_years: fiscal-year labels in this chunk, formatted YYYY-YY (e.g. 2026-27).\n"
        "- Entity names in Title Case.\n"
        "- Dollar amounts as plain floats — no $ signs, no commas. "
        "If amounts are in thousands, multiply by 1000.\n"
        "- If no appropriations are found, return "
        '{"fiscal_years": [], "entities": {}}.'
    )

    # canonical_name: lowercase key -> display name (first seen)
    # all_entities:   lowercase key -> {fy: amount}
    # all_pages:      lowercase key -> page number (start page of chunk where first seen)
    canonical_name: dict = {}
    all_entities:   dict = {}
    all_pages:      dict = {}
    all_fiscal_years: list = []

    seen_exact: set = set()  # (lowercase_name, amounts_frozenset) — skips identical re-extractions
    # Use a work queue so chunks that hit the token limit can be split in half
    # and retried automatically without losing any entities.
    work_queue = deque(initial_chunks)
    print(f"  Extracting with GPT-4o — {len(initial_chunks)} chunk(s)...")

    while work_queue:
        pages = work_queue.popleft()
        start_page = pages[0][0]
        end_page   = pages[-1][0]
        label      = f"pages {start_page}–{end_page} of {page_count}"
        chunk_text = '\n'.join(text for _, text in pages)

        print(f"    Chunk: {label}")
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                max_tokens=16384,
                temperature=0,
                seed=42,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Extract appropriations ({label}):\n\n{chunk_text}"},
                ],
            )
            finish_reason = response.choices[0].finish_reason
            raw = response.choices[0].message.content or ""

            if finish_reason == "length":
                if len(pages) > 1:
                    # Split in half and re-queue — no entities are lost
                    mid = len(pages) // 2
                    work_queue.appendleft(pages[mid:])
                    work_queue.appendleft(pages[:mid])
                    print(f"      → token limit hit — splitting into {mid} + {len(pages) - mid} pages")
                    continue
                else:
                    # Single-page chunk; can't split further
                    warnings.append(
                        f"{label}: single-page chunk hit token limit — entities on this page may be incomplete."
                    )
                    print(f"      WARNING: single-page chunk hit token limit")

            data = json.loads(raw)

            for fy in data.get("fiscal_years", []):
                if fy not in all_fiscal_years:
                    all_fiscal_years.append(fy)

            chunk_entity_count = 0
            for name, amounts in data.get("entities", {}).items():
                if not isinstance(amounts, dict):
                    continue

                key = name.strip().lower()
                valid_amounts = {}
                for fy, amt in amounts.items():
                    try:
                        valid_amounts[fy] = float(amt)
                    except (TypeError, ValueError):
                        continue
                if not valid_amounts or not any(v > 0 for v in valid_amounts.values()):
                    continue

                # Identical (name + amounts) seen again = the same entity re-extracted
                # from context that spans a chunk boundary — skip the duplicate.
                # Same name with a DIFFERENT amount = the entity appears in a separate
                # bill section and must be kept as its own row.
                amounts_sig = frozenset(valid_amounts.items())
                if (key, amounts_sig) in seen_exact:
                    continue
                seen_exact.add((key, amounts_sig))

                if key not in canonical_name:
                    canonical_name[key] = name.strip()
                    all_entities[key] = valid_amounts
                    all_pages[key] = start_page
                else:
                    # Same name already recorded with a different amount: a separate
                    # appropriation elsewhere in the bill. Store under a unique internal
                    # key; display name gets a "(p. N)" suffix for disambiguation.
                    page_key = f"{key} [p{start_page}]"
                    counter = 2
                    while page_key in canonical_name:
                        page_key = f"{key} [p{start_page}:{counter}]"
                        counter += 1
                    canonical_name[page_key] = f"{name.strip()} (p. {start_page})"
                    all_entities[page_key] = valid_amounts
                    all_pages[page_key] = start_page

                chunk_entity_count += 1

            print(f"      → {chunk_entity_count} entities found")
            if chunk_entity_count == 0:
                warnings.append(
                    f"{label}: 0 entities returned — this page range may be missing from the report."
                )
                print("      WARNING: no entities found in this chunk")

        except Exception as exc:
            warnings.append(f"{label} error: {exc}")
            print(f"    ERROR: {exc}")

    fiscal_years = all_fiscal_years or ['Total']

    # ---------------------------------------------------------------------------
    # Post-processing filters (deterministic, not AI)
    # ---------------------------------------------------------------------------

    _personal_name_re = re.compile(r'^[A-Z][A-Za-z\-]+,\s+[A-Z][a-z]+$')
    _page_suffix_re   = re.compile(r'\s*\(p\.\s*[\d:]+\)\s*$')

    def _base_name(name: str) -> str:
        """Strip the '(p. N)' page-disambiguation suffix from same-name section variants."""
        return _page_suffix_re.sub('', name).strip()

    display_names = {k: canonical_name[k] for k in all_entities}

    def _word_set(name: str) -> frozenset:
        """Normalize to a word set ignoring punctuation/order, using the base name
        so page-suffix variants group correctly with their base entry."""
        return frozenset(re.sub(r'[,\-/]', ' ', _base_name(name).lower()).split())

    # Word-set deduplication: names that are word-order rearrangements of each other
    # (e.g. "Education, State Department Of" == "State Department Of Education") are
    # duplicates. Page-suffix variants ("[pN]" internal key) are intentional separate
    # appropriations — never merge them.
    word_set_groups: dict = {}
    for key in all_entities:
        ws = _word_set(canonical_name[key])
        word_set_groups.setdefault(ws, []).append(key)

    word_set_dupes: set = set()
    for keys_in_group in word_set_groups.values():
        if len(keys_in_group) > 1:
            # Only merge entries that don't have a page-suffix internal key — those are
            # intentional separate-section rows and must stay independent.
            base_keys = [k for k in keys_in_group if '[p' not in k]
            if len(base_keys) > 1:
                best = max(base_keys, key=lambda k: sum(all_entities[k].values()))
                for k in base_keys:
                    if k != best:
                        word_set_dupes.add(k)

    def _is_sub_section(key: str) -> bool:
        """True if another entity's base name is a substring of this entity's base name
        AND that other entity has a larger appropriation — indicating this is a sub-section.
        Uses base names to avoid false positives on page-suffix variants of the same entity."""
        name = _base_name(display_names[key]).lower()
        amt  = sum(all_entities[key].values())
        for other_key, other_name in display_names.items():
            if other_key == key:
                continue
            other_base = _base_name(other_name).lower()
            if other_base == name:
                continue  # Same base name (page-suffix variant) — not a sub-section
            if other_base in name and other_base != name:
                other_amt = sum(all_entities[other_key].values())
                if other_amt > amt:
                    return True
        return False

    def _is_fragment_duplicate(key: str) -> bool:
        """True if this entity's base name is a substring of a much larger entity's
        base name (10x+ amount) — a short fragment noise entry."""
        name = _base_name(display_names[key]).lower()
        amt  = sum(all_entities[key].values())
        for other_key, other_name in display_names.items():
            if other_key == key:
                continue
            other_base = _base_name(other_name).lower()
            if other_base == name:
                continue  # Same base name (page-suffix variant) — not a fragment
            if name in other_base and other_base != name:
                other_amt = sum(all_entities[other_key].values())
                if other_amt > amt * 10:
                    return True
        return False

    def _should_exclude(key: str) -> bool:
        name = canonical_name[key]
        amts = all_entities[key]
        if not any(v > 0 for v in amts.values()):
            return True
        if _personal_name_re.match(name):
            return True
        if key in word_set_dupes:
            return True
        if _is_sub_section(key):
            return True
        if _is_fragment_duplicate(key):
            return True
        return False

    ordered_entities = {
        canonical_name[key]: amts
        for key, amts in all_entities.items()
        if not _should_exclude(key)
    }

    # Page numbers keyed by display name, for entities that survived filtering
    entity_pages = {
        canonical_name[key]: all_pages.get(key)
        for key in all_entities
        if not _should_exclude(key)
    }

    grand_totals = {
        fy: sum(e.get(fy, 0) for e in ordered_entities.values())
        for fy in fiscal_years
    }

    return {
        "entities": ordered_entities,
        "entity_pages": entity_pages,
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

    fys          = data["fiscal_years"]
    entities     = data["entities"]
    entity_pages = data.get("entity_pages", {})
    grand_totals = data["grand_totals"]
    bill_totals  = data["bill_grand_totals"]

    # Column layout: [Name] [Page] [FY1] [FY2] ...
    PAGE_COL = 2
    FY_START = 3
    num_cols = 2 + len(fys)   # name + page + fiscal years

    # --- Row 1: Title ---------------------------------------------------------
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    c = ws.cell(row=1, column=1, value="State Budget Bill — Appropriations Summary")
    c.font      = Font(name="Arial", size=16, bold=True, color=COLOR_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    # --- Row 2: Metadata ------------------------------------------------------
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    meta = (f"Source: {Path(source_file).name}   |   "
            f"Pages: {data['page_count']}   |   "
            f"Figures: all-funds TOTAL per entity")
    c = ws.cell(row=2, column=1, value=meta)
    c.font      = Font(name="Arial", size=10, italic=True, color=COLOR_META_FG)
    c.alignment = Alignment(horizontal="left")
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 6  # spacer

    # --- Row 4: Column headers ------------------------------------------------
    ws.row_dimensions[4].height = 22

    c = ws.cell(row=4, column=1, value="Agency / Cabinet / Department")
    c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
    c.fill      = _fill(COLOR_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border    = _border()

    c = ws.cell(row=4, column=PAGE_COL, value="Page")
    c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
    c.fill      = _fill(COLOR_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = _border()

    for col, fy in enumerate(fys, start=FY_START):
        c = ws.cell(row=4, column=col, value=fy)
        c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
        c.fill      = _fill(COLOR_HEADER_BG)
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border    = _border()

    # --- Rows 5+: Entity data -------------------------------------------------
    for offset, (name, amounts) in enumerate(entities.items()):
        r    = 5 + offset
        fill = COLOR_ROW_EVEN if offset % 2 == 0 else COLOR_ROW_ODD
        page = entity_pages.get(name)

        c = ws.cell(row=r, column=1, value=name)
        c.font      = Font(name="Arial", size=10)
        c.fill      = _fill(fill)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = _border()

        c = ws.cell(row=r, column=PAGE_COL, value=page)
        c.font      = Font(name="Arial", size=10)
        c.fill      = _fill(fill)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _border()

        for col, fy in enumerate(fys, start=FY_START):
            amt = amounts.get(fy, 0)
            c = ws.cell(row=r, column=col, value=amt if amt else None)
            c.font          = Font(name="Arial", size=10)
            c.fill          = _fill(fill)
            c.alignment     = Alignment(horizontal="right", vertical="center")
            c.border        = _border()
            c.number_format = DOLLAR_FMT

    next_row = 5 + len(entities)

    # --- Named-entity totals row ----------------------------------------------
    c = ws.cell(row=next_row, column=1, value="Total — Named Entity Appropriations")
    c.font      = Font(name="Arial", size=11, bold=True)
    c.fill      = _fill(COLOR_TOTAL_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border    = _border()

    # blank page cell on totals row
    c = ws.cell(row=next_row, column=PAGE_COL, value=None)
    c.fill   = _fill(COLOR_TOTAL_BG)
    c.border = _border()

    for col, fy in enumerate(fys, start=FY_START):
        amt = grand_totals.get(fy, 0)
        c = ws.cell(row=next_row, column=col, value=amt if amt else None)
        c.font          = Font(name="Arial", size=11, bold=True)
        c.fill          = _fill(COLOR_TOTAL_BG)
        c.alignment     = Alignment(horizontal="right", vertical="center")
        c.border        = _border()
        c.number_format = DOLLAR_FMT

    next_row += 1

    # --- Bill's stated all-funds total (if present) ---------------------------
    if bill_totals:
        c = ws.cell(row=next_row, column=1,
                    value="Bill's Stated All-Funds Total (incl. bonds, transfers, other funds)")
        c.font      = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
        c.fill      = _fill(COLOR_HEADER_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = _border()

        c = ws.cell(row=next_row, column=PAGE_COL, value=None)
        c.fill   = _fill(COLOR_HEADER_BG)
        c.border = _border()

        for col, fy in enumerate(fys, start=FY_START):
            amt = bill_totals.get(fy, 0)
            c = ws.cell(row=next_row, column=col, value=amt if amt else None)
            c.font          = Font(name="Arial", size=11, bold=True, color=COLOR_HEADER_FG)
            c.fill          = _fill(COLOR_HEADER_BG)
            c.alignment     = Alignment(horizontal="right", vertical="center")
            c.border        = _border()
            c.number_format = DOLLAR_FMT

        next_row += 1

    # --- Warnings -------------------------------------------------------------
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

    # --- Methodology note -----------------------------------------------------
    next_row += 1
    ws.merge_cells(start_row=next_row, start_column=1,
                   end_row=next_row, end_column=num_cols)
    c = ws.cell(row=next_row, column=1,
                value=(
                    "Methodology: Appropriations were extracted from the source PDF using AI (GPT-4o). "
                    "Each named agency, department, cabinet, or bureau is listed with its all-funds appropriation "
                    "for the identified fiscal year(s). The Page column indicates the PDF page where the "
                    "appropriation appears. Grand-total and rollup lines are excluded to prevent "
                    "double-counting. Verify all figures against the enrolled bill before citing."
                ))
    c.font      = Font(name="Arial", size=8, italic=True, color=COLOR_NOTES_FG)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[next_row].height = 54

    # --- Column widths and freeze ---------------------------------------------
    ws.column_dimensions["A"].width = 52          # entity name
    ws.column_dimensions["B"].width = 7           # page number
    for col in range(FY_START, num_cols + 1):
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

    if not Path(pdf_path).exists():
        print(f"Error: file not found — {pdf_path}")
        sys.exit(1)

    reports_dir = Path(__file__).parent / "Reports"
    reports_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = (sys.argv[2] if len(sys.argv) > 2
                else str(reports_dir / (Path(pdf_path).stem + f"_summary_{timestamp}.xlsx")))

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
