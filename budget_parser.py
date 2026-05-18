#!/usr/bin/env python3
"""
State Budget Bill Analyzer
Extracts agency/department appropriations from a PDF and produces a summary report.

Extraction strategy:
  - Looks for "TOTAL - [ENTITY NAME]" header lines, which mark the authoritative
    summary for each cabinet, department, or agency in the bill.
  - Captures the TOTAL row (all-funds combined) for fiscal year columns that follow.
  - Falls back to a line-by-line TOTAL pattern for documents that lack "TOTAL - X" headers.
"""

import sys
import os
import re
import json
import subprocess
from pathlib import Path
import pdfplumber


# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

# "TOTAL - DEPARTMENT OF EDUCATION" or "TOTAL - GENERAL GOVERNMENT"
# Requires a letter immediately after the dash to exclude "TOTAL -0-" capital project lines
TOTAL_HEADER_RE = re.compile(
    r'TOTAL\s*[-\u2013]\s*([A-Z].+)',
    re.IGNORECASE
)

# A TOTAL row with one or more dollar figures on the same line
# e.g. "           TOTAL                                  6,976,000                6,895,000"
TOTAL_ROW_RE = re.compile(
    r'^(?:\d+\s+)?TOTAL\s',
    re.IGNORECASE
)

# Plain comma-grouped number
NUMBER_RE = re.compile(r'[\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?')

# Fiscal year column header, e.g. "2026-27" or "2027-28"
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

    # Detect fiscal year columns from early pages
    fiscal_years = []
    for line in full_text_lines[:300]:
        for y in YEAR_COL_RE.findall(line):
            if y not in fiscal_years:
                fiscal_years.append(y)
        if len(fiscal_years) >= 2:
            break

    # Find every "TOTAL - X" block and extract the TOTAL row beneath it
    entities = {}
    skip_terms = ['STATE/EXECUTIVE BUDGET', 'PHASE I TOBACCO', 'FUNDS TRANSFER']

    i = 0
    while i < len(full_text_lines):
        line = full_text_lines[i]
        m = TOTAL_HEADER_RE.search(line)

        if m:
            raw_name = m.group(1).strip()
            # Strip leading line numbers (e.g. "13   STATE/EXECUTIVE BUDGET")
            raw_name = re.sub(r'^\d+\s+', '', raw_name).strip()

            if any(t in raw_name.upper() for t in skip_terms):
                i += 1
                continue

            # Scan next 20 lines for the TOTAL (all-funds) row
            found = {}
            for j in range(i + 1, min(i + 20, len(full_text_lines))):
                scan = full_text_lines[j]
                if TOTAL_ROW_RE.match(scan):
                    nums = [n for n in NUMBER_RE.findall(scan) if parse_num(n) > 10000]
                    if nums and fiscal_years:
                        # Always take the rightmost N numbers where N = number of fiscal years.
                        # This safely drops any prior-year columns regardless of how many exist.
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
                    existing = sum(entities[name].values())
                    new = sum(found.values())
                    if new > existing:
                        entities[name] = found
                else:
                    entities[name] = found

        i += 1

    # Sort by primary fiscal year descending
    primary_fy = fiscal_years[0] if fiscal_years else 'Total'
    sorted_entities = dict(
        sorted(entities.items(), key=lambda x: x[1].get(primary_fy, 0), reverse=True)
    )

    grand_totals = {}
    for fy in (fiscal_years or ['Total']):
        grand_totals[fy] = sum(e.get(fy, 0) for e in sorted_entities.values())

    # Extract the bill's own stated all-funds grand total from the "TOTAL FUNDS" line
    bill_grand_totals = {}
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
        "fiscal_years": fiscal_years or ['Total'],
        "grand_totals": grand_totals,
        "bill_grand_totals": bill_grand_totals,
        "page_count": page_count,
        "warnings": warnings,
    }


# ---------------------------------------------------------------------------
# Report (DOCX via docx-js Node script)
# ---------------------------------------------------------------------------

REPORT_JS = r"""
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType
} = require('docx');

const data = JSON.parse(fs.readFileSync('/tmp/budget_data.json', 'utf8'));
const fys = data.fiscal_years;

function fmt(n) {
  if (!n || n === 0) return '-';
  return '$' + Math.round(n).toLocaleString('en-US');
}

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

const nameWidth = fys.length > 1 ? 5040 : 6240;
const fyWidth   = fys.length > 1 ? Math.floor(4320 / fys.length) : 3120;
const totalWidth = nameWidth + fyWidth * fys.length;

function headerCell(text, width) {
  return new TableCell({
    borders, margins: cm,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: '1F3864', type: ShadingType.CLEAR },
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20 })]
    })]
  });
}

function nameHeaderCell(text, width) {
  return new TableCell({
    borders, margins: cm,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: '1F3864', type: ShadingType.CLEAR },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20 })]
    })]
  });
}

const headerRow = new TableRow({
  tableHeader: true,
  children: [
    nameHeaderCell('Agency / Cabinet / Department', nameWidth),
    ...fys.map(fy => headerCell(fy, fyWidth))
  ]
});

const dataRows = Object.entries(data.entities).map(([name, totals], i) => {
  const fill = i % 2 === 0 ? 'F5F7FA' : 'FFFFFF';
  return new TableRow({
    children: [
      new TableCell({
        borders, margins: cm,
        width: { size: nameWidth, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: name, size: 18 })] })]
      }),
      ...fys.map(fy => new TableCell({
        borders, margins: cm,
        width: { size: fyWidth, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: fmt(totals[fy] || 0), size: 18 })]
        })]
      }))
    ]
  });
});

const totalRow = new TableRow({
  children: [
    new TableCell({
      borders, margins: cm,
      width: { size: nameWidth, type: WidthType.DXA },
      shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
      children: [new Paragraph({
        children: [new TextRun({ text: 'Total — Named Entity Appropriations', bold: true, size: 20 })]
      })]
    }),
    ...fys.map(fy => new TableCell({
      borders, margins: cm,
      width: { size: fyWidth, type: WidthType.DXA },
      shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: fmt(data.grand_totals[fy] || 0), bold: true, size: 20 })]
      })]
    }))
  ]
});

const billTotalRow = Object.keys(data.bill_grand_totals).length > 0
  ? new TableRow({
      children: [
        new TableCell({
          borders, margins: cm,
          width: { size: nameWidth, type: WidthType.DXA },
          shading: { fill: '1F3864', type: ShadingType.CLEAR },
          children: [new Paragraph({
            children: [new TextRun({ text: "Bill's Stated All-Funds Total (incl. bonds, transfers, other funds)", bold: true, color: 'FFFFFF', size: 20 })]
          })]
        }),
        ...fys.map(fy => new TableCell({
          borders, margins: cm,
          width: { size: fyWidth, type: WidthType.DXA },
          shading: { fill: '1F3864', type: ShadingType.CLEAR },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: fmt(data.bill_grand_totals[fy] || 0), bold: true, color: 'FFFFFF', size: 20 })]
          })]
        }))
      ]
    })
  : null;

const allRows = billTotalRow
  ? [headerRow, ...dataRows, totalRow, billTotalRow]
  : [headerRow, ...dataRows, totalRow];

const warningParagraphs = data.warnings.length > 0
  ? [
      new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: 'Notes', bold: true, size: 20 })] }),
      ...data.warnings.map(w => new Paragraph({ children: [new TextRun({ text: '• ' + w, size: 18, color: '666666' })] }))
    ]
  : [];

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: '1F3864' },
        paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: '2E5090' },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun('State Budget Bill — Appropriations Summary')]
      }),
      new Paragraph({
        spacing: { after: 160 },
        children: [new TextRun({
          text: `Source: ${data.source_file}   |   Pages: ${data.page_count}   |   Figures: all-funds TOTAL per entity`,
          size: 18, color: '666666', italics: true
        })]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun('Appropriations by Cabinet / Department / Agency')]
      }),
      new Table({
        width: { size: totalWidth, type: WidthType.DXA },
        columnWidths: [nameWidth, ...fys.map(() => fyWidth)],
        rows: allRows
      }),
      ...warningParagraphs,
      new Paragraph({
        spacing: { before: 300 },
        children: [new TextRun({
          text: 'Methodology: "Total — Named Entity Appropriations" is drawn exclusively from "TOTAL - [Entity]" ' +
                'summary lines, representing discrete all-funds appropriations to named cabinets, departments, ' +
                'and agencies. Line-item and sub-unit figures are excluded to prevent double-counting. ' +
                'The bill\'s Stated All-Funds Total is taken from the "TOTAL FUNDS" line in the bill\'s own ' +
                'grand summary and includes bond proceeds, intergovernmental transfers, investment income, ' +
                'and other funds not attributed to named entities. Verify against the enrolled bill before citing.',
          size: 16, color: '888888', italics: true
        })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2], buf);
  console.log('OK');
});
"""


def generate_report(data: dict, source_file: str, output_path: str):
    data["source_file"] = Path(source_file).name
    with open('/tmp/budget_data.json', 'w') as f:
        json.dump(data, f)
    with open('/tmp/make_report.js', 'w') as f:
        f.write(REPORT_JS)

    npm_root = subprocess.run(
        ['npm', 'root', '-g'], capture_output=True, text=True
    ).stdout.strip()
    env = {**os.environ, 'NODE_PATH': npm_root}

    result = subprocess.run(
        ['node', '/tmp/make_report.js', output_path],
        capture_output=True, text=True, env=env
    )
    if result.returncode != 0:
        raise RuntimeError(f"Report generation failed:\n{result.stderr}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python budget_parser.py <budget.pdf> [output.docx]")
        sys.exit(1)

    pdf_path = sys.argv[1]

    # Default output: Reports/ folder next to this script
    reports_dir = Path(__file__).parent / "Reports"
    reports_dir.mkdir(exist_ok=True)
    out_path = sys.argv[2] if len(sys.argv) > 2 else str(reports_dir / (Path(pdf_path).stem + "_summary.docx"))

    if not Path(pdf_path).exists():
        print(f"Error: file not found — {pdf_path}")
        sys.exit(1)

    print(f"Analyzing: {pdf_path}")
    data = extract_budget_data(pdf_path)

    print(f"\nFound {len(data['entities'])} budget entities")
    print(f"Fiscal years: {', '.join(data['fiscal_years'])}")
    for fy in data['fiscal_years']:
        print(f"Grand total {fy}: ${data['grand_totals'].get(fy, 0):,.0f}")
    if data['warnings']:
        print(f"\nWarnings ({len(data['warnings'])}):")
        for w in data['warnings']:
            print(f"  - {w}")

    print(f"\nGenerating report → {out_path}")
    generate_report(data, pdf_path, out_path)
    print("Done.")


if __name__ == '__main__':
    main()
