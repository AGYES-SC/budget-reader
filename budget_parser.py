#!/usr/bin/env python3
"""
State Budget Bill Analyzer
Extracts agency/department appropriations from a PDF using OpenAI GPT-4o,
then produces a formatted Word document summary report.

Requires:
  - OPENAI_API_KEY environment variable
  - pip install openai pdfplumber
"""

import sys
import os
import json
import subprocess
from pathlib import Path
from datetime import datetime
import pdfplumber
import openai


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

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Error: OPENAI_API_KEY is not set.")
        print("Add it to ~/.zshrc:  export OPENAI_API_KEY=\"sk-...\"")
        sys.exit(1)

    client = openai.OpenAI(api_key=api_key)
    full_text = '\n'.join(full_text_lines)

    # Chunk into ~80K-char pieces so each fits comfortably in a single request
    chunk_size = 80_000
    chunks = [full_text[i:i + chunk_size] for i in range(0, len(full_text), chunk_size)]

    system_prompt = (
        "You are a state budget bill analyst. Extract appropriations from the provided budget "
        "bill text following these exact rules:\n\n"
        "NON-EDUCATION AGENCIES: Include every named agency, department, cabinet, or bureau "
        "that is NOT education-related. Include exactly ONE row per agency showing its total "
        "appropriation. Do not include sub-items or line items for these agencies. Do not skip "
        "any non-education agency.\n\n"
        "EDUCATION AGENCIES: For any appropriation related to education (K-12, school districts, "
        "universities, colleges, community colleges, vocational education, Board of Education, "
        "Department of Education, or any education program), list EVERY individual line item as "
        "a separate row. Do NOT include the parent agency total — only the individual line items "
        "(to avoid double-counting). Name each line item as "
        "'Parent Agency — Line Item' (e.g. 'Department Of Education — Special Education', "
        "'University Of Kansas — General Operations').\n\n"
        "Return ONLY valid JSON in exactly this shape:\n"
        "{\n"
        '  "fiscal_years": ["2026-27"],\n'
        '  "entities": {\n'
        '    "Department Of Transportation": {"2026-27": 98765432.0},\n'
        '    "Department Of Education — Special Education": {"2026-27": 45000000.0},\n'
        '    "Department Of Education — Elementary Education": {"2026-27": 32000000.0}\n'
        "  },\n"
        '  "fund_sources": {\n'
        '    "Department Of Transportation": "State Highway Fund",\n'
        '    "Department Of Education — Special Education": "General Fund",\n'
        '    "Department Of Education — Elementary Education": "General Fund"\n'
        "  }\n"
        "}\n\n"
        "Additional rules:\n"
        "- fiscal_years: labels formatted YYYY-YY (e.g. 2026-27).\n"
        "- Entity names in Title Case.\n"
        "- Dollar amounts are plain floats, no $ or commas. "
        "If amounts are listed in thousands, multiply by 1000.\n"
        "- fund_sources: use the exact fund name from the bill. Default to 'General Fund'.\n"
        "- If no appropriations found, return "
        '{"fiscal_years": [], "entities": {}, "fund_sources": {}}.'
    )

    all_entities: dict = {}
    all_fiscal_years: list = []
    all_fund_sources: dict = {}

    print(f"  Extracting with GPT-4o — {len(chunks)} chunk(s)...")
    for idx, chunk in enumerate(chunks, 1):
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                max_tokens=8192,
                temperature=0,
                seed=42,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Extract appropriations (chunk {idx}/{len(chunks)}):\n\n{chunk}"},
                ],
            )
            raw = response.choices[0].message.content or ""
            data = json.loads(raw)

            for fy in data.get("fiscal_years", []):
                if fy not in all_fiscal_years:
                    all_fiscal_years.append(fy)

            for name, amounts in data.get("entities", {}).items():
                if not isinstance(amounts, dict):
                    continue
                if name not in all_entities:
                    all_entities[name] = {}
                for fy, amt in amounts.items():
                    try:
                        amt = float(amt)
                    except (TypeError, ValueError):
                        continue
                    # Keep the higher figure when the same entity appears in multiple chunks
                    if fy not in all_entities[name] or amt > all_entities[name][fy]:
                        all_entities[name][fy] = amt

            for name, source in data.get("fund_sources", {}).items():
                if name not in all_fund_sources and isinstance(source, str):
                    all_fund_sources[name] = source

        except Exception as exc:
            warnings.append(f"Chunk {idx}/{len(chunks)} error: {exc}")

    fiscal_years = all_fiscal_years or ['Total']
    primary_fy = fiscal_years[0]

    # Drop entities where every fiscal year amount is zero (unreadable amounts)
    ordered_entities = {
        name: amts for name, amts in all_entities.items()
        if any(v > 0 for v in amts.values())
    }

    grand_totals = {fy: sum(e.get(fy, 0) for e in ordered_entities.values()) for fy in fiscal_years}

    return {
        "entities": ordered_entities,
        "fiscal_years": fiscal_years,
        "grand_totals": grand_totals,
        "fund_sources": all_fund_sources,
        "bill_grand_totals": {},
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
const fundSources = data.fund_sources || {};

function fmt(n) {
  if (!n || n === 0) return '-';
  return '$' + Math.round(n).toLocaleString('en-US');
}

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

const nameWidth   = 3600;
const sourceWidth = 1680;
const fyWidth     = fys.length > 1 ? Math.floor(4080 / fys.length) : 4080;
const totalWidth  = nameWidth + sourceWidth + fyWidth * fys.length;

function headerCell(text, width, align) {
  return new TableCell({
    borders, margins: cm,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: '1F3864', type: ShadingType.CLEAR },
    children: [new Paragraph({
      alignment: align || AlignmentType.RIGHT,
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20 })]
    })]
  });
}

const headerRow = new TableRow({
  tableHeader: true,
  children: [
    headerCell('Agency / Cabinet / Department', nameWidth, AlignmentType.LEFT),
    headerCell('Fund Source', sourceWidth, AlignmentType.LEFT),
    ...fys.map(fy => headerCell(fy, fyWidth))
  ]
});

const EDU_KEYWORDS = /education|school|university|college|vocational|k-12|higher ed/i;

const dataRows = Object.entries(data.entities).map(([name, totals], i) => {
  const isEdu = EDU_KEYWORDS.test(name);
  const fill = isEdu ? 'FFF8E7' : (i % 2 === 0 ? 'F5F7FA' : 'FFFFFF');
  const source = fundSources[name] || '';
  return new TableRow({
    children: [
      new TableCell({
        borders, margins: cm,
        width: { size: nameWidth, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: name, size: 18 })] })]
      }),
      new TableCell({
        borders, margins: cm,
        width: { size: sourceWidth, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: source, size: 16, italics: true, color: '444444' })] })]
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
    new TableCell({
      borders, margins: cm,
      width: { size: sourceWidth, type: WidthType.DXA },
      shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: '' })] })]
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
        columnWidths: [nameWidth, sourceWidth, ...fys.map(() => fyWidth)],
        rows: [headerRow, ...dataRows, totalRow]
      }),
      ...warningParagraphs,
      new Paragraph({
        spacing: { before: 300 },
        children: [new TextRun({
          text: 'Methodology: Appropriations were extracted from the source PDF using AI (GPT-4o). ' +
                'Each named agency, department, cabinet, or bureau is listed with its all-funds appropriation ' +
                'for the identified fiscal year(s). Grand-total and rollup lines are excluded to prevent ' +
                'double-counting. Verify all figures against the enrolled bill before citing.',
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

    reports_dir = Path(__file__).parent / "Reports"
    reports_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = sys.argv[2] if len(sys.argv) > 2 else str(reports_dir / (Path(pdf_path).stem + f"_summary_{timestamp}.docx"))

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
    generate_report(data, pdf_path, out_path)
    print("Done.")


if __name__ == '__main__':
    main()
