---
name: word-to-excel-lims
description: Convert Word original-record/LIMS templates (.docx) into stable, editable Excel 2019 .xlsx templates. Use when the user asks to convert Word to Excel, Word原始记录模板转Excel, LIMS原始记录导出模板, preserve visible Word content, redraw borders with Excel cell borders, keep print pages consistent, use equal-width base columns, avoid controls/drawings, and validate XLSX stability.
license: MIT
compatibility: Agent Skills format. Requires a runtime that can read uploaded DOCX files and create XLSX files. Python 3.10+ recommended; LibreOffice recommended for page-count/render QA.
metadata:
  version: "1.0.0"
  language: "zh-CN"
---

# Word → Excel LIMS Template Converter

## Goal
Convert an uploaded Word original-record template into a stable, editable, printable `.xlsx` template for Excel 2019 and LIMS export use.

Primary objective: reproduce the **visible business structure** of the Word file in Excel, not Word's hidden layout/control XML. Prioritize workbook stability, editability, print correctness, and page fidelity over pixel-perfect copying.

## When to use this skill
Use this skill when the user asks to:
- convert a Word/DOCX original record template to Excel/XLSX;
- make a LIMS raw-record export template;
- preserve Word-visible tables, fields, checkboxes, signatures, page numbers, borders, and print layout in Excel;
- avoid Excel repair warnings, damaged XML, broken merged cells, or Drawing/control objects.

Do not use this skill for data analysis spreadsheets, normal tabular reports, or conversions where the user wants the Word content pasted directly without template reconstruction.

## Core rules
1. Output must be `.xlsx` compatible with Excel 2019.
2. Do not use Excel form controls, ActiveX, text boxes, shapes, image-based tables, drawing lines, or nonstandard XML.
3. Replace Word checkbox controls with plain text `□`.
4. Recreate visible Word lines and table rules using Excel cell borders only.
5. Build the workbook as a manual Excel template reconstruction. Do not copy Word's hidden spaces, control characters, section artifacts, or raw layout XML.
6. Keep worksheet gridlines visible. Do not fill large blank areas white or hide normal Excel gridlines.
7. Use equal-width base columns plus merged cells. Do not solve width problems by making individual columns wider.
8. Fixed labels should normally display on one line by merging enough equal-width columns. Do not depend on forced line breaks to preserve the layout.
9. Excel print preview page count must match the Word print preview page count as closely as the environment allows.
10. If exact Word fidelity conflicts with XLSX stability, choose XLSX stability first and explain the tradeoff briefly.

## Fast execution workflow

### 1. Inspect inputs
- Identify the uploaded Word template and any optional reference Excel template.
- Determine paper size, orientation, margins, headers/footers, visible page count, titles, business sections, tables, signature areas, and page numbering.
- If LibreOffice or a document renderer is available, render DOCX/PDF/PNG to confirm the visible page count and layout.
- If only text/table extraction is available, proceed with best effort and state that visual page-count verification was limited.

Helpful script: `scripts/analyze_docx.py <input.docx> --out analysis.json`

### 2. Normalize visible content
- Keep all visible text, table headers, business fields, instructions, placeholders, signatures, and page numbers.
- Remove illegal XML characters and unstable control characters.
- Convert checkbox-like controls/symbols to `□` unless already visibly checked and the user asked to preserve checked state.
- Ignore hidden Word artifacts: hidden paragraph markers, excess tabs, repeated invisible spaces, non-rendered controls, and layout-only XML.

### 3. Choose an equal-width base-column grid
Choose base columns dynamically from template complexity:
- Simple form: 16–24 columns.
- Medium form with several field rows/tables: 24–36 columns.
- Dense record table or multi-level headers: 36–48 columns.
- Very dense horizontal layout: 48–64 columns.

All base columns in the content area must have the same width. Use merged ranges to represent labels, value fields, long text areas, table cells, and signature blocks. Set print area tightly around actual content, not around empty helper columns.

Target page fill: the content area should visually occupy about 88%–95% of printable width, centered horizontally, without obvious right-side blank space.

See `references/layout-algorithm.md` for the detailed grid allocation method.

### 4. Rebuild Excel layout
- Create one worksheet unless the Word template clearly requires multiple sheets.
- Use standard fonts, alignments, row heights, borders, margins, and print settings.
- Titles, section headers, long notes, and signature fields may span many base columns.
- Use cell borders for every visible Word table border, horizontal line, vertical line, and separator line.
- Avoid excessive styling. Use fills only inside real template content areas when visually necessary.
- Keep row/column gridlines enabled outside the template area.

### 5. Page setup and print control
- Match Word paper size, orientation, margins, page headers/footers, and page count where possible.
- Set horizontal centering on page.
- Set a tight print area around the actual template content.
- Use fit-to-page settings carefully: usually fit width to 1 page and control height/page breaks to match the Word page count.
- If Word has 1 print page, Excel should print as 1 page. If Word has 2 pages, Excel should print as 2 pages, etc.
- Do not add pages just because content is crowded; first optimize row heights, base columns, merged ranges, font size, and scaling.
- Do not compress so aggressively that the template becomes unreadable or visually narrow.

### 6. XLSX stability gate
Before returning the file, verify:
- XLSX ZIP opens and all XML parts parse.
- No `xl/drawings/`, `xl/ctrlProps/`, `xl/activeX/`, or VML controls are present unless explicitly unavoidable and safe.
- No duplicate, overlapping, crossed, or out-of-bounds merged ranges.
- No illegal XML characters in strings.
- No sheet gridline setting disables normal gridlines.
- Workbook, worksheet, styles, mergeCells, and print settings are structurally valid.
- Print area does not include large empty right-side space.

Helpful script: `scripts/validate_xlsx_structure.py <output.xlsx>`

### 7. Visual QA loop
Render or inspect the final workbook if the environment supports it.
Check:
- page count matches Word;
- content is horizontally centered;
- template width is visually full but not clipped;
- borders are continuous and not truncated;
- all visible text and business fields are present;
- labels do not rely on manual line breaks;
- no content is hidden, clipped, overlapped, or split onto unexpected pages.

If QA fails, patch the workbook and rerun the relevant checks. Limit iterations to focused fixes: column count/width, merge spans, print area, scaling, row heights, borders.

## Output format
Return:
1. A downloadable `.xlsx` file.
2. A short QA summary including:
   - Word page count source: renderer / document metadata / best effort.
   - Excel print page target.
   - Base-column count and equal-width rule used.
   - Major stability checks passed.
   - Any limitations or tradeoffs.

## Common failure fixes
- **Right blank area too large:** increase base-column count, widen all base columns uniformly, expand merged ranges, and reset print area tightly.
- **Labels wrap after removing manual breaks:** allocate more merged base columns to label cells; do not widen only one column.
- **Excel repair warning risk:** remove drawings, controls, VML, malformed styles, conflicting merges, and illegal XML characters.
- **Borders break:** redraw borders on every cell edge across merged regions; do not use shapes.
- **Too many pages:** reduce row height, adjust font size moderately, rebalance merged ranges, use fit-to-width, and insert page breaks intentionally.
- **Unreadably compressed:** increase page count only if Word page count also requires it; otherwise redesign the grid instead of shrinking excessively.

## Repository resources
- `references/conversion-checklist.md` — concise pass/fail checklist.
- `references/layout-algorithm.md` — base-column and merge-range strategy.
- `scripts/analyze_docx.py` — extracts DOCX visible structure signals.
- `scripts/validate_xlsx_structure.py` — validates XLSX ZIP/XML/merge/drawing-control risks.
