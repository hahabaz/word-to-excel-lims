---
name: word-to-excel-lims
description: Convert Word original-record/LIMS templates (.docx) into stable, editable Excel 2019 .xlsx templates using a rules + example-case library + layout_plan.json + deterministic generation scripts + validation gates. Use when the user asks for Word to Excel conversion, Word原始记录模板转Excel, LIMS原始记录导出模板, Excel 2019兼容, 等宽基础列, 打印页数一致, or stable editable xlsx output.
license: MIT
compatibility: Agent Skills format. Python 3.10+ recommended. Runtime can operate with an empty examples/ case library.
metadata:
  version: "2.0.0"
  language: "zh-CN"
---

# Word → Excel LIMS Skill

## Purpose

This skill converts uploaded Word original-record/LIMS templates into editable, printable, Excel 2019-compatible `.xlsx` templates.

The skill is not a direct Word-format copier. It rebuilds the visible business structure manually in Excel using standard worksheet constructs: cells, fonts, borders, fills, row heights, equal-width base columns, merged ranges, page setup, print areas, and page breaks.

## Non-negotiable output goals

1. Output must be `.xlsx` and compatible with Excel 2019.
2. Do not use form controls, ActiveX controls, text boxes, drawing shapes, shape lines, images as tables, VML objects, or complex drawing XML.
3. Convert Word checkbox-like glyphs to plain text `□` unless visible meaning requires another plain character.
4. Rebuild every visible Word table line, separator, horizontal line, and vertical line with Excel cell borders.
5. Keep normal Excel gridlines visible. Do not hide sheet gridlines and do not white-fill the whole sheet to simulate a Word page.
6. Use equal-width base columns for the content area. Never solve layout by selectively widening individual columns.
7. Use merged cells over equal-width base columns for long labels, values, titles, instruction blocks, tables, and signature areas.
8. Fixed field labels should normally display on one line. Do not depend on forced line breaks to make labels fit.
9. Content must be horizontally centered on the print page and should visually occupy about 88%–95% of printable width.
10. Print preview page count should match the Word source page count whenever the environment can determine it. If the exact count cannot be determined, use best effort and state the limitation.
11. Prefer file stability over pixel-perfect appearance when Word formatting is unusually complex.
12. Run validation before delivery. If structural validation finds repair-risk issues, regenerate before returning the file.

## System architecture

This skill has four layers:

```text
1. Rules
   SKILL.md and references/*.md define conversion policy and quality gates.

2. Case library
   examples/case-*/ may contain source.docx, ai-output.xlsx, corrected.xlsx, and notes.md.
   The case library may be empty. Empty examples must not block conversion.

3. Layout JSON
   The model or scripts create layout_plan.json, a structured plan describing page setup,
   base column count, sections, merged ranges, row heights, borders, and print settings.

4. Deterministic scripts
   scripts/analyze_docx.py extracts visible Word structure.
   scripts/build_layout_plan.py builds a first-pass layout_plan.json.
   scripts/generate_xlsx_from_plan.py generates .xlsx from layout_plan.json.
   scripts/validate_xlsx_structure.py validates XLSX structure and repair-risk signals.
   scripts/run_conversion.py runs the pipeline end to end.
```

## When to use this skill

Use this skill when the user asks to:

- convert a Word original-record template to Excel;
- create an Excel 2019-compatible LIMS export template;
- preserve Word-visible fields, tables, lines, checkboxes, signature areas, and page structure;
- improve an existing Word-to-Excel conversion that has broken layout, right-side blank space, repair warnings, unstable merges, or print-page mismatch;
- use a reference Excel template while keeping the current Word content as the source of truth.

## Required input handling

### If the user provides only a Word file

1. Analyze the Word file's visible text, paragraphs, tables, fields, checkboxes, and page orientation hints.
2. Build a first-pass layout plan using this skill's default rules and scripts.
3. Generate the `.xlsx` file.
4. Validate the file.
5. Return the generated `.xlsx` plus a short QA summary.

### If the user provides a Word file and a reference Excel file

Use the Word file as the content source. Use the Excel file only as a style and layout reference for page proportions, font sizes, border style, margins, row/column density, and print setup. Do not copy unrelated fields from the reference Excel.

### If the user provides example cases

Search `examples/case-*` for relevant `notes.md` and `corrected.xlsx`. Prefer the closest case's corrected file as a layout-style reference, but do not copy unrelated business fields. Transfer only reusable layout principles.

### If the examples directory is empty

Continue normally. Use the built-in layout heuristics, `layout_plan.schema.json`, and references. Do not ask the user for examples unless the current Word file is too ambiguous to reconstruct.

## Conversion workflow

### Step 1 — Extract Word structure

Use `scripts/analyze_docx.py` where possible:

```bash
python scripts/analyze_docx.py input.docx --out word_analysis.json
```

The analysis should capture:

- visible paragraphs;
- table count, rows, columns, and cell text;
- checkbox-like symbols normalized to `□`;
- estimated complexity signals;
- page setup hints when available;
- illegal/control character cleanup needs.

If script extraction is incomplete, use direct document inspection and manual reasoning. Never copy hidden Word control characters into Excel.

### Step 2 — Consult case library, if any

Case folders follow this pattern:

```text
examples/case-001/
├── source.docx
├── ai-output.xlsx
├── corrected.xlsx
└── notes.md
```

`corrected.xlsx` is the target-quality reference. `ai-output.xlsx` is the known imperfect version. `notes.md` explains what changed and why.

If `examples/` contains no case folders, record `case_library.status = "empty"` in the layout plan and proceed.

### Step 3 — Produce layout_plan.json

Before generating Excel, create a layout plan. The plan is the bridge between model reasoning and deterministic workbook generation.

Use the schema in `references/layout_plan.schema.json`. The plan must include:

- schema version;
- source summary;
- selected case reference or empty-case status;
- page setup;
- base column count;
- equal column width;
- sections, rows, cells, merges, borders, row heights;
- print area and page-fit strategy;
- QA assumptions and limitations.

Use `scripts/build_layout_plan.py` for a runnable baseline:

```bash
python scripts/build_layout_plan.py word_analysis.json --examples examples --out layout_plan.json
```

A human or model may then refine `layout_plan.json` before generation.

### Step 4 — Generate XLSX deterministically

Use `scripts/generate_xlsx_from_plan.py`:

```bash
python scripts/generate_xlsx_from_plan.py layout_plan.json --out output.xlsx
```

Generation rules:

- use only worksheet cells and standard Excel styles;
- keep all base columns equal width;
- use merged ranges for layout spans;
- style borders as cell borders only;
- keep `showGridLines` enabled;
- avoid drawings, pictures, controls, and text boxes;
- set page setup, margins, print area, horizontal centering, and fit-to-page settings.

### Step 5 — Validate before delivery

Run:

```bash
python scripts/validate_xlsx_structure.py output.xlsx --json validation_report.json
```

The file fails the quality gate if validation reports:

- XML parse errors;
- drawing/control/media/VML parts;
- disabled gridlines;
- duplicate or overlapping merged ranges;
- malformed merge references;
- suspicious print settings;
- workbook or worksheet XML loading errors.

If the file fails, regenerate or simplify before returning.

## Layout strategy

### Equal-width base columns

Choose the base-column count from visible complexity:

| Template complexity | Suggested base columns |
|---|---:|
| Very simple fields, no dense table | 16–20 |
| Normal forms, several field rows | 24–32 |
| Multi-field rows or normal record tables | 32–40 |
| Dense tables, many narrow headers, multi-level headers | 40–56 |
| Highly dense horizontal records | 56–64 |

The exact count must not be fixed globally. Increase the count when right-side blank space is obvious, headers need narrower granularity, or labels cannot survive without forced line breaks.

### Merge strategy

- Title: merge across all base columns.
- Long instruction block: merge across all or most base columns.
- Short label: merge 2–4 base columns.
- Normal label: merge 4–6 base columns.
- Long label: merge 6–10 base columns.
- Value area: merge enough remaining columns to keep the total row width full.
- Dense table column: allocate proportional spans across the base grid.

### Page strategy

- Use A4 portrait unless Word analysis/reference indicates another page size or orientation.
- Set horizontal centering.
- Fit width to 1 page by default.
- Match Word page count if known.
- Do not solve page count by making content unreadably small.
- Keep print area tight to actual template content, not empty helper columns.

### Border strategy

- Redraw all visible Word lines using Excel cell borders.
- Apply outer borders to merged regions and continuous borders to table areas.
- Do not use drawing lines.
- Do not rely on underlines alone for Word separator lines.

## Quality gate checklist

Before final answer, verify:

- `.xlsx` opens structurally as a ZIP package;
- all XML files parse;
- no `xl/drawings/`, `xl/ctrlProps/`, `xl/activeX/`, `xl/media/`, or VML drawing parts exist;
- no duplicate/overlapping merged ranges;
- worksheet gridlines are not disabled;
- print area is tight;
- base columns in the content area are equal width;
- visible content is preserved;
- labels do not depend on forced line breaks;
- page is horizontally centered;
- right-side blank space is not obvious;
- border lines are continuous;
- file is editable and printable.

## How to improve the skill over time

When a conversion is unsatisfactory, add a new case:

```text
examples/case-###/source.docx       original Word file
examples/case-###/ai-output.xlsx    unsatisfactory AI result
examples/case-###/corrected.xlsx    human-corrected target result
examples/case-###/notes.md          what was wrong and how it was corrected
```

Then update `references/common-corrections.md` if the correction is general. This creates a reusable case library without requiring model fine-tuning.

## Final response format

Return:

1. the generated `.xlsx` file;
2. a brief QA summary: validation status, base column count, page setup, whether examples were used, and any limitations;
3. no long reproduction of the internal layout plan unless the user asks for it.
