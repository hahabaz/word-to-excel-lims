# word-to-excel-lims

A reusable Agent Skill for converting Word original-record/LIMS templates into stable Excel 2019 `.xlsx` templates.

## What it does

This skill turns a long, fragile prompt into a repeatable workflow for agents:

- interprets visible Word layout and business structure;
- rebuilds the template manually in Excel;
- uses equal-width base columns and merged cells;
- redraws Word-visible lines with Excel cell borders;
- avoids Excel controls, shapes, text boxes, image tables, ActiveX, and risky XML;
- preserves worksheet gridlines;
- validates XLSX structure before delivery.

## Files

```text
word-to-excel-lims/
├── SKILL.md
├── README.md
├── examples/
│   └── user_prompt.md
├── references/
│   ├── conversion-checklist.md
│   └── layout-algorithm.md
└── scripts/
    ├── analyze_docx.py
    └── validate_xlsx_structure.py
```

## Install in ChatGPT Skills

1. Compress the `word-to-excel-lims` folder as a `.zip`.
2. In ChatGPT, open your profile menu.
3. Select **Skills**.
4. Select **New skill** → **Upload from your computer**.
5. Upload the zip file.
6. Test it with a Word template and ask for a Word-to-Excel LIMS conversion.

## Validate locally

```bash
python scripts/analyze_docx.py input.docx --out analysis.json
python scripts/validate_xlsx_structure.py output.xlsx --json
```

These scripts are helper checks. Final Excel quality still requires visual print-preview inspection when possible.
