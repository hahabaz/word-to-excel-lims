# word-to-excel-lims

A reusable Agent Skill for converting Word original-record/LIMS templates into stable, editable, printable Excel 2019-compatible `.xlsx` files.

This upgraded version uses a four-layer design:

1. **Rules** — `SKILL.md` and `references/*.md`
2. **Case library** — `examples/case-*` with Word, AI output, corrected Excel, and notes
3. **Layout JSON** — `layout_plan.json` as a structured layout contract
4. **Scripts** — deterministic generation and validation scripts

The case library can be empty. The skill still runs using default rules and layout heuristics.

## Quick start

```bash
cd word-to-excel-lims
python scripts/run_conversion.py path/to/input.docx --out output.xlsx
```

This will create intermediate files next to the output:

```text
output.word_analysis.json
output.layout_plan.json
output.validation_report.json
output.xlsx
```

## Optional manual workflow

```bash
python scripts/analyze_docx.py input.docx --out word_analysis.json
python scripts/build_layout_plan.py word_analysis.json --examples examples --out layout_plan.json
python scripts/generate_xlsx_from_plan.py layout_plan.json --out output.xlsx
python scripts/validate_xlsx_structure.py output.xlsx --json validation_report.json
```

## Dependencies

- Python 3.10+
- `openpyxl` for XLSX generation

Install:

```bash
pip install openpyxl
```

The analysis and validation scripts use Python standard library only.

## Add example cases over time

Each case should look like this:

```text
examples/case-001/
├── source.docx
├── ai-output.xlsx
├── corrected.xlsx
└── notes.md
```

`corrected.xlsx` is the target-quality reference. `notes.md` should explain what the AI result got wrong and how the corrected file fixed it.

## Empty examples are valid

At the beginning, `examples/` may contain no real `case-*` folders. The scripts will mark the case library as empty and continue.

## GitHub update flow

```bash
git add .
git commit -m "Upgrade word-to-excel-lims skill with layout JSON pipeline"
git push
```

## ChatGPT Skill upload

Zip the folder so that `SKILL.md` is inside the top-level `word-to-excel-lims/` folder, then upload the zip in ChatGPT Skills.
