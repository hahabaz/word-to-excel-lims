# Case Library Guide

The case library turns one-off corrections into reusable conversion knowledge.

## Recommended case structure

```text
examples/case-001/
├── source.doc or source.docx
├── ai-output.xlsx
├── corrected.xlsx
└── notes.md
```

## Empty case library

An empty `examples/` folder is allowed. The pipeline must still run and should set:

```json
{"case_library": {"status": "empty", "selected_case": null}}
```

## What the model should learn from a case

Use the corrected Excel to infer:

- base column count range, including when to reduce overly dense 56/64-column grids;
- merge strategy;
- page width usage;
- table border style;
- font sizes;
- margins, orientation, print settings, and page-count constraints;
- signature area layout;
- how to avoid previous AI mistakes.

Do not blindly copy unrelated business fields from corrected examples.

## Good notes.md content

A useful `notes.md` should answer:

1. What was wrong in `ai-output.xlsx`?
2. What did the human change in `corrected.xlsx`?
3. What general rule should be reused?
4. What is specific to this document and should not be generalized?


## Seeded cases

### case-001-cjjkjl-a012-pathogen

Use this case when converting A012-like microbiology/pathogenic-bacteria original-record forms. The human `corrected.xlsx` is not absolute authority, but it is the preferred style reference for this class. Reuse these lessons:

- keep portrait source documents portrait;
- if Word is one page, Excel must print as one page;
- do not apply default fill colors;
- do not add extra blank rows under each detection project;
- avoid redundant page-top organization text in the worksheet body if the corrected case omits it;
- prefer a compact equal-width column grid over a visually noisy 56/64-column grid.
