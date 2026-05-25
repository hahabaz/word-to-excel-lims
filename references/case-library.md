# Case Library Guide

The case library turns one-off corrections into reusable conversion knowledge.

## Recommended case structure

```text
examples/case-001/
├── source.docx
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

- base column count range;
- merge strategy;
- page width usage;
- table border style;
- font sizes;
- margins and print settings;
- signature area layout;
- how to avoid previous AI mistakes.

Do not blindly copy unrelated business fields from corrected examples.

## Good notes.md content

A useful `notes.md` should answer:

1. What was wrong in `ai-output.xlsx`?
2. What did the human change in `corrected.xlsx`?
3. What general rule should be reused?
4. What is specific to this document and should not be generalized?
