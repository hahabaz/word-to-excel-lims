# Case library

This folder may be empty at the beginning. Empty case library is valid and must not block conversion.

Add real cases as you accumulate examples:

```text
examples/case-001/
├── source.docx
├── ai-output.xlsx
├── corrected.xlsx
└── notes.md
```

## File roles

- `source.docx`: original Word template.
- `ai-output.xlsx`: earlier AI-generated result that was not good enough.
- `corrected.xlsx`: human-corrected Excel target. This is the best style reference.
- `notes.md`: what was wrong, how it was corrected, and what rule should be reused.

Case folders starting with `_` are templates or documentation and are ignored by scripts.
