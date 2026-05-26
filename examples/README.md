# Case library

This folder may be empty at the beginning. Empty case library is valid and must not block conversion.

Add real cases as you accumulate examples:

```text
examples/case-001/
├── source.doc or source.docx
├── ai-output.xlsx
├── corrected.xlsx
└── notes.md
```

## File roles

- `source.doc` / `source.docx`: original Word template.
- `ai-output.xlsx`: earlier AI-generated result that was not good enough.
- `corrected.xlsx`: human-corrected Excel target. This is the best style reference.
- `notes.md`: what was wrong, how it was corrected, and what rule should be reused.

Case folders starting with `_` are templates or documentation and are ignored by scripts.


## Current included case

- `case-001-cjjkjl-a012-pathogen`: A012 食品微生物检验原始记录 case. Includes the original `.doc`, the earlier AI output, the user-made corrected Excel, and notes that capture the six regression rules from the user feedback.
