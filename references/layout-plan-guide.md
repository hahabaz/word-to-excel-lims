# Layout Plan Guide

`layout_plan.json` is a stable intermediate format. It makes the conversion less dependent on free-form reasoning.

## Why use a layout plan?

Instead of asking the model to directly create a workbook, first create a structured plan:

1. what page setup to use;
2. how many equal-width base columns to use;
3. where sections begin;
4. what text appears in each merged region;
5. which borders and alignment are needed;
6. what print area should be set.

The generator script then converts the plan to `.xlsx` in a deterministic way.

## Minimal plan shape

```json
{
  "schema_version": "1.0",
  "source": {"docx_path": "input.docx", "summary": "..."},
  "case_library": {"status": "empty", "selected_case": null},
  "workbook": {
    "sheets": [
      {
        "name": "Sheet1",
        "page_setup": {"paper_size": "A4", "orientation": "portrait", "horizontal_centered": true},
        "base_grid": {"base_columns": 32, "equal_column_width": 2.4},
        "sections": []
      }
    ]
  }
}
```

## Manual refinement

After generating a first-pass plan, a model or human may edit:

- base column count;
- merged spans;
- row heights;
- borders;
- print settings;
- section ordering.

Then run the generator again.
