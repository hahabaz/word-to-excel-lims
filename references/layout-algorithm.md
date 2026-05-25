# Equal-Width Base Column Layout Algorithm

This skill uses equal-width base columns plus merged cells. This is the main rule that makes Word-like templates editable, printable, and stable in Excel.

## 1. Estimate template complexity
Use visible Word structure signals:

| Signal | Effect |
|---|---|
| Few fields, no dense tables | 16–24 base columns |
| Several field rows, normal tables | 24–36 base columns |
| Many horizontal fields or multi-level headers | 36–48 base columns |
| Dense record tables, narrow headers, many side-by-side fields | 48–64 base columns |

Do not hard-code one column count for every document. Increase columns when right-side blank space appears or when field/header density cannot be represented cleanly.

## 2. Keep base columns equal
All participating columns must have the same width. Solve layout by merging columns, not by widening individual columns.

Examples:
- Short label: merge 2–4 base columns.
- Normal label: merge 4–6 base columns.
- Long label: merge 6–10 base columns.
- Value field: merge remaining columns in the row or section.
- Title: merge all content columns.
- Signature line: merge a label region and a writable region.

## 3. Reserve label width without forced wrapping
Fixed labels should normally survive if the user removes manual line breaks. If a label wraps only because its merged range is too narrow, widen the merged range by adding base columns.

## 4. Fill printable width
The content grid should occupy roughly 88%–95% of printable width. Use page setup centering and a tight print area. Do not include empty helper columns in the print area.

## 5. Border strategy
Apply borders to actual cell edges. For merged regions, style the outer edge of the entire merged rectangle and adjacent internal boundaries where the original Word table has visible lines.

Avoid all shape/drawing line objects.

## 6. Print strategy
- Match paper size and orientation.
- Center horizontally.
- Fit width to 1 page unless the Word template itself is wider than one page.
- Use manual page breaks to match Word page count.
- Prefer layout rebalance over extreme print scaling.
