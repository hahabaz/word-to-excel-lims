# Equal-Width Base Column Layout Algorithm

This skill uses equal-width base columns plus merged cells. This is the core technique for making Word-like templates editable, printable, and stable in Excel.

## 1. Estimate template complexity

| Signal | Effect |
|---|---|
| Few fields, no dense tables | 16–20 base columns |
| Several field rows, normal tables | 24–32 base columns |
| Many horizontal fields or multi-level headers | 32–48 base columns |
| Dense record tables, narrow headers, many side-by-side fields | 48–64 base columns |

Do not hard-code one column count for every document.

## 2. Keep base columns equal

All participating columns must have the same width. Solve layout by merging columns, not by widening individual columns.

## 3. Reserve label width without forced wrapping

Fixed labels should normally survive if the user removes manual line breaks. If a label wraps only because its merged range is too narrow, widen the merged range by adding base columns.

## 4. Fill printable width

The content grid should occupy roughly 88%–95% of printable width. Use page setup centering and a tight print area.

## 5. Border strategy

Apply borders to actual cell edges. Avoid all shape or drawing line objects.

## 6. Print strategy

- Match paper size and orientation.
- Center horizontally.
- Fit width to 1 page unless the Word template itself is wider than one page.
- Use page breaks to match Word page count when known.
- Prefer layout rebalance over extreme print scaling.
