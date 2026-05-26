# Equal-Width Base Column Layout Algorithm

This skill uses equal-width base columns plus merged cells. This is the core technique for making Word-like templates editable, printable, and stable in Excel.

## 1. Estimate template complexity

| Signal | Effect |
|---|---|
| Few fields, no dense tables | 16–20 base columns |
| Several field rows, normal tables | 24–32 base columns |
| Many horizontal fields or multi-level headers | 32–48 base columns |
| Dense record tables, narrow headers, many side-by-side fields | 40–56 base columns; use 56+ only for truly extreme layouts |

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


## A012-style compact exception

For one-page portrait microbiology original-record templates similar to `case-001-cjjkjl-a012-pathogen`, do **not** automatically escalate to 56/64 base columns just because some table headers are dense. Prefer 28–36 equal-width base columns, with merged cells for section titles and proportional spans for subcolumns. If the first-pass layout spills to a second print page, reduce blank rows and row heights before increasing column count or changing orientation.

## No-fill default

The default visual style is no cell fill. Header hierarchy should be expressed with bold text, alignment, row height, and borders. Use fill colors only when the Word source or the selected `corrected.xlsx` case visibly uses them.


## Hidden-row-safe case learning

When using a `corrected.xlsx` as reference, the visible print result is authoritative, not hidden helper rows or hidden columns. If hidden rows/columns exist, record them as warnings and ignore them when estimating target row count, section spacing, and data-entry row counts.

## Multi-level header alignment algorithm

For dense tables with nested headers:

1. List every leaf column in order.
2. Assign each leaf column to exactly one parent group.
3. Calculate parent merged ranges from the first to last leaf assigned to that parent.
4. Only then place visible header text and rowspans.
5. Validate that standalone columns with vertical merges do not steal columns from adjacent parent groups.

A rendered table where all text appears but parent/child boundaries are shifted is a failed conversion.

## Readable one-page scaling

For one-page Word sources, keep one-page Excel output as the goal. However, if the result becomes visibly much smaller than the Word source or the closest corrected case, rebalance the layout before delivery: remove unnecessary blank rows, reduce excessive row heights, tighten margins, adjust merged spans, and choose a compact base grid. Do not rely on extreme print scaling as the main solution.
