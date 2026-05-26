# Common Corrections

Update this file when repeated human corrections appear across multiple cases.

## Right-side blank space

Common causes:

- too few base columns;
- print area includes wrong blank columns;
- title/table merges do not span the full content grid;
- content copied from Word at narrow width instead of rebuilt for Excel page width.

Preferred fix:

- increase base column count;
- keep columns equal width;
- expand merged ranges;
- set tight print area;
- enable horizontal centering.

## Labels depend on forced line breaks

Preferred fix:

- remove artificial label wrapping where possible;
- allocate more merged base columns to labels;
- keep value fields and total row width balanced.

## Broken borders

Preferred fix:

- use cell borders only;
- apply edges to merged ranges consistently;
- avoid drawing lines;
- apply continuous table grid borders.

## Excel repair warning risk

Preferred fix:

- remove drawing/control parts;
- avoid overlapping merges;
- validate XML after generation;
- regenerate with simpler standard styles if needed.


## A012 one-page portrait microbiology forms

Seeded from `examples/case-001-cjjkjl-a012-pathogen`.

Common causes:

- page-top organization/header text is treated as duplicate body text;
- source portrait orientation is changed to landscape or otherwise not preserved;
- Word source has 1 print page but Excel spills to 2 pages;
- extra blank rows are inserted below each detection-project section;
- table headers receive default gray fills even though the template should be no-fill;
- base-column count is too high, especially 56/64 columns on a portrait one-page form.

Preferred fix:

- use `corrected.xlsx` in the case folder as the closest visual reference for A012-like forms;
- keep the page portrait and fit the print area to exactly 1 page high when the Word source is 1 page;
- remove redundant page-header/organization text from the worksheet body when the corrected case omits it;
- keep all cells no-fill unless the source/reference visibly uses fill;
- reduce the base grid to the smallest clean equal-width grid, usually about 28–36 columns for this style;
- do not add filler rows after 沙门氏菌、志贺氏菌、金黄色葡萄球菌、单核细胞增生李斯特菌 sections.


## Hidden rows or columns in corrected examples

Common causes:

- user-made reference files may contain hidden helper rows from manual editing;
- copied templates may keep old rows hidden outside the visible print area;
- hidden rows can mislead automated layout extraction into adding extra blank rows.

Preferred fix:

- treat corrected examples as visible-print references, not raw row/column blueprints;
- record hidden rows/columns in notes or validation warnings;
- do not generate hidden rows or hidden columns by default;
- compare against the visible printed page and the source Word content.

## Multi-level table header misalignment

Common causes:

- distributing leaf columns evenly without checking parent group boundaries;
- placing subheaders on the wrong row level;
- using text presence as the only check and ignoring merged range start/end columns.

Preferred fix:

- build an explicit parent-child header map before merging cells;
- verify each parent header's merged range exactly covers its leaf columns;
- check vertical merges for standalone columns such as 系统生化、血清学实验、结果;
- inspect the rendered or print-preview layout before delivery.

## One-page output is too small

Common causes:

- forcing `fit_to_height=1` without rebalancing rows, columns, margins, and section spacing;
- using too many rows/columns or unnecessary blank rows;
- treating page count as the only success metric.

Preferred fix:

- keep Word 1 page as Excel 1 page whenever possible, but maintain readability;
- first tighten print area, remove extra blank rows, reduce only excessive row heights, and use compact merged cells;
- compare visual scale to the closest `corrected.xlsx` case;
- do not deliver a technically one-page file that is visibly much smaller than the source/reference.

## A013 Salmonella qualitative + quantitative forms

Seeded from `examples/case-002-cjjkjl-a013-salmonella-qual-quant`.

Common causes:

- 沙门氏菌定性表的多级表头错位；
- “增菌培养 / 分离培养 / 生化反应”父级合并范围与子列不匹配；
- 为了保持一页，把整体缩放得明显小于原文档和人工 Excel。

Preferred fix:

- use the A013 corrected workbook's visible layout as the closest style reference;
- explicitly map BPW/TTB/RVS to 增菌培养, XLD/沙显/BS to 分离培养, and the biochemical leaves to 生化反应;
- preserve no-fill styling and portrait orientation;
- keep one page only if the resulting scale remains readable and close to the reference.
