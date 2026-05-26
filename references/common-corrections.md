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
