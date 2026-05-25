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
