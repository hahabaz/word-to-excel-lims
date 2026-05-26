# Word → Excel LIMS Conversion Checklist

Use this as the final gate before delivery.

## Compatibility

- [ ] Output is `.xlsx`.
- [ ] Compatible with Excel 2019.
- [ ] No form controls, ActiveX controls, text boxes, shape lines, drawing objects, media, or picture tables.
- [ ] Word checkboxes are represented as plain text `□`.
- [ ] Visible Word lines are rebuilt with Excel borders.

## Content preservation

- [ ] Visible titles, fields, placeholders, table headers, table body areas, instructions, page numbers, and signatures are retained.
- [ ] Hidden Word artifacts are not copied into Excel.
- [ ] Illegal XML/control characters are removed.

## Layout

- [ ] Content is rebuilt manually in Excel.
- [ ] Content area uses equal-width base columns, using the smallest clean grid rather than defaulting to 56/64 columns.
- [ ] Individual columns are not widened selectively.
- [ ] Long fields and labels use merged equal-width columns.
- [ ] Fixed labels do not depend on forced line breaks.
- [ ] Worksheet gridlines remain visible outside the template area.
- [ ] No large white-filled canvas is used.

## Page and print

- [ ] Word print-preview page count is identified or best-effort estimated; if Word is 1 page, Excel is verified/fit to 1 page.
- [ ] Excel print-preview page count matches the Word page count where environment permits.
- [ ] Paper size, orientation, margins, centering, print area, and page breaks are set.
- [ ] Print area tightly covers actual template content.
- [ ] Content is horizontally centered and visually occupies about 88%–95% of printable width.
- [ ] No obvious right-side blank area.

## Borders

- [ ] Every visible Word table border and separator line is represented by Excel borders.
- [ ] Borders are continuous and aligned.
- [ ] Borders start from the correct visual locations, not only after text.

## XLSX stability

- [ ] XLSX ZIP opens successfully.
- [ ] XML files parse successfully.
- [ ] No duplicate or overlapping merged ranges.
- [ ] No drawing/control/XML-risk parts are present.
- [ ] workbook.xml, sheet XML, styles.xml, mergeCells, and print settings are structurally valid.
- [ ] Worksheet gridlines are not disabled.

## Delivery

- [ ] Return the `.xlsx` file.
- [ ] Include a short QA summary.
- [ ] State any limitation honestly.


## A012 / seeded-case regression checks

Apply when the current document resembles `examples/case-001-cjjkjl-a012-pathogen`:

- [ ] Source `.doc`/`.docx` portrait orientation is preserved as Excel portrait.
- [ ] Word 1-page template does not become 2 Excel print pages.
- [ ] Page-top redundant organization/header text is not duplicated in the worksheet body when the corrected case omits it.
- [ ] 沙门氏菌、志贺氏菌、金黄色葡萄球菌、单核细胞增生李斯特菌 sections have no AI-added filler rows.
- [ ] Table headers and body cells have no fill color unless source/reference visibly uses fill.
- [ ] Base columns are visually compact; avoid 56/64-column grids for this portrait one-page style.
