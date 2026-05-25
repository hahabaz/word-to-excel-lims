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
- [ ] Content area uses equal-width base columns.
- [ ] Individual columns are not widened selectively.
- [ ] Long fields and labels use merged equal-width columns.
- [ ] Fixed labels do not depend on forced line breaks.
- [ ] Worksheet gridlines remain visible outside the template area.
- [ ] No large white-filled canvas is used.

## Page and print

- [ ] Word print-preview page count is identified or best-effort estimated.
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
