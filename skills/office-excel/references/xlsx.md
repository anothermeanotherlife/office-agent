# XLSX reference (openpyxl)

- openpyxl reads/edits; XlsxWriter only writes (faster, better charts)
  and cannot edit. pandas prepares data only; it must not be used to
  format, or it overwrites formatting.
- Styles are immutable: create Font/Fill/Border once and reuse them.
- Formulas are written, not calculated. Verify values after a
  LibreOffice recalc + an independent Python cross-check.
- openpyxl cannot create a native PivotTable from scratch (it can keep
  or edit one in a template). Summary tables must be labelled honestly.
- Charts: BarChart/LineChart/PieChart, `Reference`, `add_data`,
  `set_categories`, title, axis labels, no empty ranges.
- Print: orientation, paperSize, fitToWidth/fitToHeight, print_title_rows,
  print_area, margins, oddHeader/oddFooter.
- Freeze: `freeze_panes='A2'`; filter: `auto_filter.ref`. In write-only
  mode set them before appending data.
- Merged cells: only the top-left cell keeps a value; others become
  MergedCell. Do not merge a data area.
- `copy_worksheet` does not copy everything: copy data_validations,
  page_setup, margins, print titles/area and conditional formatting by hand.
