---
name: office-excel
description: Build a professional Excel workbook with openpyxl, with formula-linked summaries, Excel Tables, charts and print setup.
---

# Office Excel

Use this skill when the task is a `.xlsx` workbook.

Workflow:

1. Read `design_tokens.json` and `STANDARDS.md`.
2. Sheets: README / Data / Summary / Dashboard. Never Sheet1.
3. Summary values are formulas linked to Data, never hand-typed totals.
4. openpyxl does NOT calculate formulas and does NOT create a native
   PivotTable from scratch. Label Summary honestly, and verify values
   after a LibreOffice recalc + an independent Python cross-check.
5. Numbers are numbers, dates are datetimes. Number formats follow the
   en-US Excel standard in the tokens; locale display is Excel's job.
6. Freeze panes, auto filter, print title rows, print area, fitToPage.
7. Reuse style objects; do not create a new Font/Fill per cell.
8. Run `office_builder/audit.py` on the real file before delivering.

See `references/xlsx.md`.
