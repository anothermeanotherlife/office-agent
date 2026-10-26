"""Build an Excel workbook with openpyxl.

Rules enforced here (see STANDARDS.md):
- Sheets are named README / Data / Summary / Dashboard, never Sheet1.
- Summary values are FORMULAS linked to Data, never hand-typed totals.
  openpyxl does not calculate formulas; values are verified after a
  LibreOffice recalc (see tools/render_check.py) and by an independent
  Python cross-check in audit.py.
- openpyxl does not create a native PivotTable from scratch, so Summary
  is labelled a summary table, not a PivotTable.
- Number formats follow the en-US Excel standard in design_tokens.json;
  locale display is Excel's job on open.
- Style objects are created once and reused (creating a new Font/Fill per
  cell bloats the file).
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .theme import hex_to_rgb, load_tokens


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def build_xlsx(data: dict, out_path: str | Path, tokens: dict | None = None) -> Path:
    tokens = tokens or load_tokens()
    colors = tokens["colors"]
    fonts = tokens["fonts"]
    fmts = tokens["excel"]["number_formats"]

    wb = Workbook()
    readme = wb.active
    readme.title = "README"
    readme["A1"] = data["title"]
    readme["A1"].font = Font(name=fonts["heading"], size=16, bold=True, color=colors["primary"])
    readme["A2"] = "Nguồn số liệu và định nghĩa chỉ số — xem Data và Summary."
    readme["A3"] = "Lưu ý: Summary là bảng tổng hợp bằng công thức, không phải PivotTable native."
    for i, source in enumerate(data.get("sources", []), start=5):
        readme.cell(row=i, column=1, value=source)
    readme.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Data")
    headers = data["table"]["headers"]
    rows = data["table"]["rows"]
    header_font = Font(name=fonts["body"], size=11, bold=True, color=colors["header_text"])
    header_fill = _fill(colors["header_bg"])
    thin = Side(style="thin", color=colors["muted"])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ci, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    banded = _fill(colors["banded_row"])
    for ri, row in enumerate(rows, start=2):
        for ci, value in enumerate(row, start=1):
            cell = ws.cell(row=ri, column=ci, value=value)
            cell.font = Font(name=fonts["body"], size=11)
            cell.border = border
            if ri % 2 == 0:
                cell.fill = banded
            if ci >= 2:  # numeric columns in the example data
                cell.number_format = fmts["int"]
    ws.add_table(Table(displayName="DataTable", ref=f"A1:{ws.cell(row=1, column=len(headers)).column_letter}{len(rows) + 1}"))
    ws.tables["DataTable"].tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_title_rows = "1:1"
    ws.print_area = ws.dimensions
    ws.column_dimensions["A"].width = 22
    for col in ("B", "C", "D"):
        ws.column_dimensions[col].width = 18

    # Data validation: numeric columns must be whole numbers >= 0.
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv.error = "Vui lòng nhập số nguyên >= 0."
    dv.prompt = "Số nguyên >= 0"
    ws.add_data_validation(dv)
    dv.add(f"B2:D{len(rows) + 1}")

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Chỉ số"
    summary["B1"] = "Tổng (công thức liên kết Data)"
    for cell in (summary["A1"], summary["B1"]):
        cell.font = header_font
        cell.fill = header_fill
    n = len(rows) + 1
    summary["A2"] = "Tổng doanh thu"
    summary["B2"] = f"=SUM(Data!B2:B{n})"
    summary["A3"] = "Tổng chi phí"
    summary["B3"] = f"=SUM(Data!C2:C{n})"
    summary["A4"] = "Lợi nhuận (doanh thu - chi phí)"
    summary["B4"] = "=B2-B3"
    for row in (2, 3, 4):
        summary.cell(row=row, column=2).number_format = fmts["int"]
        summary.cell(row=row, column=1).font = Font(name=fonts["body"], size=11)
    summary["A5"] = "Tỷ lệ lợi nhuận"
    summary["B5"] = "=IF(B2=0,0,B4/B2)"
    summary["B5"].number_format = fmts["pct"]
    summary.freeze_panes = "A2"
    summary.sheet_properties.pageSetUpPr.fitToPage = True
    summary.page_setup.fitToWidth = 1
    summary.page_setup.fitToHeight = 1
    summary.print_title_rows = "1:1"
    summary.column_dimensions["A"].width = 36
    summary.column_dimensions["B"].width = 28
    # Conditional formatting: highlight profit ratio below 10%.
    summary.conditional_formatting.add(
        "B5", CellIsRule(operator="lessThan", formula=["0.1"], fill=_fill("FECACA"))
    )

    dashboard = wb.create_sheet("Dashboard")
    dashboard["A1"] = "Biểu đồ doanh thu theo kênh (nguồn: Data)"
    chart = BarChart()
    chart.type = "col"
    chart.title = "Doanh thu theo kênh"
    chart.y_axis.title = "Doanh thu"
    chart.x_axis.title = "Kênh"
    chart.style = 10
    chart.shape = 4
    chart.legend = None
    chart_data = Reference(ws, min_col=2, min_row=1, max_row=n)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n)
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(cats)
    chart.dataLabels = None
    dashboard.add_chart(chart, "A3")
    dashboard.sheet_properties.pageSetUpPr.fitToPage = True

    wb.properties.title = data["title"]
    wb.properties.creator = data.get("author", tokens["brand"])

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return out
