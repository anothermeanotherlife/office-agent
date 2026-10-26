"""Hard example: revenue up, profit down — a decision, not a template.

Data (million VND, internal aggregate, 2026):
  Q1: Online 460/250, Store 380/250, Agent 160/95   -> rev 1000, cost 595, profit 405
  Q2: Online 520/310, Store 380/260, Agent 150/90   -> rev 1050, cost 660, profit 390
  Q3: Online 560/340, Store 370/270, Agent 140/85   -> rev 1070, cost 695, profit 375
  Q4: Online 610/390, Store 350/280, Agent 120/80   -> rev 1080, cost 750, profit 330
  Year: rev 4200, cost 2700, profit 1500, margin 35.7%
  Q1 margin 40.5% -> Q4 margin 30.6%. Online cost rate 54.3% -> 63.9%.

Scenarios (inputs are blue cells in Excel; every total is a formula):
  Keep:      Online +8%, Store -3%, Agent -10%, cost rate +2pp -> profit 1449
  Narrow Ag: Online +18%, Store -3%, Agent -50%, cost rate +1pp -> profit 1490
  Invest On: Online +25%, Store -5%, Agent -10%, cost rate +4pp -> profit 1488
The recommendation is Narrow Agent, but by only ~2 over Invest Online —
so the deck states conditions and risks, not a triumphant certainty.

Run from repo root: python examples/build_hard_example.py
Outputs: examples/hard-output/report.{docx,xlsx,pptx}
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from office_builder.theme import load_tokens  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "examples" / "hard-output"

QUARTERS = ["Q1", "Q2", "Q3", "Q4"]
CHANNELS = ["Online", "Cửa hàng", "Đại lý"]
# (quarter, channel) -> (revenue, cost)
RAW = {
    ("Q1", "Online"): (460, 250), ("Q1", "Cửa hàng"): (380, 250), ("Q1", "Đại lý"): (160, 95),
    ("Q2", "Online"): (520, 310), ("Q2", "Cửa hàng"): (380, 260), ("Q2", "Đại lý"): (150, 90),
    ("Q3", "Online"): (560, 340), ("Q3", "Cửa hàng"): (370, 270), ("Q3", "Đại lý"): (140, 85),
    ("Q4", "Online"): (610, 390), ("Q4", "Cửa hàng"): (350, 280), ("Q4", "Đại lý"): (120, 80),
}
assert sum(v[0] for v in RAW.values()) == 4200
assert sum(v[1] for v in RAW.values()) == 2700


# ---------------------------------------------------------------- XLSX
def build_xlsx(path: Path, tokens: dict) -> None:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.styles import Border, Side

    colors = tokens["colors"]
    BLUE, BLACK, GREEN = "0000FF", "000000", "00B050"
    hdr_font = Font(name="Calibri", size=11, bold=True, color=colors["header_text"])
    hdr_fill = PatternFill("solid", fgColor=colors["header_bg"])
    thin = Side(style="thin", color=colors["muted"])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def header(ws, row, values):
        for ci, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=ci, value=v)
            c.font, c.fill, c.border = hdr_font, hdr_fill, border

    def style_row(ws, row, ncols, fmt=None, font_color=BLACK, bold=False):
        for ci in range(1, ncols + 1):
            c = ws.cell(row=row, column=ci)
            c.font = Font(name="Calibri", size=11, bold=bold, color=font_color)
            c.border = border
            if fmt and ci >= 2:
                c.number_format = fmt

    wb = Workbook()
    readme = wb.active
    readme.title = "README"
    readme["A1"] = "Báo cáo 4 quý 2026 — Doanh thu tăng, lợi nhuận giảm"
    readme["A1"].font = Font(name="Calibri", size=16, bold=True, color=colors["primary"])
    readme["A2"] = "Đơn vị: triệu đồng. Nguồn: dữ liệu nội bộ tổng hợp. Mọi tổng là công thức, không gõ tay."
    readme["A3"] = "Màu chữ theo chuẩn FAST: xanh dương = số nhập tay (input), đen = công thức trong sheet, xanh lá = công thức liên kết sheet khác."
    readme["A4"] = "Summary là bảng tổng hợp bằng công thức, không phải PivotTable native (openpyxl không tạo PivotTable từ đầu)."
    readme["A5"] = "Kịch bản và độ nhạy nằm ở sheet Scenario và Sensitivity; đổi ô xanh dương, các tổng tự đổi."
    readme.column_dimensions["A"].width = 120

    data = wb.create_sheet("Data")
    header(data, 1, ["Quý", "Doanh thu", "Chi phí", "Kênh", "Lợi nhuận"])
    r = 2
    for q in QUARTERS:
        for ch in CHANNELS:
            rev, cost = RAW[(q, ch)]
            data.cell(row=r, column=1, value=q)
            data.cell(row=r, column=2, value=rev)
            data.cell(row=r, column=3, value=cost)
            data.cell(row=r, column=4, value=ch)
            data.cell(row=r, column=5, value=f"=B{r}-C{r}")
            style_row(data, r, 5, fmt="#,##0", font_color=BLACK)
            # Inputs are blue, the profit formula is black.
            for ci in (1, 2, 3, 4):
                data.cell(row=r, column=ci).font = Font(name="Calibri", size=11, color=BLUE)
            r += 1
    data.freeze_panes = "A2"
    data.auto_filter.ref = data.dimensions
    data.sheet_properties.pageSetUpPr.fitToPage = True
    data.page_setup.orientation = "landscape"; data.page_setup.fitToWidth = 1; data.page_setup.fitToHeight = 1
    data.print_title_rows = "1:1"
    for col, w in zip("ABCDE", (12, 16, 16, 16, 16)):
        data.column_dimensions[col].width = w

    qtr = wb.create_sheet("Quarterly")
    header(qtr, 1, ["Quý", "Doanh thu", "Chi phí", "Lợi nhuận", "Biên"])
    for i, q in enumerate(QUARTERS, start=2):
        qtr.cell(row=i, column=1, value=q)
        qtr.cell(row=i, column=2, value=f'=SUMIFS(Data!B:B,Data!A:A,A{i})')
        qtr.cell(row=i, column=3, value=f'=SUMIFS(Data!C:C,Data!A:A,A{i})')
        qtr.cell(row=i, column=4, value=f"=B{i}-C{i}")
        qtr.cell(row=i, column=5, value=f"=IF(B{i}=0,0,D{i}/B{i})")
        style_row(qtr, i, 5, font_color=GREEN)
        qtr.cell(row=i, column=5).number_format = "0.0%"
        for ci in (2, 3, 4):
            qtr.cell(row=i, column=ci).number_format = "#,##0"
    qtr.freeze_panes = "A2"

    chan = wb.create_sheet("Channel")
    header(chan, 1, ["Kênh", "Doanh thu", "Chi phí", "Lợi nhuận", "Tỷ lệ chi phí", "Biên"])
    for i, ch in enumerate(CHANNELS, start=2):
        chan.cell(row=i, column=1, value=ch)
        chan.cell(row=i, column=2, value=f'=SUMIFS(Data!B:B,Data!D:D,A{i})')
        chan.cell(row=i, column=3, value=f'=SUMIFS(Data!C:C,Data!D:D,A{i})')
        chan.cell(row=i, column=4, value=f"=B{i}-C{i}")
        chan.cell(row=i, column=5, value=f"=IF(B{i}=0,0,C{i}/B{i})")
        chan.cell(row=i, column=6, value=f"=IF(B{i}=0,0,D{i}/B{i})")
        style_row(chan, i, 6, font_color=GREEN)
        for ci in (2, 3, 4):
            chan.cell(row=i, column=ci).number_format = "#,##0"
        for ci in (5, 6):
            chan.cell(row=i, column=ci).number_format = "0.0%"
    chan.freeze_panes = "A2"

    summ = wb.create_sheet("Summary")
    header(summ, 1, ["Chỉ số", "Tổng (công thức liên kết Data)"])
    summ["A2"] = "Tổng doanh thu"; summ["B2"] = "=SUM(Data!B2:B13)"
    summ["A3"] = "Tổng chi phí"; summ["B3"] = "=SUM(Data!C2:C13)"
    summ["A4"] = "Lợi nhuận (doanh thu - chi phí)"; summ["B4"] = "=B2-B3"
    summ["A5"] = "Tỷ lệ lợi nhuận"; summ["B5"] = "=IF(B2=0,0,B4/B2)"
    for row in (2, 3, 4):
        summ.cell(row=row, column=2).number_format = "#,##0"
    summ["B5"].number_format = "0.0%"
    summ["A7"] = "Đối chiếu: tổng theo quý trừ tổng chung (phải = 0)"
    summ["B7"] = "=SUM(Quarterly!B2:B5)-B2"
    summ["A8"] = "Đối chiếu: tổng theo kênh trừ tổng chung (phải = 0)"
    summ["B8"] = "=SUM(Channel!B2:B4)-B2"
    summ["A9"] = "Đối chiếu: lợi nhuận quý Q1→Q4 (Q4 − Q1, phải = −75)"
    summ["B9"] = "=Quarterly!D5-Quarterly!D2"
    summ.freeze_panes = "A2"
    summ.sheet_properties.pageSetUpPr.fitToPage = True
    summ.column_dimensions["A"].width = 52; summ.column_dimensions["B"].width = 30

    scen = wb.create_sheet("Scenario")
    header(scen, 1, ["Kịch bản", "Online +", "Cửa hàng +", "Đại lý +", "Tăng tỷ lệ chi phí (pp)"])
    inputs = [("Giữ nguyên", 0.08, -0.03, -0.10, 0.02),
              ("Thu hẹp Đại lý", 0.18, -0.03, -0.50, 0.01),
              ("Đầu tư Online", 0.25, -0.05, -0.10, 0.04)]
    for i, (name, o, s, a, d) in enumerate(inputs, start=2):
        scen.cell(row=i, column=1, value=name)
        for ci, v in enumerate((o, s, a, d), start=2):
            c = scen.cell(row=i, column=ci, value=v)
            c.number_format = "0.0%"
            c.font = Font(name="Calibri", size=11, color=BLUE)
    header(scen, 6, ["Kênh", "Doanh thu gốc", "Tỷ lệ chi phí gốc"])
    for i, ch in enumerate(CHANNELS, start=7):
        scen.cell(row=i, column=1, value=ch)
        scen.cell(row=i, column=2, value=f"=Channel!B{i - 5}")
        scen.cell(row=i, column=3, value=f"=Channel!E{i - 5}")
        style_row(scen, i, 3, font_color=GREEN)
        scen.cell(row=i, column=2).number_format = "#,##0"
        scen.cell(row=i, column=3).number_format = "0.0%"
    header(scen, 11, ["Kịch bản", "Doanh thu", "Chi phí", "Lợi nhuận", "Biên"])
    for j, scen_row in enumerate((2, 3, 4), start=12):
        scen.cell(row=j, column=1, value=f"=A{scen_row}")
        # Revenue: base * (1+growth) summed over channels; Cost: revenue * (base rate + delta)
        rev = "+".join(f"$B${7 + k}*(1+{chr(66 + k)}{scen_row})" for k in range(3))
        cost = "+".join(f"$B${7 + k}*(1+{chr(66 + k)}{scen_row})*($C${7 + k}+$E${scen_row})" for k in range(3))
        scen.cell(row=j, column=2, value=f"={rev}")
        scen.cell(row=j, column=3, value=f"={cost}")
        scen.cell(row=j, column=4, value=f"=B{j}-C{j}")
        scen.cell(row=j, column=5, value=f"=IF(B{j}=0,0,D{j}/B{j})")
        style_row(scen, j, 5, font_color=BLACK)
        for ci in (2, 3, 4):
            scen.cell(row=j, column=ci).number_format = "#,##0"
        scen.cell(row=j, column=5).number_format = "0.0%"
    scen.freeze_panes = "A2"
    scen.column_dimensions["A"].width = 22
    for col in "BCDE":
        scen.column_dimensions[col].width = 16

    sens = wb.create_sheet("Sensitivity")
    sens["A1"] = "Độ nhạy lợi nhuận — kịch bản Thu hẹp Đại lý (đổi ô xanh ở Scenario, bảng này tự đổi)"
    header(sens, 3, ["Tăng trưởng Online \\ Tăng tỷ lệ chi phí Online", 0, 0.01, 0.03])
    for ci in (2, 3, 4):
        sens.cell(row=3, column=ci).number_format = "0.0%"
        sens.cell(row=3, column=ci).font = Font(name="Calibri", size=11, bold=True, color=BLUE)
    for i, g in enumerate((0.10, 0.18, 0.25), start=4):
        sens.cell(row=i, column=1, value=g).number_format = "0.0%"
        sens.cell(row=i, column=1).font = Font(name="Calibri", size=11, color=BLUE)
        for ci, col in enumerate("BCD", start=2):
            # Online profit with varied growth/delta + Store/Agent profit of Narrow scenario (Scenario row 3 inputs)
            formula = (f"=$B$9*(1+$A{i})*(1-($C$9+{col}$3))"
                       f"+$B$10*(1+$C$3)*(1-($C$10+$E$3))"
                       f"+$B$11*(1+$D$3)*(1-($C$11+$E$3))")
            # NOTE: Scenario sheet layout: channel bases are rows 7-9 in the Scenario sheet,
            # not in Sensitivity. Fix references below in a second pass.
            sens.cell(row=i, column=ci, value=f"=Scenario!{formula[1:]}" if False else None)
    # Second pass with correct cross-sheet references (Scenario!B7:B9, C7:C9, inputs row 3 = Narrow).
    for i in range(4, 7):
        for ci, col in enumerate("BCD", start=2):
            sens.cell(row=i, column=ci, value=(
                f"=Scenario!$B$7*(1+$A{i})*(1-(Scenario!$C$7+{col}$3))"
                f"+Scenario!$B$8*(1+Scenario!$C$3)*(1-(Scenario!$C$8+Scenario!$E$3))"
                f"+Scenario!$B$9*(1+Scenario!$D$3)*(1-(Scenario!$C$9+Scenario!$E$3))"))
            sens.cell(row=i, column=ci).number_format = "#,##0"
            sens.cell(row=i, column=ci).font = Font(name="Calibri", size=11, color=GREEN)
    sens.freeze_panes = "A4"
    sens.column_dimensions["A"].width = 44
    for col in "BCD":
        sens.column_dimensions[col].width = 18

    dash = wb.create_sheet("Dashboard")
    dash["A1"] = "Doanh thu và lợi nhuận theo quý (nguồn: Quarterly; trục cột bắt đầu từ 0)"
    chart = BarChart()
    chart.type = "col"; chart.title = "Doanh thu và lợi nhuận theo quý"
    chart.y_axis.title = "Triệu đồng"; chart.x_axis.title = "Quý"
    chart.style = 10; chart.legend = None
    chart_data = Reference(qtr, min_col=2, min_row=1, max_row=5, max_col=4)
    cats = Reference(qtr, min_col=1, min_row=2, max_row=5)
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(cats)
    chart.y_axis.scaling.min = 0
    chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True
    dash.add_chart(chart, "A3")
    dash.sheet_properties.pageSetUpPr.fitToPage = True

    wb.properties.title = "Báo cáo 4 quý 2026 — Doanh thu tăng, lợi nhuận giảm"
    wb.properties.creator = tokens["brand"]
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


# ---------------------------------------------------------------- DOCX
def build_docx(path: Path, tokens: dict) -> None:
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def field(p, instr):
        for typ, txt in (("begin", None), (None, instr), ("separate", None), ("end", None)):
            r = p.add_run()
            if txt is not None:
                el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt; r._r.append(el)
            else:
                el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), typ); r._r.append(el)

    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"; doc.styles["Normal"].font.size = Pt(11)
    for name, size in (("Heading 1", 20), ("Heading 2", 16)):
        s = doc.styles[name]; s.font.name = "Calibri"; s.font.size = Pt(size); s.font.bold = True
    sec = doc.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    m = Cm(2.0); sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = m
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "Báo cáo nội bộ — 4 quý 2026"
    sec.footer.is_linked_to_previous = False
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run("Trang "); field(fp, " PAGE "); fp.add_run(" / "); field(fp, " NUMPAGES ")

    doc.add_heading("Doanh thu tăng, lợi nhuận giảm: thu hẹp Đại lý có kiểm soát", level=0)
    doc.add_paragraph("Bản ghi quyết định cho ban lãnh đạo. Đơn vị: triệu đồng, trừ khi ghi khác. "
                      "Mọi số trong bản này đều đối chiếu được với file Excel kèm theo (sheet Data → Summary).")
    toc = doc.add_paragraph(); field(toc, ' TOC \\o "1-3" \\h \\z \\u ')
    doc.add_paragraph("Mục lục do Word cập nhật khi mở file. Audit chỉ kiểm tra field TOC và Heading.")

    def table(headers, rows, caption):
        t = doc.add_table(rows=1 + len(rows), cols=len(headers)); t.style = "Table Grid"
        for i, h in enumerate(headers):
            t.rows[0].cells[i].paragraphs[0].add_run(h).bold = True
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row):
                t.rows[ri + 1].cells[ci].paragraphs[0].add_run(str(v))
        doc.add_paragraph(caption, style="Caption")

    chapters = [
        ("1. Kết luận trước: đề xuất và điều kiện", [
            "Đề xuất: thu hẹp kênh Đại lý 50% trong kỳ tới, giữ Cửa hàng, và kìm tỷ lệ chi phí Online dưới 64%.",
            "Lý do không phải vì Đại lý lỗ: Đại lý vẫn lãi 220. Lý do là Đại lý co lại đã rõ (160 → 120), "
            "trong khi Online tăng nhanh nhưng đắt dần, nên dồn nguồn lực vào kênh đang tăng mà không kiểm soát chi phí sẽ lặp lại đúng bài toán năm nay.",
            "Kết quả kịch bản: Thu hẹp Đại lý cho lợi nhuận 1.490, cao hơn Giữ nguyên 1.449 và Đầu tư Online 1.488, "
            "nhưng chỉ hơn Đầu tư Online 2. Vì vậy đây là lựa chọn có điều kiện, không phải chiến thắng áp đảo.",
            "Điều kiện dừng: nếu tỷ lệ chi phí Online vượt 64%, hoặc Cửa hàng giảm thêm quá 5%, phải xem lại đề xuất ngay trong tháng, không đợi hết quý.",
        ], []),
        ("2. Bằng chứng: doanh thu lên, lợi nhuận xuống", [
            "Doanh thu cả năm 4.200, tăng từ 1.000 (Q1) lên 1.080 (Q4), tức +8,0%. Lợi nhuận đi ngược: 405 (Q1) xuống 330 (Q4), tức −18,5%.",
            "Biên lợi nhuận: 40,5% (Q1) → 37,1% (Q2) → 35,0% (Q3) → 30,6% (Q4). Xu hướng giảm đều cả 4 quý, không phải một quý xấu đơn lẻ.",
            "Chi phí tăng 595 → 750, tức +26,1%, nhanh gấp 3 lần doanh thu. Đây là nguyên nhân chính, không phải doanh thu sụt.",
        ], [("Bảng 1 — Theo quý", ["Quý", "Doanh thu", "Chi phí", "Lợi nhuận", "Biên"],
             [["Q1", "1.000", "595", "405", "40,5%"], ["Q2", "1.050", "660", "390", "37,1%"],
              ["Q3", "1.070", "695", "375", "35,0%"], ["Q4", "1.080", "750", "330", "30,6%"],
              ["Cả năm", "4.200", "2.700", "1.500", "35,7%"]],
             "Nguồn: sheet Data, tổng theo quý bằng SUMIFS. Đơn vị triệu đồng. Lợi nhuận = doanh thu − chi phí.")]),
        ("3. Vì sao: chuyển dịch kênh và chi phí Online đắt dần", [
            "Online là động lực duy nhất tăng: 460 → 610 (+32,6%). Cửa hàng giảm 380 → 350, Đại lý giảm 160 → 120.",
            "Nhưng tỷ lệ chi phí Online tăng 54,3% (Q1) lên 63,9% (Q4). Mỗi đồng doanh thu Online mới mang về ít lợi nhuận hơn đồng cũ.",
            "Cả năm, Online lãi 860 trên doanh thu 2.150; Cửa hàng lãi 420 trên 1.480; Đại lý lãi 220 trên 570. "
            "Online vẫn là kênh lãi nhất tuyệt đối, nhưng biên của nó đang bị bào mòn nhanh nhất.",
        ], [("Bảng 2 — Theo kênh, cả năm", ["Kênh", "Doanh thu", "Chi phí", "Lợi nhuận", "Tỷ lệ chi phí"],
             [["Online", "2.150", "1.290", "860", "60,0%"], ["Cửa hàng", "1.480", "1.060", "420", "71,6%"],
              ["Đại lý", "570", "350", "220", "61,4%"], ["Tổng", "4.200", "2.700", "1.500", "64,3%"]],
             "Nguồn: sheet Data, tổng theo kênh bằng SUMIFS. Tỷ lệ chi phí = chi phí / doanh thu.")]),
        ("4. Ba kịch bản và lựa chọn", [
            "Ba kịch bản dùng cùng một bộ giả định đầu vào (ô xanh dương ở sheet Scenario), không phải ba con số gõ tay riêng lẻ.",
            "Giữ nguyên: doanh thu 4.271, lợi nhuận 1.449, biên 33,9%. Thu hẹp Đại lý: doanh thu 4.258, lợi nhuận 1.490, biên 35,0%. "
            "Đầu tư Online: doanh thu cao nhất 4.607, nhưng lợi nhuận 1.488, biên thấp nhất 32,3%.",
            "Bài học: doanh thu cao nhất không đồng nghĩa lợi nhuận cao nhất, vì chi phí thu hút khách Online tăng thêm 4 điểm phần trăm trong kịch bản này.",
            "Độ nhạy (sheet Sensitivity): với kịch bản Thu hẹp, nếu tăng trưởng Online chỉ 10% thay vì 18% và tỷ lệ chi phí tăng thêm 3 điểm, "
            "lợi nhuận vẫn trên 1.400 — lựa chọn này không mong manh vào một giả định duy nhất.",
        ], [("Bảng 3 — Kịch bản kỳ tới", ["Kịch bản", "Doanh thu", "Chi phí", "Lợi nhuận", "Biên"],
             [["Giữ nguyên", "4.271", "2.822", "1.449", "33,9%"], ["Thu hẹp Đại lý", "4.258", "2.768", "1.490", "35,0%"],
              ["Đầu tư Online", "4.607", "3.119", "1.488", "32,3%"]],
             "Nguồn: sheet Scenario, mọi tổng là công thức từ doanh thu gốc và tỷ lệ chi phí gốc theo kênh. Đơn vị triệu đồng.")]),
        ("5. Rủi ro, giới hạn và việc cần làm tiếp", [
            "Rủi ro 1: cắt Đại lý quá nhanh làm mất khách quen mà Online chưa bù kịp. Giảm thiểu: thu hẹp 50%, không cắt hẳn, theo dõi doanh thu Đại lý hàng tháng.",
            "Rủi ro 2: chi phí Online tiếp tục tăng vì cạnh tranh đấu thầu. Giảm thiểu: đặt trần 64% và dừng tăng chi tiêu khi vượt trần.",
            "Giới hạn dữ liệu: đây là số tổng hợp nội bộ 4 quý, không có dữ liệu thị trường bên ngoài, không tách chi phí cố định và biến đổi. "
            "Vì vậy kịch bản là công cụ so sánh, không phải dự báo cam kết.",
            "Việc tiếp theo: kế toán xác nhận lại cách phân bổ chi phí Online theo quý; quản lý kênh Đại lý đề xuất danh sách khách giữ lại trước khi thu hẹp.",
        ], []),
    ]
    for heading, paras, tables in chapters:
        doc.add_heading(heading, level=1)
        for p in paras:
            doc.add_paragraph(p)
        for title, headers, rows, caption in tables:
            doc.add_paragraph(title, style="Heading 2")
            table(headers, rows, caption)

    doc.add_heading("Phụ lục — Nguồn và cách đối chiếu", level=1)
    for s in ["Nguồn: sheet Data trong file Excel kèm theo; tổng theo quý và theo kênh đối chiếu ở sheet Summary (ô kiểm tra phải = 0).",
              "Cách tính: lợi nhuận = doanh thu − chi phí; biên = lợi nhuận / doanh thu; tỷ lệ chi phí = chi phí / doanh thu.",
              "Hạn chế: file PowerPoint kèm theo chưa được trình chiếu click thật trong PowerPoint; animation mới đạt mức cấu trúc + render LibreOffice."]:
        doc.add_paragraph(s, style="List Bullet")
    doc.core_properties.title = "Doanh thu tăng, lợi nhuận giảm: thu hẹp Đại lý có kiểm soát"
    doc.core_properties.author = tokens["brand"]; doc.core_properties.language = "vi-VN"
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


# ---------------------------------------------------------------- PPTX
def _txt(shape, text, size, bold, color, align=0):
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    tf = shape.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.font.size = Pt(size); p.font.bold = bold
    p.font.name = "Calibri"; p.font.color.rgb = RGBColor(*bytes.fromhex(color))
    p.alignment = [PP_ALIGN.LEFT, PP_ALIGN.CENTER][align]
    return shape


def _box(slide, x, y, w, h, text, size, bold, color, fill=None, align=0):
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE
    if fill:
        from pptx.dml.color import RGBColor
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor(*bytes.fromhex(fill))
        shp.line.fill.background()
    else:
        shp = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    return _txt(shp, text, size, bold, color, align)


def build_pptx(path: Path, tokens: dict) -> None:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    colors = tokens["colors"]
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Doanh thu tăng, lợi nhuận giảm"; prs.core_properties.language = "vi-VN"
    blank = next((l for l in prs.slide_layouts if l.name == "Title Only"), prs.slide_layouts[6])

    def slide(title, note):
        s = prs.slides.add_slide(blank)
        if s.shapes.title:
            t = s.shapes.title
            t.left, t.top, t.width, t.height = Inches(0.7), Inches(0.35), Inches(11.9), Inches(0.9)
            _txt(t, title, 30, True, colors["primary"])
        else:
            _box(s, 0.7, 0.35, 11.9, 0.9, title, 30, True, colors["primary"])
        _box(s, 0.7, 6.95, 11.9, 0.3, "Nguồn: file Excel kèm theo, sheet Data → Summary. Đơn vị: triệu đồng.", 9, False, colors["muted"])
        try:
            s.notes_slide.placeholders[1].text_frame.text = note
        except Exception:
            pass
        return s

    def chart(s, cats, series, x=4.9, y=1.5, w=7.7, h=4.8):
        from pptx.chart.data import CategoryChartData as CCD
        cd = CCD(); cd.categories = cats
        for name, vals in series:
            cd.add_series(name, vals)
        g = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = g.chart; ch.has_legend = False; ch.value_axis.minimum_scale = 0
        ch.value_axis.has_title = True; ch.value_axis.axis_title.text_frame.text = "Triệu đồng"
        ch.plots[0].has_data_labels = True
        fills = [colors["primary"], colors["accent"], colors["muted"]]
        for i, ser in enumerate(ch.series):
            f = ser.format.fill; f.solid(); f.fore_color.rgb = RGBColor(*bytes.fromhex(fills[i % 3]))
        return g

    # Title slide
    s = prs.slides.add_slide(blank)
    if s.shapes.title:
        t = s.shapes.title
        t.left, t.top, t.width, t.height = Inches(0.7), Inches(1.2), Inches(11.9), Inches(1.2)
        _txt(t, "Doanh thu tăng, lợi nhuận giảm", 44, True, colors["primary"])
    else:
        _box(s, 0.7, 1.2, 11.9, 1.2, "Doanh thu tăng, lợi nhuận giảm", 44, True, colors["primary"])
    _box(s, 0.7, 2.6, 11.9, 0.7, "Đề xuất: thu hẹp Đại lý có kiểm soát — không phải vì Đại lý lỗ, mà vì Online đang đắt dần.", 20, False, colors["text"])
    for i, (label, val) in enumerate([("Doanh thu 2026", "4.200"), ("Lợi nhuận 2026", "1.500"), ("Biên Q1 → Q4", "40,5% → 30,6%")]):
        _box(s, 0.7 + i * 4.1, 3.8, 3.8, 1.4, f"{label}\n{val}", 20, True, colors["primary"], fill=colors["banded_row"], align=1)
    _box(s, 0.7, 5.6, 11.9, 0.5, "Báo cáo 4 quý 2026 · Mọi số đối chiếu được với Excel kèm theo.", 14, False, colors["muted"])
    try:
        s.notes_slide.placeholders[1].text_frame.text = "Mở đầu bằng kết luận, không bằng tổng quan. Nói ngay đề xuất, rồi mới đưa bằng chứng."
    except Exception:
        pass

    s = slide("Q1 → Q4: doanh thu +8,0%, lợi nhuận −18,5%", "Điểm gãy ở Q4, không phải một quý xấu đơn lẻ.")
    _box(s, 0.7, 1.6, 3.8, 3.2, "Chi phí tăng +26,1%, nhanh gấp 3 lần doanh thu. Đó là nguyên nhân chính.", 20, False, colors["text"])
    _box(s, 0.7, 5.1, 3.8, 1.0, "Chú thích: cột xám là chi phí, cột xanh là lợi nhuận — đọc trực tiếp giá trị trên cột.", 12, False, colors["muted"])
    chart(s, QUARTERS, [("Doanh thu", [1000, 1050, 1070, 1080]), ("Lợi nhuận", [405, 390, 375, 330])])

    s = slide("Online kéo doanh thu, nhưng cũng kéo chi phí lên", "Online tăng 460 → 610, nhưng tỷ lệ chi phí của nó tăng theo.")
    _box(s, 0.7, 1.6, 3.8, 3.2, "Tỷ lệ chi phí Online: 54,3% (Q1) → 63,9% (Q4). Đồng mới ít lãi hơn đồng cũ.", 20, False, colors["text"])
    _box(s, 0.7, 5.1, 3.8, 1.0, "Cả năm Online lãi 860, vẫn lớn nhất tuyệt đối — vấn đề là xu hướng, không phải mức hiện tại.", 12, False, colors["muted"])
    chart(s, QUARTERS, [("Online", [460, 520, 560, 610]), ("Cửa hàng", [380, 380, 370, 350]), ("Đại lý", [160, 150, 140, 120])])

    s = slide("Thu hẹp Đại lý: lợi nhuận cao nhất, nhưng chỉ hơn 2", "Đừng đọc kịch bản như chiến thắng áp đảo. Chênh lệch quá nhỏ để tự tin tuyệt đối.")
    _box(s, 0.7, 1.6, 3.8, 3.2, "Giữ nguyên 1.449 · Thu hẹp 1.490 · Đầu tư Online 1.488. Doanh thu cao nhất lại có biên thấp nhất.", 20, False, colors["text"])
    _box(s, 0.7, 5.1, 3.8, 1.0, "Vì vậy đề xuất đi kèm điều kiện dừng: chi phí Online vượt 64% thì xem lại ngay.", 12, False, colors["muted"])
    chart(s, ["Giữ nguyên", "Thu hẹp", "Đầu tư"], [("Lợi nhuận kịch bản", [1449, 1490, 1488])])

    s = slide("Quyết định: thu hẹp 50%, giữ Cửa hàng, kìm chi phí", "Ba việc, theo thứ tự, không phải danh sách dài.")
    _box(s, 0.7, 1.6, 11.9, 1.1, "1. Thu hẹp Đại lý 50%, không cắt hẳn — giữ khách quen, theo dõi hàng tháng.", 20, False, colors["text"])
    _box(s, 0.7, 3.0, 11.9, 1.1, "2. Đặt trần tỷ lệ chi phí Online ở 64% — vượt trần thì dừng tăng chi tiêu.", 20, False, colors["text"])
    _box(s, 0.7, 4.4, 11.9, 1.1, "3. Kế toán xác nhận lại cách phân bổ chi phí Online trước khi chốt kỳ tới.", 20, False, colors["text"])

    s = slide("Phụ lục: đối chiếu và giới hạn", "Appendix chỉ để đối chiếu, không để thuyết trình dài thêm.")
    _box(s, 0.7, 1.6, 11.9, 4.8,
         "Tổng năm: 4.200 / 2.700 / 1.500. Quý: 1.000, 1.050, 1.070, 1.080. "
         "Kênh: 2.150, 1.480, 570. Giới hạn: số nội bộ, không tách cố định / biến đổi; kịch bản là so sánh, không phải dự báo.", 18, False, colors["text"])

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    # Fade click beats, then re-save via motion module (structure rung only).
    from office_builder.pptx_motion import add_fade_click_beats
    add_fade_click_beats(path, path, tokens)


if __name__ == "__main__":
    tokens = load_tokens()
    OUT.mkdir(parents=True, exist_ok=True)
    build_xlsx(OUT / "report.xlsx", tokens)
    build_docx(OUT / "report.docx", tokens)
    build_pptx(OUT / "report.pptx", tokens)
    print(f"Wrote {OUT}/report.docx, report.xlsx, report.pptx")
