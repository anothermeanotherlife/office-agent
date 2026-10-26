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


# ---------------------------------------------------------------- XLSX (XlsxWriter)
# Why XlsxWriter, not openpyxl, for this file: openpyxl native charts render
# in LibreOffice with a garbled legend ("Q1; Doanh thu; Q2; Doanh thu; ...")
# even with legend=None and with populated caches. XlsxWriter charts render
# cleanly in LibreOffice (tested 2026-10-26). Visual truth beats audit score.
def build_xlsx(path: Path, tokens: dict) -> None:
    import xlsxwriter

    colors = tokens["colors"]
    wb = xlsxwriter.Workbook(str(path))
    hdr = wb.add_format({"bold": True, "font_color": colors["header_text"], "bg_color": colors["header_bg"], "font_name": "Calibri", "font_size": 11, "border": 1})
    blue = wb.add_format({"font_color": "0000FF", "font_name": "Calibri", "font_size": 11, "border": 1})
    black = wb.add_format({"font_color": "000000", "font_name": "Calibri", "font_size": 11, "border": 1})
    green = wb.add_format({"font_color": "00B050", "font_name": "Calibri", "font_size": 11, "border": 1})
    money_blue = wb.add_format({"font_color": "0000FF", "font_name": "Calibri", "font_size": 11, "border": 1, "num_format": "#,##0"})
    money_black = wb.add_format({"font_color": "000000", "font_name": "Calibri", "font_size": 11, "border": 1, "num_format": "#,##0"})
    money_green = wb.add_format({"font_color": "00B050", "font_name": "Calibri", "font_size": 11, "border": 1, "num_format": "#,##0"})
    pct_blue = wb.add_format({"font_color": "0000FF", "font_name": "Calibri", "font_size": 11, "border": 1, "num_format": "0.0%"})
    pct_green = wb.add_format({"font_color": "00B050", "font_name": "Calibri", "font_size": 11, "border": 1, "num_format": "0.0%"})
    pct_black = wb.add_format({"font_color": "000000", "font_name": "Calibri", "font_size": 11, "border": 1, "num_format": "0.0%"})
    title_fmt = wb.add_format({"bold": True, "font_color": colors["primary"], "font_name": "Calibri", "font_size": 16})

    def setup(ws, widths, landscape=True):
        ws.freeze_panes(1, 0)
        if landscape:
            ws.set_landscape()
        ws.fit_to_pages(1, 1)
        for col, w in widths.items():
            ws.set_column(col, w)

    readme = wb.add_worksheet("README")
    readme.write(0, 0, "Báo cáo 4 quý 2026 — Doanh thu tăng, lợi nhuận giảm", title_fmt)
    readme.write(1, 0, "Đơn vị: triệu đồng. Nguồn: dữ liệu nội bộ tổng hợp. Mọi tổng là công thức, không gõ tay.")
    readme.write(2, 0, "Màu chữ theo chuẩn FAST: xanh dương = số nhập tay (input), đen = công thức trong sheet, xanh lá = công thức liên kết sheet khác.")
    readme.write(3, 0, "Summary là bảng tổng hợp bằng công thức, không phải PivotTable native.")
    readme.write(4, 0, "Kịch bản và độ nhạy nằm ở sheet Scenario và Sensitivity; đổi ô xanh dương, các tổng tự đổi.")
    readme.set_column("A:A", 120)

    data = wb.add_worksheet("Data")
    for ci, h in enumerate(["Quý", "Doanh thu", "Chi phí", "Kênh", "Lợi nhuận"]):
        data.write(0, ci, h, hdr)
    r = 1
    for q in QUARTERS:
        for ch in CHANNELS:
            rev, cost = RAW[(q, ch)]
            data.write(r, 0, q, blue); data.write_number(r, 1, rev, money_blue); data.write_number(r, 2, cost, money_blue)
            data.write(r, 3, ch, blue); data.write_formula(r, 4, f"=B{r+1}-C{r+1}", money_black, rev-cost)
            r += 1
    data.autofilter(0, 0, 12, 4)
    setup(data, {"A:A": 12, "B:E": 16}); data.repeat_rows(0); data.print_area("A1:E13")

    qtr = wb.add_worksheet("Quarterly")
    for ci, h in enumerate(["Quý", "Doanh thu", "Chi phí", "Lợi nhuận", "Biên"]):
        qtr.write(0, ci, h, hdr)
    for i, q in enumerate(QUARTERS, start=1):
        qtr.write(i, 0, q, green)
        qtr.write_formula(i, 1, f"=SUMIFS(Data!B:B,Data!A:A,A{i+1})", money_green, sum(RAW[(q, ch)][0] for ch in CHANNELS))
        qtr.write_formula(i, 2, f"=SUMIFS(Data!C:C,Data!A:A,A{i+1})", money_green, sum(RAW[(q, ch)][1] for ch in CHANNELS))
        qtr.write_formula(i, 3, f"=B{i+1}-C{i+1}", money_green, sum(RAW[(q, ch)][0]-RAW[(q, ch)][1] for ch in CHANNELS))
        qtr.write_formula(i, 4, f"=IF(B{i+1}=0,0,D{i+1}/B{i+1})", pct_green, sum(RAW[(q, ch)][0]-RAW[(q, ch)][1] for ch in CHANNELS)/sum(RAW[(q, ch)][0] for ch in CHANNELS))
    setup(qtr, {"A:A": 12, "B:D": 16, "E:E": 12})

    chan = wb.add_worksheet("Channel")
    for ci, h in enumerate(["Kênh", "Doanh thu", "Chi phí", "Lợi nhuận", "Tỷ lệ chi phí", "Biên"]):
        chan.write(0, ci, h, hdr)
    for i, ch in enumerate(CHANNELS, start=1):
        chan.write(i, 0, ch, green)
        chan.write_formula(i, 1, f"=SUMIFS(Data!B:B,Data!D:D,A{i+1})", money_green, sum(RAW[(qq, ch)][0] for qq in QUARTERS))
        chan.write_formula(i, 2, f"=SUMIFS(Data!C:C,Data!D:D,A{i+1})", money_green, sum(RAW[(qq, ch)][1] for qq in QUARTERS))
        chan.write_formula(i, 3, f"=B{i+1}-C{i+1}", money_green, sum(RAW[(qq, ch)][0]-RAW[(qq, ch)][1] for qq in QUARTERS))
        chan.write_formula(i, 4, f"=IF(B{i+1}=0,0,C{i+1}/B{i+1})", pct_green, sum(RAW[(qq, ch)][1] for qq in QUARTERS)/sum(RAW[(qq, ch)][0] for qq in QUARTERS))
        chan.write_formula(i, 5, f"=IF(B{i+1}=0,0,D{i+1}/B{i+1})", pct_green, sum(RAW[(qq, ch)][0]-RAW[(qq, ch)][1] for qq in QUARTERS)/sum(RAW[(qq, ch)][0] for qq in QUARTERS))
    setup(chan, {"A:A": 16, "B:D": 16, "E:F": 16})

    summ = wb.add_worksheet("Summary")
    summ.write(0, 0, "Chỉ số", hdr); summ.write(0, 1, "Tổng (công thức liên kết Data)", hdr)
    summ.write(1, 0, "Tổng doanh thu"); summ.write_formula(1, 1, "=SUM(Data!B2:B13)", money_black, 4200)
    summ.write(2, 0, "Tổng chi phí"); summ.write_formula(2, 1, "=SUM(Data!C2:C13)", money_black, 2700)
    summ.write(3, 0, "Lợi nhuận (doanh thu - chi phí)"); summ.write_formula(3, 1, "=B2-B3", money_black, 1500)
    summ.write(4, 0, "Tỷ lệ lợi nhuận"); summ.write_formula(4, 1, "=IF(B2=0,0,B4/B2)", pct_black, 1500/4200)
    summ.write(6, 0, "Đối chiếu: tổng theo quý trừ tổng chung (phải = 0)"); summ.write_formula(6, 1, "=SUM(Quarterly!B2:B5)-B2", money_black, 0)
    summ.write(7, 0, "Đối chiếu: tổng theo kênh trừ tổng chung (phải = 0)"); summ.write_formula(7, 1, "=SUM(Channel!B2:B4)-B2", money_black, 0)
    summ.write(8, 0, "Đối chiếu: lợi nhuận quý Q1→Q4 (Q4 − Q1, phải = −75)"); summ.write_formula(8, 1, "=Quarterly!D5-Quarterly!D2", money_black, -75)
    setup(summ, {"A:A": 52, "B:B": 30})

    scen = wb.add_worksheet("Scenario")
    for ci, h in enumerate(["Kịch bản", "Online +", "Cửa hàng +", "Đại lý +", "Tăng tỷ lệ chi phí (pp)"]):
        scen.write(0, ci, h, hdr)
    inputs = [("Giữ nguyên", 0.08, -0.03, -0.10, 0.02), ("Thu hẹp Đại lý", 0.18, -0.03, -0.50, 0.01), ("Đầu tư Online", 0.25, -0.05, -0.10, 0.04)]
    for i, (name, o, s, a, d) in enumerate(inputs, start=1):
        scen.write(i, 0, name, black)
        for ci, v in enumerate((o, s, a, d), start=1):
            scen.write_number(i, ci, v, pct_blue)
    for ci, h in enumerate(["Kênh", "Doanh thu gốc", "Tỷ lệ chi phí gốc"]):
        scen.write(5, ci, h, hdr)
    for i, ch in enumerate(CHANNELS, start=6):
        scen.write(i, 0, ch, green)
        scen.write_formula(i, 1, f"=Channel!B{i-4}", money_green, [2150, 1480, 570][i-6])
        scen.write_formula(i, 2, f"=Channel!E{i-4}", pct_green, [1290/2150, 1060/1480, 350/570][i-6])
    for ci, h in enumerate(["Kịch bản", "Doanh thu", "Chi phí", "Lợi nhuận", "Biên"]):
        scen.write(10, ci, h, hdr)
    _base_rev = [2150, 1480, 570]; _base_rate = [1290/2150, 1060/1480, 350/570]
    for j, scen_row in enumerate((2, 3, 4), start=11):
        scen.write_formula(j, 0, f"=A{scen_row}", black, inputs[scen_row-2][0])
        rev = "+".join(f"$B${7+k}*(1+{chr(66+k)}{scen_row})" for k in range(3))
        cost = "+".join(f"$B${7+k}*(1+{chr(66+k)}{scen_row})*($C${7+k}+$E${scen_row})" for k in range(3))
        _g = inputs[scen_row-2][1:4]; _d = inputs[scen_row-2][4]
        _rev_v = sum(_base_rev[k] * (1 + _g[k]) for k in range(3))
        _cost_v = sum(_base_rev[k] * (1 + _g[k]) * (_base_rate[k] + _d) for k in range(3))
        scen.write_formula(j, 1, f"={rev}", money_black, _rev_v)
        scen.write_formula(j, 2, f"={cost}", money_black, _cost_v)
        scen.write_formula(j, 3, f"=B{j+1}-C{j+1}", money_black, _rev_v - _cost_v)
        scen.write_formula(j, 4, f"=IF(B{j+1}=0,0,D{j+1}/B{j+1})", pct_black, (_rev_v - _cost_v) / _rev_v)
    setup(scen, {"A:A": 24, "B:E": 16})

    sens = wb.add_worksheet("Sensitivity")
    sens.merge_range(0, 0, 0, 3, "Độ nhạy lợi nhuận — kịch bản Thu hẹp Đại lý (đổi ô xanh ở Scenario, bảng này tự đổi)")
    for ci, h in enumerate(["Tăng trưởng Online \\ Tăng tỷ lệ chi phí Online", 0, 0.01, 0.03]):
        if ci == 0:
            sens.write(2, ci, h, hdr)
        else:
            sens.write_number(2, ci, h, pct_blue)
    for i, g in enumerate((0.10, 0.18, 0.25), start=3):
        sens.write_number(i, 0, g, pct_blue)
        for ci, (col, d) in enumerate(zip("BCD", (0, 0.01, 0.03)), start=1):
            _sens_v = (2150 * (1 + g) * (1 - (1290/2150 + d))
                       + 1480 * 0.97 * (1 - (1060/1480 + 0.01))
                       + 570 * 0.50 * (1 - (350/570 + 0.01)))
            sens.write_formula(i, ci, (
                f"=Scenario!$B$7*(1+$A{i+1})*(1-(Scenario!$C$7+{col}$3))"
                f"+Scenario!$B$8*(1+Scenario!$C$3)*(1-(Scenario!$C$8+Scenario!$E$3))"
                f"+Scenario!$B$9*(1+Scenario!$D$3)*(1-(Scenario!$C$9+Scenario!$E$3))"), money_green, _sens_v)
    setup(sens, {"A:A": 40, "B:D": 16}); sens.print_area("A1:D6"); sens.repeat_rows(2)

    dash = wb.add_worksheet("Dashboard")
    dash.write(0, 0, "Doanh thu theo quý (nguồn: Quarterly; trục cột bắt đầu từ 0)")
    chart = wb.add_chart({"type": "column"})
    chart.add_series({"name": "Doanh thu", "categories": ["Quarterly", 1, 0, 4, 0], "values": ["Quarterly", 1, 1, 4, 1], "data_labels": {"value": True}})
    chart.set_title({"name": "Doanh thu theo quý"})
    chart.set_x_axis({"name": "Quý"}); chart.set_y_axis({"name": "Triệu đồng", "min": 0})
    chart.set_legend({"none": True})
    dash.insert_chart("A3", chart)
    setup(dash, {"A:A": 100})

    wb.set_properties({"title": "Báo cáo 4 quý 2026 — Doanh thu tăng, lợi nhuận giảm", "creator": tokens["brand"]})
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.close()



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
    # Visual QA finding: a TOC field alone renders as an empty paragraph in
    # LibreOffice until the user updates it. So also print a readable contents
    # list directly, so the rendered file is usable even before any update.
    doc.add_paragraph("Mục lục — 1. Kết luận trước · 2. Bằng chứng · 3. Vì sao · "
                      "4. Ba kịch bản · 5. Rủi ro và việc tiếp theo · Phụ lục: Nguồn và cách đối chiếu.")

    def _shade(cell, hex_color):
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
        tc_pr.append(shd)

    def table(headers, rows, caption):
        from docx.shared import RGBColor
        t = doc.add_table(rows=1 + len(rows), cols=len(headers)); t.style = "Table Grid"
        # Repeat header row when a table splits across pages (visual QA finding:
        # the continuation page previously showed data rows with no header).
        tr_pr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); tr_pr.append(th)
        for i, h in enumerate(headers):
            cell = t.rows[0].cells[i]
            r = cell.paragraphs[0].add_run(h); r.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _shade(cell, "0F2A44")
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row):
                t.rows[ri + 1].cells[ci].paragraphs[0].add_run(str(v))
        cap = doc.add_paragraph(caption, style="Caption")
        for r in cap.runs:
            r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

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
        _box(s, 0.7, 6.95, 11.9, 0.3, "Nguồn: Excel kèm theo. Đơn vị: triệu đồng.", 9, False, colors["muted"])
        try:
            s.notes_slide.placeholders[1].text_frame.text = note
        except Exception:
            pass
        return s

    def chart(s, cats, series, x=4.9, y=1.5, w=7.7, h=4.8, legend=True, labels=True):
        from pptx.chart.data import CategoryChartData as CCD
        from pptx.enum.chart import XL_LEGEND_POSITION
        cd = CCD(); cd.categories = cats
        for name, vals in series:
            cd.add_series(name, vals)
        g = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = g.chart; ch.has_legend = legend
        if legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
        ch.value_axis.minimum_scale = 0
        ch.value_axis.has_title = True; ch.value_axis.axis_title.text_frame.text = "Triệu đồng"
        ch.plots[0].has_data_labels = labels
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
    # Visual QA correction: the old annotation said the grey column was cost,
    # but this chart has no cost series at all — navy is revenue, teal is profit.
    _box(s, 0.7, 5.1, 3.8, 1.0, "Chú thích đúng: cột xanh đậm là doanh thu, cột xanh lá là lợi nhuận.", 12, False, colors["muted"])
    chart(s, QUARTERS, [("Doanh thu", [1000, 1050, 1070, 1080]), ("Lợi nhuận", [405, 390, 375, 330])])

    s = slide("Online kéo doanh thu, nhưng cũng kéo chi phí lên", "Online tăng 460 → 610, nhưng tỷ lệ chi phí của nó tăng theo.")
    _box(s, 0.7, 1.6, 3.8, 3.2, "Tỷ lệ chi phí Online: 54,3% (Q1) → 63,9% (Q4). Đồng mới ít lãi hơn đồng cũ.", 20, False, colors["text"])
    _box(s, 0.7, 5.1, 3.8, 1.0, "Cả năm Online lãi 860, vẫn lớn nhất tuyệt đối — vấn đề là xu hướng, không phải mức hiện tại.", 12, False, colors["muted"])
    # No data labels on the 3-series chart: they overlapped the bars in the render.
    chart(s, QUARTERS, [("Online", [460, 520, 560, 610]), ("Cửa hàng", [380, 380, 370, 350]), ("Đại lý", [160, 150, 140, 120])], labels=False)

    s = slide("Thu hẹp Đại lý: lợi nhuận cao nhất, nhưng chỉ hơn 2", "Đừng đọc kịch bản như chiến thắng áp đảo.")
    _box(s, 0.7, 1.6, 3.8, 3.2, "Giữ 1.449 · Thu hẹp 1.490 · Đầu tư 1.488.", 20, False, colors["text"])
    _box(s, 0.7, 5.1, 3.8, 1.0, "Điều kiện dừng: chi phí Online vượt 64% thì xem lại.", 12, False, colors["muted"])
    # Visual QA correction: three bars from 0 to ~1490 look identical, because
    # the difference is 2. A bar chart hides the very point of this slide.
    # Use a comparison table instead, and say so on the slide.
    for i, (name, rev, profit, margin) in enumerate([
            ("Giữ nguyên", "4.271", "1.449", "33,9%"),
            ("Thu hẹp", "4.258", "1.490", "35,0%"),
            ("Đầu tư", "4.607", "1.488", "32,3%")]):
        _box(s, 4.9, 1.6 + i * 1.5, 7.7, 1.2,
             f"{name}: DT {rev} · LN {profit} · Biên {margin}", 20, True, colors["primary"], fill=colors["banded_row"])
    _box(s, 4.9, 6.1, 7.7, 0.5, "Cột từ 0 giấu chênh lệch 2.", 12, False, colors["muted"])

    s = slide("Quyết định: thu hẹp 50%, giữ Cửa hàng, kìm chi phí", "Ba việc, theo thứ tự.")
    _box(s, 0.7, 1.6, 11.9, 1.0, "1. Thu hẹp Đại lý 50%, giữ khách quen.", 20, False, colors["text"])
    _box(s, 0.7, 2.9, 11.9, 1.0, "2. Trần chi phí Online 64%.", 20, False, colors["text"])
    _box(s, 0.7, 4.2, 11.9, 1.0, "3. Kế toán xác nhận phân bổ chi phí.", 20, False, colors["text"])
    _box(s, 0.7, 5.5, 11.9, 1.0, "Vượt trần thì dừng.", 20, True, colors["warning"], fill=colors["banded_row"])

    s = slide("Phụ lục: đối chiếu và giới hạn", "Appendix chỉ để đối chiếu, không để thuyết trình dài thêm.")
    _box(s, 0.7, 1.6, 11.9, 1.0, "Tổng năm: doanh thu 4.200 · chi phí 2.700 · lợi nhuận 1.500.", 18, False, colors["text"])
    _box(s, 0.7, 2.9, 11.9, 1.0, "Theo quý: 1.000, 1.050, 1.070, 1.080. Theo kênh: 2.150, 1.480, 570.", 18, False, colors["text"])
    _box(s, 0.7, 4.2, 11.9, 1.0, "Giới hạn: số nội bộ, không tách cố định / biến đổi; kịch bản là so sánh, không phải dự báo.", 18, False, colors["text"])

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
