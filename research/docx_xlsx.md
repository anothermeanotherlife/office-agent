# Nghiên cứu DOCX + XLSX cho AI agent — nền tảng skill Office đa agent

**Kết luận kiến trúc trước khi đi vào chi tiết:** Đừng để agent viết format trực tiếp từng đoạn/cell. Chuẩn đúng là: `theme.json + template (.docx/.xlsx mẫu) + data.json đã kiểm chứng → builder → audit bằng code → render lại`. Agent tìm kiếm chỉ lo dữ liệu, agent builder chỉ lo cấu trúc theo theme, agent audit chỉ chạy checklist dưới đây và **fail là không giao file**.

## 1. DOCX — python-docx (khuyến nghị 1.2.0)

### 1.1. Những khối phải nắm

| Khối | API / cách làm chuẩn | Lưu ý cho agent |
|---|---|---|
| Styles | `doc.styles['Normal']`, `['Heading 1'..'Heading 3']`, `['Title']`, `['Caption']`, `['List Bullet'/'List Number']`. Sửa `style.font.name/size/color`, `paragraph_format.space_before/after, line_spacing, keep_with_next`. Đặt font cho cả `w:rFonts` ascii + hAnsi + eastAsia nếu có tiếng Việt | **Không format trực tiếp bằng run** trừ trường hợp đặc biệt. Style không tồn tại → `KeyError`, nên builder phải assert style có trong template trước |
| Heading / Outline | `doc.add_heading(text, level=0..3)`. Level 0 = Title. Heading là xương sống của TOC và Navigation Pane | Kiểm chứng: thứ tự H1→H2→H3 không nhảy cấp, không có heading rỗng, không dùng bold giả làm heading |
| Section / Page setup | `doc.sections[0]`, `doc.add_section(WD_SECTION.NEW_PAGE)`. `section.page_width/page_height, orientation, top/bottom/left/right_margin, header_distance, footer_distance, gutter`. Đổi sang Landscape phải hoán đổi width/height | Mỗi section có header/footer riêng. Đổi margin mà không đổi section = đổi toàn bộ tài liệu |
| Header / Footer | `section.header`, `section.footer`, `section.first_page_header/footer`, `section.different_first_page_header_footer=True`. **Phải set `is_linked_to_previous=False` trước khi sửa**, nếu không sẽ sửa luôn section trước | Logo đặt trong header bằng bảng 1 hàng 3 cột hoặc paragraph + tab stop, không chèn ảnh trôi tự do bằng python-docx thuần |
| Page number | Không có API sẵn. Chèn field bằng lxml: `w:fldChar begin → w:instrText PAGE → w:fldChar end`, tương tự `NUMPAGES`. Footer dạng “Trang X / Y” | Word/LibreOffice mới tính số trang khi mở file, python-docx **không tính** |
| Table | `doc.add_table(rows, cols)`, `table.style='Table Grid' / 'Light Grid Accent 1'`, `table.autofit=False`, set `tblW` + `gridCol` sao cho tổng = chiều rộng trang trừ margin. Header lặp lại: set `tblHeader` cho hàng đầu, `cantSplit` cho hàng không muốn tách trang | `cell.text = ...` sẽ reset run về plain. Muốn giữ format phải ghi vào `cell.paragraphs[0].runs[0]` hoặc tạo run mới với style |
| TOC | python-docx thuần **không có API TOC**. Cách chuẩn: chèn field `TOC \o "1-3" \h \z \u` bằng `OxmlElement('w:fldChar'/'w:instrText')` begin → separate → end, và set `updateFields` để Word hỏi cập nhật khi mở | TOC chỉ là field rỗng cho tới khi Word/LibreOffice update. Audit chỉ kiểm được “có field TOC + có đủ Heading 1-3”, không kiểm được số trang TOC nếu chưa render |
| Numbering | `style='List Bullet'/'List Number'` dùng numbering mặc định của Word. Muốn restart, multilevel, định dạng riêng phải tự viết `numbering.xml` (abstractNum/num) hoặc dùng thư viện mở rộng | **Lỗi kinh điển:** 2 danh sách riêng trong 1 tài liệu bị nối số tiếp nhau thay vì restart |
| Image / Caption | `doc.add_picture(path, width=Inches(x))` — chỉ truyền width để giữ tỉ lệ. Caption: python-docx thuần không có `add_caption`; làm bằng paragraph style `Caption` + field `SEQ Figure \* ARABIC` qua lxml | Kiểm tra ảnh: tồn tại trong `word/media/`, có relationship, không vượt quá chiều rộng trang, có alt text (`descr`) nếu làm báo cáo chuyên nghiệp |
| Template | **Ưu tiên template-first:** tạo `.docx` mẫu trong Word sẵn style/header/footer, rồi `Document('template.docx')` hoặc `docxtpl.DocxTemplate` với `{{ bien }}`, `{% for %}`, `{% if %}`, `{%tr` cho hàng bảng | docxtpl giữ style của placeholder, nhưng tag Jinja phải nằm trong cùng 1 run, nếu Word tự tách run sẽ vỡ template |

### 1.2. Lỗi thường gặp DOCX và cách kiểm chứng bằng code

| Lỗi | Nguyên nhân | Cách sửa / kiểm chứng |
|---|---|---|
| Style trôi, mỗi đoạn một font | Format trực tiếp, copy từ nguồn khác, style không tồn tại | Audit: duyệt mọi paragraph, assert `p.style.name` thuộc whitelist; duyệt run, cấm font/size/color lệch theme trừ danh sách cho phép |
| Bảng vỡ, tràn lề, cột méo | Không set `tblW/gridCol`, `autofit=True`, tổng cột > khổ giấy | Audit: tính tổng `gridCol` (twips) ≤ `page_width - left - right_margin`; reload và đếm `len(table.columns)`, kiểm tra hàng header có `tblHeader` |
| Font thiếu / lỗi tiếng Việt | Dùng font không có glyph tiếng Việt, không set eastAsia/cs | Dùng font phổ biến (Calibri, Arial, Times New Roman) và kiểm tra render bằng LibreOffice → PDF, không chỉ tin file XML |
| TOC trống, số trang sai | Tưởng python-docx tự tính TOC/PAGE | Audit XML: tìm `TOC` và `PAGE` trong `document.xml`/header/footer; ghi chú rõ cho người dùng “mở Word → Update Fields” |
| Heading giả, TOC không bắt | Dùng bold + cỡ lớn thay Heading | Audit: H1/H2 phải dùng đúng style, không có heading rỗng |
| File mở được nhưng Word báo lỗi | XML hỏng, relationship ảnh/style thiếu | Audit: `zipfile.ZipFile.testzip() is None`, đủ part `word/document.xml`, `word/styles.xml`, mọi `r:embed` đều có target trong `word/_rels/document.xml.rels` |
| Thay text làm mất format | `paragraph.text = ...` / `cell.text = ...` xóa run | Khi replace, giữ run đầu và chỉ đổi `run.text`, hoặc dùng `python-docx-replace` / docxtpl; với run bị tách, phải normalize trước |

**Script audit DOCX tối thiểu:**
```python
from docx import Document
import zipfile
assert zipfile.ZipFile("out.docx").testzip() is None
d = Document("out.docx")
assert [p.style.name for p in d.paragraphs if p.text.strip()]
assert all(s in [s.name for s in d.styles] for s in ["Heading 1","Heading 2","Normal"])
xml = zipfile.ZipFile("out.docx").read("word/document.xml").decode()
assert "TOC" in xml and "PAGE" in xml  # nếu tài liệu yêu cầu TOC + số trang
for t in d.tables:
    assert t.style.name in {"Table Grid","Light Grid Accent 1"}
```

### 1.3. Khi nào cần vượt qua python-docx thuần

`python-docx-ng` (fork, cùng namespace `docx`, **không cài chung với python-docx**) bổ sung: TOC/field builder, `add_caption()` kèm SEQ + bookmark `_Ref`, numbering đọc/ghi + restart, floating image, watermark, footnote/endnote, delete/copy có sửa relationship. Chỉ dùng khi skill thật sự cần các tính năng đó, vì nó tăng rủi ro tương thích.

## 2. XLSX — openpyxl (đọc/sửa) + XlsxWriter (tạo mới đẹp, nhanh)

**Chọn thư viện theo việc, không chọn theo thói quen:**

- **openpyxl:** đọc + sửa file có sẵn, template, formula, chart, validation, conditional formatting. Là lựa chọn mặc định cho skill này.
- **XlsxWriter:** chỉ ghi, không đọc/sửa. Nhanh hơn, format/chart/rich-text trong cell tốt hơn, `constant_memory` cho file lớn. Dùng khi tạo báo cáo mới hoàn toàn, dung lượng lớn.
- **pandas:** chỉ chuẩn bị dữ liệu (`pivot_table`, groupby), xuất qua `ExcelWriter(engine='openpyxl'/'xlsxwriter')`. Không dùng pandas để format — nó làm mất format khi ghi đè.

### 2.1. Những khối phải nắm (openpyxl)

| Khối | API chuẩn | Lưu ý |
|---|---|---|
| Sheet | `wb.active`, `wb.create_sheet("Data")`, `ws.sheet_properties.pageSetUpPr.fitToPage=True`, `ws.sheet_state='hidden'` cho sheet dữ liệu thô | Đặt tên sheet không dấu, ≤31 ký tự, không chứa `[]:*?/\`. Sheet đầu tiên phải là sheet người đọc cần thấy, không phải sheet dữ liệu thô |
| Style | `Font(name, size, bold, color)`, `PatternFill`, `Border/Side`, `Alignment(wrap_text, vertical, horizontal)`, `NamedStyle`. **Tạo 1 lần, tái sử dụng** — style object là immutable, muốn đổi phải tạo object mới | Tạo Font/Fill mới trong vòng lặp hàng nghìn dòng là nguyên nhân file phình và chậm |
| Number format | `cell.number_format = '#,##0'`, `'#,##0.00'`, `'#,##0_);[Red](#,##0)'`, `'"$"#,##0.00'`, `'yyyy-mm-dd'`, `'dd/mm/yyyy'`, `'0.00%'`. Giá trị lưu dạng số/ngày thật, format chỉ là cách hiển thị | Xem mục locale ở 2.2 — đây là lỗi hay gặp nhất |
| Formula | `cell.value = "=SUM(B2:B10)"`, `"=IF(...)"`, `"=VLOOKUP(...)"`, `"=SUMIFS(...)"`. openpyxl **chỉ ghi công thức, không tính** | Muốn kiểm giá trị phải mở bằng Excel/LibreOffice để recalc rồi đọc lại bằng `load_workbook(data_only=True)` |
| Conditional formatting | `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule`, `ws.conditional_formatting.add("B2:B100", rule)` | Rule chồng nhau gây khó đọc; giới hạn 2–3 rule/sheet, ưu tiên CellIs/FormulaRule có ngưỡng rõ ràng |
| Data validation | `DataValidation(type="list"/"decimal"/"date", formula1='"A,B,C"' hoặc "0,100", allow_blank=True, showErrorMessage=True)`, `ws.add_data_validation(dv)`, `dv.add("C2:C100")` | Validation chỉ cảnh báo khi nhập tay trong Excel, không chặn code ghi sai — audit vẫn phải kiểm dữ liệu đầu vào |
| Pivot-like summary | openpyxl **không tạo native PivotTable mới từ đầu** (chỉ giữ/sửa pivot có sẵn). Cách thực tế cho agent: sheet `Data` (dữ liệu sạch) + sheet `Summary` tính bằng `SUMIFS/COUNTIFS/AVERAGEIFS` hoặc pandas `pivot_table` xuất thành bảng thường, ghi rõ “bảng tổng hợp, không phải PivotTable native” | Nếu bắt buộc PivotTable thật: cần template Excel đã có pivot + refresh khi mở, hoặc xlwings/Excel thật (không phù hợp chạy server) |
| Chart (openpyxl) | `BarChart/LineChart/PieChart`, `Reference(ws, min_col, min_row, max_row)`, `chart.add_data()`, `chart.set_categories()`, `chart.title, chart.style, chart.y_axis.numFmt`, `ws.add_chart(chart, "E2")` | Chart phải trỏ đúng range, có title, có nhãn trục, không chart từ range rỗng; XlsxWriter cho chart đẹp và tùy biến hơn nếu tạo mới |
| Print setup | `ws.page_setup.orientation/paperSize/fitToWidth/fitToHeight`, `ws.print_title_rows='1:1'`, `ws.print_area='A1:H50'`, `ws.page_margins`, `ws.oddHeader/oddFooter` | Báo cáo không set print setup sẽ in vỡ trang — audit bắt buộc cho file gửi khách |
| Freeze / Filter | `ws.freeze_panes='A2'` (đóng hàng 1), `'B2'` (đóng hàng 1 + cột A), `ws.auto_filter.ref='A1:H100'` | Trong `write_only=True`, phải set freeze/filter **trước** khi append dữ liệu |

### 2.2. Lỗi thường gặp XLSX và cách kiểm chứng

| Lỗi | Nguyên nhân | Cách sửa / kiểm chứng |
|---|---|---|
| Formula lỗi / hiện 0 / `#VALUE!` | Ghi công thức sai cú pháp, tham chiếu range sai, openpyxl không tính nên tưởng đúng | Audit 2 lớp: (1) `data_only=False` kiểm tra chuỗi formula, range tồn tại, không có `#REF!`; (2) sau khi LibreOffice recalc, `data_only=True` kiểm tra giá trị, tổng cross-check bằng Python (tổng Summary phải bằng tổng Data) |
| Số thành text, ngày thành text | Ghi `"1,000"`, `"12/01/2026"` dạng string thay vì `1000`, `datetime` | Audit: `isinstance(cell.value, (int,float))` với cột số, `isinstance(..., datetime)` với cột ngày; không có ô số nào `data_type=='s'` |
| Định dạng số sai locale | Dùng format theo locale máy dev, dấu `,`/`.` đảo ngược, tiền tệ không rõ đơn vị | **Quy tắc:** lưu số thuần, format code luôn viết theo chuẩn Excel en-US (`#,##0.00`, `yyyy-mm-dd`); hiển thị theo locale là việc của Excel khi mở. Test: đọc lại `cell.number_format` và so với bảng format chuẩn của theme, không so chuỗi hiển thị |
| File nặng, mở chậm | Style tạo mới từng cell, ảnh không nén, hàng nghìn merged cell, toàn bộ cột được format | Tái dùng style object, `write_only=True` cho file lớn, nén ảnh trước khi chèn, chỉ format đúng range dữ liệu, tránh merge tràn lan. Audit: kích thước file + thời gian `load_workbook(read_only=True)` duyệt |
| Merged cell mất dữ liệu | Chỉ ô top-left giữ giá trị, các ô còn lại thành `MergedCell` | Audit: không merge vùng có dữ liệu ở nhiều ô; style chỉ đặt ở top-left; khi đọc phải skip `MergedCell` |
| Copy sheet mất validation/print | `copy_worksheet` không copy hết | Khi copy phải copy tay: `data_validations`, `page_setup`, `page_margins`, `print_titles`, `print_area`, `conditional_formatting` |
| Chart không dữ liệu / sai trục | Range sai, categories không khớp data, chart đặt đè lên dữ liệu | Audit: `len(ws._charts)`, mỗi chart có `series` trỏ đúng sheet/range, title không rỗng; kiểm tra bằng render LibreOffice → PDF nếu là báo cáo quan trọng |
| Sheet lộn xộn, không in được | Không freeze, không filter, không print title, cột quá rộng/hẹp | Audit: `freeze_panes is not None`, `auto_filter.ref` đúng header, `print_title_rows` có, `page_setup.fitToWidth/fitToHeight` đã set |

**Script audit XLSX tối thiểu:**
```python
from openpyxl import load_workbook
wb = load_workbook("out.xlsx")  # data_only=False: kiểm formula
ws = wb["Summary"]
assert ws.freeze_panes == "A2"
assert ws["B2"].value.startswith("=")  # formula, không phải giá trị gõ tay
assert ws["B2"].number_format in {"#,##0", "#,##0.00"}
assert list(ws.conditional_formatting._cf_rules)
assert list(ws.data_validations.dataValidation)
assert len(ws._charts) >= 1
wb2 = load_workbook("out.xlsx", data_only=True)  # chỉ có nghĩa sau khi Excel/LibreOffice đã recalc
```

## 3. Giữ nhất quán thương hiệu giữa Word, Excel, PowerPoint

Đừng hardcode màu/font ở từng builder. Một file `theme.json` duy nhất là source of truth, 3 builder cùng đọc:

```json
{
  "brand": "Tên thương hiệu",
  "colors": {"primary": "0F2A44", "accent": "1B7A6E", "text": "1F2937", "muted": "6B7280", "header_bg": "0F2A44", "header_text": "FFFFFF", "banded_row": "F2F4F7"},
  "fonts": {"heading": "Calibri", "body": "Calibri", "mono": "Consolas"},
  "sizes": {"h1": 20, "h2": 16, "body": 11, "table_header": 11, "caption": 9},
  "logo": "assets/logo.png",
  "page": {"size": "A4", "margin_cm": 2.0},
  "number_formats": {"int": "#,##0", "money": "#,##0", "pct": "0.0%", "date": "dd/mm/yyyy"}
}
```

Bảng ánh xạ:

| Thành phần theme | DOCX | XLSX | PPTX (để đồng bộ, dù task này tập trung DOCX/XLSX) |
|---|---|---|---|
| Màu chính | `Heading 1` color, table header shading, đường kẻ | Header fill `PatternFill`, chart series color | Shape fill, title color, theme accent |
| Font | `Normal` + Heading styles, `rFonts` đủ ascii/hAnsi/eastAsia | `Font(name=...)` cho mọi cell, `NamedStyle` | Placeholder font, chart font |
| Logo/header | Header: logo + tên tài liệu; Footer: “Trang X / Y” + confidential nếu cần | Print header/footer: `oddHeader.center.text`, logo qua `ws.add_image` ở sheet bìa nếu cần | Slide master: logo góc, footer số slide |
| Bảng | 1 style bảng duy nhất, header đậm + nền primary + chữ trắng, hàng xen kẽ `banded_row` | Header fill primary, border mỏng, banded row bằng fill xen kẽ hoặc Table style | Bảng 1 style, cùng màu header |
| Quy tắc chung | Tối đa 3–4 màu, tỉ lệ 60-30-10; contrast chữ/nền ≥ 4.5:1; không dùng màu đỏ-xanh lá làm cặp phân biệt duy nhất (mù màu) | Cùng quy tắc; conditional formatting dùng cùng palette, không tự phát minh màu mới | Cùng quy tắc |

Lưu ý quan trọng: theme màu của Office (theme color + tint) và màu RGB trực tiếp là 2 cơ chế khác nhau. Với agent, **luôn dùng RGB hex cố định** để DOCX/XLSX/PPTX ra đúng màu, không phụ thuộc theme máy người nhận. Chỉ khi làm template Office native mới dùng theme color.

## 4. Tiêu chuẩn chất lượng kiểm chứng được (checklist PASS/FAIL cho agent Audit)

### DOCX — FAIL nếu vi phạm bất kỳ dòng nào
- [ ] Mở được bằng `zipfile` + python-docx, không lỗi XML/relationship.
- [ ] Có Title + H1, outline không nhảy cấp, không heading rỗng, không heading giả bằng bold.
- [ ] Mọi paragraph/table đều dùng style trong whitelist của theme; không còn run lệch font/size/màu ngoài danh sách cho phép.
- [ ] Section: khổ giấy, orientation, margin đúng spec; header/footer đã unlink, có logo/tên + số trang (field PAGE/NUMPAGES) nếu yêu cầu.
- [ ] Bảng: tổng chiều rộng cột ≤ khổ giấy trừ margin, header lặp lại, không có cell trống bất ngờ, số liệu trong bảng khớp data nguồn (đối chiếu từng cell).
- [ ] Ảnh: tồn tại, không vượt lề, có caption đúng dạng `Hình X. ...` + SEQ field nếu yêu cầu cross-reference.
- [ ] TOC: có field `TOC \o "1-3"`, và tài liệu có đủ Heading 1–3 để TOC bắt được; ghi chú “cần Update Fields khi mở Word”.
- [ ] Nội dung: không placeholder (`{{ }}`, TBD, lorem), không text bị tách run làm sai chính tả, số liệu + ngày tháng đúng nguồn.

### XLSX — FAIL nếu vi phạm bất kỳ dòng nào
- [ ] Mở được bằng openpyxl, sheet đúng tên/thứ tự, không sheet trống, không sheet rác.
- [ ] Kiểu dữ liệu đúng: số là số, ngày là datetime, text là text; không ô số dạng text.
- [ ] Formula: mọi ô tổng/phần trăm đều là formula (không gõ tay), range đúng, sau recalc giá trị khớp Python cross-check (sai số cho phép < 0.01).
- [ ] Format: number_format đúng bảng theme cho từng cột; ngày dùng 1 format thống nhất toàn workbook.
- [ ] Summary khớp Data: tổng, count, average trên Summary = tính lại từ Data; không có `#REF!`, `#VALUE!`, `#DIV/0!` sau recalc.
- [ ] Conditional formatting + validation: rule đúng range, validation đúng loại/range, không rule trùng lặp vô nghĩa.
- [ ] Chart: có title, trục có nhãn, series trỏ đúng range, không chart rỗng.
- [ ] Trải nghiệm + in: freeze panes, auto filter, print title rows, print area, fitToPage, orientation/paper size đúng; cột đủ rộng để không hiện `###`.
- [ ] Hiệu năng: file không phình bất thường so với dữ liệu (cảnh báo nếu > vài MB cho dữ liệu nhỏ), không tạo style mới trong vòng lặp.

### Quy trình audit đề xuất cho repo đa agent
1. **Search agent** trả về data có nguồn, không tự bịa số.
2. **Filter agent** chốt data.json + spec (khổ giấy, style, bảng, chart) — đây là hợp đồng, builder không được tự đổi.
3. **Builder agent** chỉ chạy builder từ spec + theme, không sửa số liệu.
4. **Trial agent** chạy builder trên dữ liệu mẫu nhỏ trước, kiểm tra mở file + render LibreOffice → PDF để xem vỡ trang.
5. **Audit agent** chạy checklist trên + cross-check số liệu, trả về PASS/FAIL kèm danh sách lỗi cụ thể theo file/dòng/cell.
6. **Standard agent** chỉ cập nhật theme/spec khi có lỗi lặp lại ≥ 2 lần, không sửa từng file thủ công.
7. Lặp lại tới khi Audit PASS toàn bộ — và giữ lại audit log làm bằng chứng, không chỉ nói “đã tốt nhất” bằng cảm tính.

## 5. Nguồn tham khảo (chỉ dùng URL trả về từ search, không tự đoán)

- python-docx, Sections / page setup: https://python-docx.readthedocs.io/en/latest/user/sections.html
- python-docx, Document API (heading, paragraph, picture, section): https://python-docx.readthedocs.io/en/latest/api/document.html?highlight=page_break
- python-docx, Header/Footer phân tích: https://python-docx.readthedocs.io/en/develop/dev/analysis/features/header.html
- TOC bằng field lxml cho python-docx 1.2.0, cú pháp `TOC \o "1-3" \h \z \u`: https://copyprogramming.com/howto/update-the-toc-table-of-content-of-ms-word-docx-documents-with-python
- Pitfalls python-docx (field phải do Word tính, style phải tồn tại, numbering restart, cell.text reset format): https://www.getclaudeskills.com/skills/docx-skill-nousresearch
- python-docx-ng — TOC/field, caption SEQ, numbering restart, floating image: https://libraries.io/pypi/python-docx-ng
- docxtpl / Jinja template, `{%tr` cho bảng, tag phải cùng run: https://www.softkraft.co/python-word-automation/
- openpyxl, Conditional Formatting (CellIs, ColorScale, DataBar): https://www.openpyxl.com/conditional-format.html
- openpyxl, Styles (Font/Fill/Border/NamedStyle, style immutable): https://openpyxl.readthedocs.io/en/2.5/styles.html
- openpyxl, Pivot API — chỉ giữ/sửa pivot có sẵn, không tạo mới từ đầu: https://openpyxl.readthedocs.io/en/stable/api/openpyxl.pivot.table.html
- openpyxl, freeze/filter/iter_rows/read_only: https://Zetcode.com/python/openpyxl/
- So sánh openpyxl vs XlsxWriter vs pandas (đọc/sửa vs chỉ ghi vs nhanh): https://micropyramid.com/blog/generating-csv-excel-files-using-python/
- Office theme dùng chung Word/Excel/PowerPoint: https://www.techrepublic.com/article/how-to-create-and-use-the-same-office-themes-across-files-in-microsoft-365/
- Office document themes (Microsoft Learn): https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/use-document-themes-in-your-powerpoint-add-ins

**Giới hạn cần ghi vào README của repo, tránh hứa quá:** python-docx không render layout/PDF và không tính TOC/số trang; openpyxl không tính formula và không tạo PivotTable native mới; cả hai đều không kiểm tra được “đẹp” nếu không render qua Word/Excel/LibreOffice thật. Vì vậy audit cuối bắt buộc phải có bước render (LibreOffice headless → PDF) cho file quan trọng, còn audit code chỉ là vòng 1.
