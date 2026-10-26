# Office Agent — Hệ đa agent tạo Word, Excel, PowerPoint

> **Mục tiêu:** không phải "nhiều agent nói chuyện cho tới khi đồng ý".
> Đây là **một state machine có cổng kiểm chứng bằng code**: Researcher tìm nguồn,
> Distiller chốt spec, Builder làm file thật, Auditor độc lập + script kiểm tra
> file thật, Standards Keeper giữ luật. Một bản chỉ được gọi là "tốt nhất" khi
> artifact cụ thể vượt qua cổng đó — nếu không, repo trả về bản tốt nhất kèm
> danh sách lỗi chưa sửa, không tuyên bố cảm tính.

Repo này đi xa hơn repo mẫu
[`togcoder/skill-for-pptx`](https://github.com/togcoder/skill-for-pptx)
(PPTX Motion Lab) ở 3 điểm:

1. **Cả 3 định dạng, chung một nguồn sự thật:** Word, Excel và PowerPoint cùng
   đọc `design_tokens.json` + `STANDARDS.md` + `rubric.md`. Không copy logic
   animation PPTX sang Word/Excel.
2. **Animation PPTX sâu nhưng trung thực:** animation native, sửa được trong
   PowerPoint, chèn XML `<p:timing>` theo "file vàng", có thang bằng chứng
   rõ ràng (schema → LibreOffice → PowerPoint render → trình chiếu click thật).
   Chưa qua bước cuối thì không được ghi "đã kiểm chứng trong PowerPoint".
3. **Audit không tự khen nhau:** Auditor chạy context riêng, không thấy Builder
   tự đánh giá, bắt buộc trích dẫn giá trị đo được, vị trí slide/sheet/cell,
   ngưỡng và cách sửa. Deterministic gate có quyền phủ quyết mọi lời khen.

## Cấu trúc

```text
design_tokens.json          # Nguồn sự thật duy nhất: màu, font, cỡ chữ, thời lượng animation
STANDARDS.md                # Tiêu chuẩn, hard fail, ngưỡng phát hành
rubric.md                   # Rubric chấm điểm 0-5, trọng số, BLOCKER/MAJOR/MINOR
office_builder/             # Code sinh file thật (không phải prompt suông)
  theme.py                  # Đọc design_tokens.json, tính contrast WCAG
  docx_builder.py           # python-docx, template-first, Styles, TOC field, số trang
  xlsx_builder.py           # openpyxl, Excel Table, formula, chart, print setup
  pptx_builder.py           # python-pptx, template-first, layout/master, chart, bảng
  pptx_motion.py            # Chèn animation native: timing XML, click beat, Morph
  audit.py                  # Deterministic gate cho cả 3 loại file
agents/
  orchestrator.py           # State machine, đếm vòng, giữ bản tốt nhất, decision.log
  researcher.md             # Vai Researcher: evidence.md có nguồn + mức tin cậy
  distiller.md              # Vai Distiller: distilled_spec.md, hợp đồng cho Builder
  builder.md                # Vai Builder: chỉ làm file thật + build_log.md
  trial.md                  # Vai Trial: chạy thử trên dữ liệu mẫu nhỏ, render sớm
  auditor.md                # Vai Auditor: audit_report.json, không chấm cảm tính
  standards_keeper.md       # Vai Standards Keeper: version + changelog, không cứu bản đang audit
  standards_scout.md        # Vai Standards Scout: tìm tiêu chuẩn ứng viên + test tương ứng
skills/                     # Agent Skills (SKILL.md, progressive disclosure)
  office-word/SKILL.md
  office-excel/SKILL.md
  office-pptx/SKILL.md
  office-pptx-motion/SKILL.md
  office-audit/SKILL.md
tools/
  audit.py                  # CLI: python -m tools.audit output/...
  render_check.py           # LibreOffice headless -> PDF/PNG để kiểm tra tràn/mép
examples/
  data.json                 # Dữ liệu mẫu: báo cáo Q3 (đã kiểm chứng tổng)
  build_example.py          # Chạy orchestrator trên dữ liệu mẫu, sinh 3 file
  README.md
tests/
  test_builders_and_audit.py
runs/<task-id>/             # Mỗi lần chạy: task.json, evidence, spec, output, audit_report.json
```

## Chạy thử nhanh

```bash
python -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python examples/build_example.py
.venv/bin/python -m tools.audit examples/output/report.pptx examples/output/report.docx examples/output/report.xlsx
.venv/bin/python -m pytest tests/ -q
```

Ví dụ sẽ sinh `examples/output/report.docx`, `report.xlsx`, `report.pptx`
(tiêu đề khẳng định, 1 ý/slide, bar chart bắt đầu từ 0, bảng Excel dùng
formula liên kết Data → Summary, PPTX có animation Fade theo click beat),
sau đó audit bằng code. Nếu audit FAIL, orchestrator sửa đúng
`must_fix_next` và chạy lại, tối đa 3 vòng mặc định (5 vòng cho tài liệu
phức tạp), dừng sớm khi 2 vòng liên tiếp không cải thiện.

## Vòng làm việc

```text
Researcher -> Distiller -> Spec
  -> Builder -> bản thử v1
  -> Trial (dữ liệu mẫu nhỏ, render sớm)
  -> Auditor (độc lập) + Verifier bằng code
  -> FAIL: Builder sửa đúng lỗi liệt kê
  -> PASS hoặc hết vòng: Standards Keeper xác nhận + Orchestrator chốt
```

**Quality gate (đồng thuận có bằng chứng):**

1. Deterministic gate 100% PASS: mở được, không lỗi OOXML, không tham chiếu
   hỏng, số liệu khớp kiểm tra độc lập.
2. Không BLOCKER, MAJOR ≤ 2 (tài liệu thường), = 0 (tài liệu gửi khách hàng).
3. Điểm rubric ≥ 85/100, không tiêu chí nào ở mức 0.
4. Không thoái lui: bản mới ≥ bản tốt nhất đã lưu.
5. Auditor + Verifier bằng code + Standards Keeper đều PASS trên cùng một
   artifact hash.

## Giới hạn phải ghi rõ (tránh hứa quá)

- `python-docx` không render layout/PDF và không tính TOC/số trang; TOC và
  PAGE chỉ là field, Word/LibreOffice mới tính khi mở.
- `openpyxl` không tính formula và không tạo PivotTable native mới; kiểm tra
  giá trị phải qua LibreOffice recalc, và "bảng tổng hợp" phải ghi rõ không
  phải PivotTable native.
- `python-pptx` không đo text thật như PowerPoint render và không có API
  animation; animation phải viết XML tay theo file vàng và kiểm chứng bằng
  render. LibreOffice không có Morph và chỉ chơi subset effect, nên
  "LibreOffice chạy được" ≠ "PowerPoint chạy được".
- Trigger bấm shape chỉ kiểm chứng được bằng trình chiếu tương tác thật,
  CreateVideo không bấm được.

## Nguồn tham khảo chính

- Repo mẫu PPTX Motion Lab: https://github.com/togcoder/skill-for-pptx
- `python-pptx`: https://python-pptx.readthedocs.io/en/latest/
- `python-docx`: https://python-docx.readthedocs.io/en/latest/
- `openpyxl`: https://openpyxl.readthedocs.io/en/stable/
- WCAG 2.1 AA: https://www.equalweb.com/wcag/2-1/
- Xem đầy đủ trong `research/` (bản nghiên cứu) và từng `skills/*/references/`.

## Giấy phép

MIT. Xem `LICENSE`.
