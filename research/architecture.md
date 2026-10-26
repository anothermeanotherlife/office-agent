# Nghiên cứu kiến trúc đa agent + vòng audit
Nguồn: nhóm nghiên cứu kiến trúc, 2026-10-26.

## Nguyên tắc
- Orchestrator bằng code, dạng Supervisor / state machine, không chat tự do.
- Luồng: Orchestrator -> Researcher (song song) -> Distiller -> Spec -> Builder -> bản thử v1 -> Auditor độc lập + Verifier bằng code -> FAIL sửa đúng lỗi -> PASS hoặc hết vòng -> Standards Keeper xác nhận -> chốt.

## Các vai
- Orchestrator: quyết định chuyển trạng thái, đếm vòng, giữ bản tốt nhất. Không tự sửa tiêu chuẩn để cho qua.
- Researcher: evidence.md, mỗi khẳng định kèm URL, trích dẫn ngắn, ngày kiểm tra, mức tin cậy.
- Distiller: distilled_spec.md, chỉ giữ quy tắc áp dụng được, ghi xung đột nguồn và cách giải quyết.
- Builder/Executor: file thật .docx/.xlsx/.pptx + build_log.md. Không tự chấm điểm sản phẩm của mình.
- Auditor: audit_report.json, không nhận lịch sử suy luận của Builder. Chỉ đọc code sinh file rồi khen là không đủ.
- Standards Keeper: standards.md có version, changelog, ví dụ tốt/xấu. Không sửa tiêu chuẩn trong lúc audit một bản cụ thể để cứu bản đó.
- Standards Scout: danh sách tiêu chuẩn ứng viên + nguồn + ví dụ kiểm chứng được; chỉ đưa vào repo khi có test tương ứng.

## Artifact mỗi lần chạy runs/<task-id>/
task.json, evidence.md, distilled_spec.md, standards.md (bản sao áp dụng), output/, audit_report.json, decision.log.
audit_report.json tối thiểu: artifact, artifact_sha256, verdict, score, blockers [{rule, where, measured, evidence}], majors, minors, regression_vs_best, must_fix_next.
Quy tắc vàng: Builder chỉ sửa đúng must_fix_next; không sửa phần đã PASS nếu không liên quan.

## Quality gate - đồng thuận có bằng chứng
Bản chỉ được gọi là tốt khi đồng thời:
1. Deterministic gate 100% PASS: mở được, không lỗi OOXML, không tham chiếu hỏng, số liệu khớp kiểm tra độc lập.
2. Không BLOCKER, MAJOR ≤ ngưỡng (≤2 thường, =0 gửi khách hàng).
3. Điểm rubric ≥ ngưỡng, ví dụ ≥85/100, không tiêu chí nào ở mức 0.
4. Không thoái lui: bản mới ≥ bản tốt nhất đã lưu.
5. Auditor độc lập xác nhận, không dùng chung context với Builder.
Đồng thuận = Auditor + Verifier bằng code + Standards Keeper đều PASS trên cùng một artifact hash. LLM khen nhau mà check code FAIL thì vẫn FAIL.

## Giới hạn vòng
- Mặc định tối đa 3 vòng sửa; phức tạp tối đa 5 vòng.
- Dừng sớm nếu 2 vòng liên tiếp không cải thiện điểm, hoặc cùng BLOCKER lặp lại 2 lần.
- Hết vòng mà vẫn FAIL: trả về bản tốt nhất + lỗi chưa sửa + phương án đơn giản hoá, không tuyên bố "tốt nhất" chỉ vì hết vòng.

## Audit file thật
- Mở lại file bằng python-pptx / python-docx / openpyxl ngay sau khi sinh.
- Kiểm tra OOXML bằng openxml-audit (pytest assert PPTX/DOCX/XLSX).
- Không còn placeholder, Lorem ipsum, [TODO], ảnh/tham chiếu thiếu.
- So sánh nội dung trích xuất từ file với dữ liệu nguồn, không so với prompt của Builder.
- PPTX animation: kiểm tra XML thật: slide, shape, left+width ≤ slide_width, top+height ≤ slide_height, timing tree có target shape ID tồn tại, dur, trigger/click hợp lệ, số click/slide, số animation đồng thời ≤4, Morph tên !! khớp chuỗi slide.
- Render thang bằng chứng: schema -> LibreOffice mở -> LibreOffice headless convert PDF/PNG -> PowerPoint render -> trình chiếu click thật. Chưa qua bước cuối thì không ghi "đã kiểm chứng trong PowerPoint".
- Word: thứ tự heading, mục lục, bảng, chú thích, số trang sau convert PDF; độ rộng cột, ô trống, tiêu đề bảng lặp.
- Excel: sheet, vùng dữ liệu, kiểu dữ liệu, công thức, named range, thiết lập in; tính lại độc lập bằng Python, không đọc cached value; dùng LibreOffice recalc để phát hiện công thức lỗi; biểu đồ nguồn tồn tại, trục/nhãn không rỗng.

## Lưu tiêu chuẩn
- standards.md là nguồn duy nhất của sự thật, có version và changelog.
- Rubric trọng số ví dụ: dữ liệu 30, đầy đủ yêu cầu 20, bố cục/render 20, thiết kế/animation 15, accessibility + toàn vẹn file 15. Mỗi trục định nghĩa BLOCKER/MAJOR/MINOR.
- Mỗi luật quan trọng kèm: tốt, xấu, test trong tools/audit.py. Tiêu chuẩn không có test thì chỉ là lời khuyên, không làm gate.
- Skill theo Agent Skills / SKILL.md, progressive disclosure: skills/office-word, office-excel, office-pptx, office-pptx-motion, office-audit; standards/; scripts/audit.py, render_check.py; templates/.
- Builder chỉ load skill loại tài liệu đang làm; Auditor load rubric + script, không load ý đồ thiết kế của Builder.

## Rủi ro và giảm thiểu
- Tự khen nhau: Auditor context riêng, ẩn danh khi cần, pre-mortem liệt kê ít nhất 3 cách bản có thể sai, có critic chuyên tìm lỗi, deterministic gate phủ quyết.
- Audit nông: Auditor không được PASS nếu chưa chạy script và chưa trích dẫn bằng chứng đo được; tách audit code + LLM.
- Tốn token / lặp vô hạn: research song song giới hạn nguồn, Distiller nén spec, model nhỏ cho routine, audit full chỉ khi cần, cache tiêu chuẩn/template/render.
- Thoái lui: lưu mọi phiên bản, giữ bản tốt nhất, so sánh diff, dừng khi 2 vòng không cải thiện.

## Nguồn chính
- https://github.com/togcoder/skill-for-pptx
- https://pypi.org/project/openxml-audit/0.7.0/
- https://arxiv.org/pdf/2605.02801
- https://arxiv.org/html/2509.23055v1
- https://www.alphaxiv.org/abs/2305.11738v4
