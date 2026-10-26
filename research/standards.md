# Nghiên cứu tiêu chuẩn cho repo Word / Excel / PowerPoint
Nguồn: nhóm nghiên cứu tiêu chuẩn thiết kế, 2026-10-26. Dùng làm đầu vào cho STANDARDS.md và design_tokens.json.

## Nguyên tắc
- Một slide / một trang = một ý; tiêu đề dạng khẳng định, nêu kết luận.
- Slide tối đa ~6 dòng, tối đa 6 thẻ / 8 KPI; vượt thì tách hoặc đưa sang phụ lục.
- Thứ tự đọc: tiêu đề → con số/kết luận → bằng chứng → nguồn/chú thích.
- Dùng slide master / layout có sẵn, không tạo slide trắng rồi chèn textbox rời rạc.
- Whitespace 20-30%; slide phủ >85% là chật, <15% là trống rỗng (tham khảo pptx-mcp).
- Ngưỡng máy kiểm: không tràn lề, không chồng lấn >0,05 inch, khoảng cách giữa các khối ≥0,1 inch.
- Tối đa 2 font, tối đa 3-4 màu; mọi màu trong theme phải có tên và HEX trong token file.
- Word: Styles Heading 1/2/3, bìa, mục lục tự động, đánh số trang, header/footer.
- Excel: 1 bảng dữ liệu / 1 vùng, chuyển thành Excel Table, không dùng bảng để trình bày layout.
- PowerPoint: khổ 16:9 mặc định, theme/master duy nhất, mọi slide đều có tiêu đề.

## Trực quan hoá dữ liệu
- Bar chart phải bắt đầu trục từ 0; không chart 3D, không dual-axis 2 thang độc lập.
- Một chart chỉ 1 thông điệp; tiêu đề chart nêu kết luận, nhãn trực tiếp.
- Pie 3-7 lát, bar 5-15 nhóm; nhiều hơn thì gộp thành "Khác" hoặc tách chart.
- Mỗi chart/bảng phải có: nguồn, ngày/kỳ dữ liệu, đơn vị, mẫu/n nếu có, định nghĩa chỉ số.
- Làm tròn nhất quán: % 1 chữ số thập phân; ghi chú khi tổng ≠ 100% do làm tròn.
- Excel: số liệu trình bày phải là công thức liên kết tới sheet dữ liệu, không gõ tay.

## Accessibility WCAG 2.1 AA
- Tương phản chữ thường 4,5:1, chữ lớn 3:1 (≥18pt thường hoặc ≥14pt đậm).
- Không dùng màu làm kênh duy nhất; thêm nhãn, icon, ký hiệu.
- PowerPoint: body tối thiểu 18pt, tiêu đề ≥28-32pt; line spacing 1,15-1,5.
- Word: body 11-12pt, line spacing 1,15-1,5.
- Alt text ≤125 ký tự, nêu nội dung và mục đích; ảnh trang trí đánh dấu decorative.
- Mỗi slide có tiêu đề duy nhất 5-10 từ; slide tiếp nối ghi "1 of 2".
- Bảng: có header row, không merge/split/nested cell trong vùng dữ liệu.
- Metadata bắt buộc: Title, Author, Language, ngày, phiên bản. Tên file/sheet/slide có nghĩa.
- Excel tách rõ README, Data, Calc, Dashboard và ghi nguồn ở README.

## Animation PPTX
- Dùng animation native, sửa được trong PowerPoint, không render video thay thế.
- 3 lớp: chính, phụ (highlight), nền (chậm, không tranh chú ý).
- Giới hạn: ngắn, không nhấp nháy, không tự chuyển slide theo thời gian; ≤4 vật thể chuyển động cùng lúc; slow-in/slow-out, so le có chủ đích.
- Thang bằng chứng: cấu trúc → mô phỏng → LibreOffice đọc → PowerPoint đọc/render → trình chiếu click thật.

## Checklist chấm điểm 0-5 và luật đồng thuận
- 0 thiếu hẳn/lỗi vỡ, 1 sai nhiều, 2 một phần đúng, 3 đạt tối thiểu, 4 tốt, 5 mẫu mực.
- 9 hạng mục: cấu trúc, layout, typography/màu, trực quan hoá, độ chính xác số liệu, bảng/Excel, accessibility, nhất quán/metadata, motion (PPTX).
- Hard fail: trục bar không từ 0, dual-axis gây lệch, 3D, contrast <4,5:1/3:1, thiếu alt text chart/ảnh chính, không nguồn, tổng sai, merge header, tràn lề, chồng lấn, slide phủ >85%.
- Ngưỡng phát hành: không hạng mục nào <3, trung bình ≥4, hạng mục 4, 5, 7 phải ≥4. Bất kỳ hard fail nào cũng phải sửa và chạy lại audit.
- Audit phải ghi rõ: luật vi phạm, giá trị đo được, ngưỡng, vị trí slide/sheet/cell và cách sửa.

## Nguồn tham khảo chính
- https://deckary.com/blog/pillar-powerpoint-design-guide
- https://www.slidor.agency/blog/professional-powerpoint-design-tips-presentation-experts
- https://www.infocepts.ai/glossary/data-visualization/
- https://useeasytool.com/blog/articles/data-visualization-best-practices.html
- https://www.equalweb.com/wcag/2-1/
- https://www.aims.edu/accessibility-learning-hub/accessibility-checklists/microsoft-powerpoint
- https://pypi.org/project/pptx-mcp/0.8.9/
- https://github.com/togcoder/skill-for-pptx

Gợi ý: lưu các ngưỡng thành design_tokens.json để mọi agent đọc cùng một file, đồng thuận theo số đo.
