# Standards v1.0

Nguồn sự thật duy nhất cho cái đẹp và kỹ thuật là `design_tokens.json`.
Generator và Auditor **cùng đọc một bộ số đó**. Nếu hai bên dùng hai bộ
luật khác nhau, audit vô nghĩa.

Changelog: xem cuối file. Standards Keeper chỉ cập nhật khi có lỗi lặp
lại ≥ 2 lần, kèm test tương ứng. Không sửa tiêu chuẩn trong lúc đang audit
một bản cụ thể để cứu bản đó.

## 1. Nguyên tắc chung

- Một slide / một trang = một ý. Tiêu đề dạng khẳng định, nêu kết luận,
  không nêu chủ đề chung chung. Ví dụ: "Doanh thu Q3 tăng nhờ kênh online"
  thay vì "Doanh thu".
- Thứ tự đọc: tiêu đề → con số/kết luận → bằng chứng → nguồn/chú thích.
- Tối đa 2 font, tối đa 3–4 màu. Mọi màu phải có tên và HEX trong token
  file; cấm màu tự chọn lẻ ngoài palette.
- Dùng style/master/layout có sẵn, không format thủ công rời rạc.
- Metadata bắt buộc: Title, Author, Language, ngày, phiên bản. Tên
  file/sheet/slide có nghĩa, không để `Document1`, `Sheet1`, `Slide 1`.

## 2. Word

- Dùng Styles Heading 1/2/3 thay vì bôi đậm thủ công; có bìa, mục lục tự
  động từ Heading (field TOC), đánh số trang (field PAGE/NUMPAGES),
  header/footer, chú thích bảng/biểu đồ, phụ lục ghi rõ A, B.
- Header/footer: phải set `is_linked_to_previous=False` trước khi sửa
  section mới, nếu không sẽ sửa luôn section trước.
- Bảng: tổng chiều rộng cột ≤ khổ giấy trừ margin, header lặp lại
  (`tblHeader`), không merge trong vùng dữ liệu.
- Ảnh: tồn tại trong `word/media/`, có relationship, không vượt lề, có
  alt text (`descr`) nếu là ảnh có ý nghĩa, caption dạng `Hình X. ...`.
- TOC và PAGE chỉ là field; Word/LibreOffice mới tính khi mở. Audit chỉ
  kiểm được "có field + có đủ Heading", không kiểm được số trang TOC nếu
  chưa render — và phải ghi chú rõ điều này cho người dùng.

## 3. Excel

- 1 bảng dữ liệu / 1 vùng, chuyển thành Excel Table (header, filter, vùng
  động). Không dùng bảng để trình bày layout.
- Tách rõ `README`, `Data`, `Summary`, `Dashboard`; ghi nguồn ở `README`.
- Số liệu trình bày phải là công thức liên kết tới sheet dữ liệu, không gõ
  tay. Tổng phải đối chiếu được với nguồn.
- Kiểu dữ liệu đúng: số là số, ngày là datetime. Lưu số thuần, format code
  viết theo chuẩn Excel en-US (`#,##0.00`, `yyyy-mm-dd`); hiển thị theo
  locale là việc của Excel khi mở.
- `openpyxl` không tạo PivotTable native mới; "bảng tổng hợp" phải ghi rõ
  không phải PivotTable native, hoặc dùng template đã có pivot + refresh.
- Freeze panes, auto filter, print title rows, print area, fitToPage,
  orientation/paper size là bắt buộc cho file gửi khách.
- Conditional formatting tối đa 2–3 rule/sheet; validation chỉ cảnh báo khi
  nhập tay trong Excel, audit vẫn phải kiểm dữ liệu đầu vào.

## 4. PowerPoint

- Khổ 16:9 mặc định; theme/master duy nhất; mọi slide đều có tiêu đề.
- Dùng slide master / layout có sẵn, không tạo slide trắng rồi chèn textbox
  rời rạc — vì phá reading order cho screen reader.
- Lưới và căn lề nhất quán; căn trái cho khối văn bản dài.
- Whitespace 20–30%; slide phủ >85% là chật, <15% là trống rỗng.
- Ngưỡng máy kiểm: không tràn lề, không chồng lấn >0,05 inch, khoảng cách
  giữa các khối ≥0,1 inch.
- Body tối thiểu 18pt, tiêu đề ≥28–32pt; số cỡ chữ/deck tốt 3–6, giới hạn 2–9.
- Mật độ chữ: tốt ≤45 từ/slide nội dung, giới hạn ≤80. Đếm từ tiếng Việt
  bằng khoảng trắng sẽ đếm quá tay, nên dùng token tiếng Việt hoặc đếm ký
  tự kèm từ khi cần.

## 5. Trực quan hoá dữ liệu

- Ma trận chọn chart: so sánh → bar; xu hướng → line; cơ cấu → stacked
  bar/treemap, pie chỉ khi ít nhóm; phân phối → histogram/box; quan hệ →
  scatter/bubble; cần số chính xác → bảng.
- **Hard fail:** bar chart không bắt đầu trục từ 0; chart 3D; dual-axis với
  2 thang độc lập; chọn cửa sổ thời gian thiên vị; gộp nhóm che giấu xu
  hướng ngược bên trong.
- Một chart chỉ 1 thông điệp; tiêu đề chart nêu kết luận, nhãn trực tiếp
  thay vì chỉ dựa vào legend.
- Pie 3–7 lát, bar 5–15 nhóm; nhiều hơn thì gộp thành "Khác" hoặc tách chart.
- Mỗi chart/bảng phải có: nguồn, ngày/kỳ dữ liệu, đơn vị, mẫu/n nếu có,
  định nghĩa chỉ số. Ghi chú làm tròn khi tổng ≠ 100% do làm tròn.
- Làm tròn nhất quán: % 1 chữ số thập phân, tiền tệ không lẻ không cần
  thiết; không trộn 12,5% với 13% trong cùng báo cáo.

## 6. Animation (chỉ PPTX)

- Dùng animation native, sửa được trong PowerPoint, không render video thay thế.
- 3 lớp: chính (hiện dần, build chart, zoom), phụ (highlight theo tiêu
  điểm), nền (chuyển động chậm, không tranh chú ý).
- Mô hình nhịp click: `Slide → Click Beat → Stage → Effects`. Click là
  ranh giới người thuyết trình, không dùng delay dài thay click. Trong một
  beat, `With Previous` chạy cùng lúc, `After Previous` chạy nối tiếp.
- Giới hạn: animation ngắn, không nhấp nháy, không tự chuyển slide theo
  thời gian; ≤4 vật thể chuyển động cùng lúc; slow-in/slow-out, so le có
  chủ đích. Tốt ≤6 click/slide, giới hạn ≤8; mỗi click ≤8 effect.
- Thời lượng theo `design_tokens.json` (entrance 250–700ms, exit 200–500ms,
  emphasis 250–600ms, move 400–1200ms, Morph 1000–1600ms, stagger 60–150ms).
  Easing PowerPoint chỉ có accel/decel, không có cubic-bezier tự do.
- Morph: nhân đôi slide, giữ tên `!!` cho cặp shape, chỉ thay vị trí/kích
  thước/màu/hình dạng cần morph. Morph là transition ở slide đích.
- Trigger bấm shape: nằm trong `interactiveSeq`, điều kiện `evt="onClick"`
  trỏ tới shape trigger. CreateVideo không bấm được, nên trigger chỉ kiểm
  chứng được bằng trình chiếu tương tác thật.
- Thang bằng chứng: cấu trúc → mô phỏng → LibreOffice đọc → PowerPoint
  đọc/render → trình chiếu click thật. Chưa qua bước cuối thì không được
  ghi "đã kiểm chứng trong PowerPoint".

## 7. Accessibility (WCAG 2.1 AA)

- Tương phản chữ thường 4,5:1, chữ lớn 3:1 (≥18pt thường hoặc ≥14pt đậm);
  thành phần đồ họa có ý nghĩa ≥3:1.
- Không dùng màu làm kênh duy nhất: thêm nhãn, icon, ký hiệu hoặc pattern.
- Chữ trên ảnh/gradient phải đạt tương phản ở điểm nền xấu nhất, hoặc thêm
  lớp nền mờ đặc.
- Tránh ALL CAPS, italic, gạch chân cho khối văn bản dài; gạch chân chỉ
  dành cho link.
- Alt text ≤125 ký tự, nêu nội dung và mục đích; ảnh trang trí đánh dấu
  decorative; không viết "image of…".
- Mỗi slide có tiêu đề duy nhất 5–10 từ; slide tiếp nối ghi "1 of 2".
- Reading order kiểm tra bằng Selection Pane / Reading Order Pane.
- Bảng: có header row, không merge/split/nested cell, không ô trống trong
  vùng dữ liệu; tránh dùng bảng để dàn layout.
- Link có văn bản mô tả đích đến, không dùng "click here" hay URL trần.

## 8. Hard fail và ngưỡng phát hành

Hard fail (bất kỳ cái nào cũng phải sửa và chạy lại audit):

- File hỏng, XML không parse được, repair prompt khi mở.
- Số liệu sai, tổng không khớp nguồn, số gõ tay trong Excel thay vì formula.
- Bar chart không từ 0, dual-axis gây lệch, chart 3D.
- Contrast <4,5:1 / 3:1, thiếu alt text cho chart/ảnh chính.
- Tràn lề, chồng lấn, slide phủ >85%.
- Animation trỏ tới shape không tồn tại, trùng `cTn id`, hai `<p:timing>`
  trong một slide, timing sai thứ tự schema.
- Placeholder, `Lorem ipsum`, `[TODO]`, ảnh/tham chiếu thiếu.

Ngưỡng phát hành: không hạng mục nào <3/5, trung bình ≥4/5, hạng mục
trực quan hoá, độ chính xác số liệu và accessibility phải ≥4/5. Theo
rubric 100 điểm: ≥85 điểm, không tiêu chí nào 0, không BLOCKER,
MAJOR ≤2 (tài liệu thường) hoặc =0 (tài liệu gửi khách hàng).

## 9. Phản biện liên tục

- "Xong" chỉ là checkpoint. Mỗi lần chạy, Critic phải tấn công lại bản
  tốt nhất, kể cả khi bản đó đã PASS.
- Critic bắt buộc pre-mortem: ít nhất 3 cách bản này có thể sai, trước
  khi chấm. Defender bảo vệ bằng bằng chứng đo được, không bằng lời khen.
- Builder chỉ sửa đúng `must_fix_next`. Bản mới phải ≥ bản tốt nhất,
  nếu không giữ bản cũ.
- Dừng tạm thời khi 2 vòng liên tiếp không cải thiện điểm và không còn
  BLOCKER/MAJOR. Dừng tạm thời ≠ đóng dự án: backlog mở phải được mở
  lại ở lần chạy sau.
- Mục không đóng được bằng code (ví dụ playback PowerPoint thật, hiệu
  chuẩn critic bằng xếp hạng người thật) phải nằm trong backlog mở, ghi
  rõ mức bằng chứng đã đạt, và không được ghi là "đã kiểm chứng" khi
  chưa đạt.

## Changelog

- v1.0 (2026-10-26): tổng hợp từ nghiên cứu PPTX (skill-for-pptx),
  DOCX/XLSX (python-docx, openpyxl), kiến trúc đa agent và tiêu chuẩn
  thiết kế/accessibility. Mọi ngưỡng số đều có test tương ứng trong
  `office_builder/audit.py`.
