# Visual Auditor — người (hoặc agent) phải MỞ file ra xem

Bài học ngày 2026-10-26: audit bằng code đã cho 100/100/100, nhưng khi mở
file render ra xem, file vẫn xấu và có lỗi thật. Vì vậy:

## Audit code KHÔNG phải là xem file

`office_builder/audit.py` chỉ mở file bằng python-pptx/python-docx/openpyxl
và đo: có title không, có tràn không, trục có từ 0 không, Summary có phải
công thức không. Nó không đọc được:

- Chú thích có đúng với chart không (slide 2 từng ghi "cột xám là chi phí"
  trong khi chart không có series chi phí nào).
- Chart có legend không, legend có rối không, data label có chồng lên nhau không.
- Bảng Word có bị tách trang mà mất header không, mục lục field có render trống không.
- Sheet Excel in ra có bị cắt trang, tiêu đề bị cụt, chart legend rối không.
- Slide có nửa dưới trống trơn không, khoảng trắng có phải thiết kế hay là bỏ trống.

`tools/render_check.py` cũ còn tệ hơn: nó chỉ kiểm tra LibreOffice convert
trả về exit 0, tức là "chuyển được sang PDF", không phải "PDF đẹp".

## Quy trình Visual QA bắt buộc trước khi giao file

1. Convert từng file riêng (không dùng chung tên report.pdf, sẽ ghi đè nhau)
   sang PDF bằng LibreOffice headless, rồi sang PNG bằng pdftoppm.
2. Mở từng ảnh PNG ra xem, slide/trang một, và ghi lại từng lỗi nhìn thấy:
   legend, annotation vs data, label chồng, bảng tách trang, in cắt trang,
   khoảng trắng, tiêu đề bị cụt.
3. Sửa lỗi đó trong builder, không sửa trong file đã xuất.
4. Audit code lại + render lại + xem lại ảnh mới. Chỉ giao khi cả hai vòng
   đều qua: code PASS và mắt nhìn không còn lỗi chặn.

Một file chỉ được gọi là "đã xem" khi có bước 2 trong log — xem ảnh render,
không phải xem log audit. Nếu chưa xem ảnh, phải nói rõ "chưa xem bằng mắt",
không được nói "đã kiểm tra kỹ".
