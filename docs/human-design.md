# Làm file trông "người làm, chuyên nghiệp" thay vì "mùi AI"

Mức bằng chứng: các nguồn dưới đây được chắt lọc từ mô tả video, transcript/tóm tắt
đi kèm và bài viết kèm theo, không phải từ việc xem trọn từng video. Dùng chúng làm
nguyên tắc thiết kế, không trích nguyên văn từng phút video.

## Nguồn tham khảo chính

- Nancy Duarte — "The Secret Structure of Great Talks" (TED):
  https://www.ted.com/talks/nancy_duarte_the_secret_structure_of_great_talks
  Bài nói hay dao động giữa "what is" và "what could be", kết bằng hành động.
- Duarte — contrast trong thuyết trình: https://www.duarte.com/blog/ultimate-guide-to-contrast/
- Cole Nussbaumer Knaflic — "Storytelling with Data", Talks at Google:
  https://www.youtube.com/watch?v=8EMW7io4rSI — đừng chỉ "show data", hãy kể chuyện
  bằng data, dùng màu dẫn mắt tới đúng một điểm mỗi lần, annotation trực tiếp lên chart.
- Edward Tufte — data-ink ratio, chartjunk, small multiples, graphical integrity:
  https://www.graphomate.com/en/2022/08/my-heroes-of-dataviz-4-5-edward-tufte-the-information-design-guru/
  Tối đa mực cho data, bỏ 3D, lưới dày, màu trang trí; trục cắt làm sai lệch tỷ lệ.
- Barbara Minto — Pyramid Principle (tóm tắt): https://deckary.com/blog/pillar-consulting-presentations-guide
  Kết luận trước, rồi 3–4 luận điểm MECE, rồi mới bằng chứng. Đọc riêng tiêu đề cũng phải thấy lập luận.
- PowerPoint Masterclass / Morph / masking (One Skill, BuckPPT, Akshrika):
  https://www.classcentral.com/course/youtube-if-i-wanted-to-master-powerpoint-in-2025-this-is-what-i-d-watch-full-masterclass-421091
  https://www.youtube.com/watch?v=sXoacTnzLPc — kỹ thuật đỉnh không phải nhiều hiệu ứng,
  mà là Morph có đặt tên `!!`, mặt nạ, nhịp click có chủ ý.
- Financial modelling — Wall Street Prep, "Modeling Best Practices", chuẩn FAST:
  https://wallstreetprep.com/learn-financial-modeling/wp-content/uploads/2024/05/Modeling-Best-Practices_2024_FINAL.pdf
  Xanh = input gõ tay, đen = công thức, xanh lá = link sheet khác; không nhét số vào công thức.
- Editorial design — grid, hierarchy, typography, white space:
  https://blog.joomag.com/transform-magazine-design-and-layout-master-modern-eye-catching-publications

## 10 nguyên tắc áp dụng trong ví dụ khó

1. Một slide / một trang / một bảng = một kết luận viết thành câu, có số.
2. Hierarchy 3 tầng: kết luận lớn nhất, bằng chứng chính, chi tiết/nguồn nhỏ và xám.
3. Khoảng trắng là thiết kế, không phải chỗ trống chưa lấp.
4. Grid + alignment thật, cùng lề, cùng khoảng cách dọc.
5. Số liệu kèm 4 thứ: đơn vị, kỳ, nguồn, cách tính.
6. Chart có annotation trực tiếp, làm mờ phần phụ, tô màu đúng điểm cần nói.
7. Màu ít và có lý do: 1 màu chính + 1 màu nhấn + xám; màu nhấn chỉ cho điểm cần chú ý.
8. Tối đa 2 font, 3–4 cỡ chữ có hệ thống xuyên suốt.
9. Tránh bullet tràn: bullet chỉ khi liệt kê song song thật sự; còn lại thành câu + bằng chứng.
10. Cùng hệ thống nhưng layout đổi theo nội dung: slide số liệu thì chart lớn, slide quyết định thì chỉ 1 khuyến nghị + điều kiện.

## 5 dấu hiệu "mùi AI" và cách sửa

1. Mọi slide giống nhau, cùng khung, cùng icon stock → giữ hệ thống nhưng đổi layout theo nội dung, bỏ icon không mang nghĩa.
2. Tiêu đề là chủ đề ("Tổng quan") + 7 bullet → viết tiêu đề thành câu kết luận có số, cắt còn ≤ 4 ý ngắn, phần còn lại xuống notes/phụ lục.
3. Chart mặc định: cầu vồng, 3D, trục cắt, không nhãn, không nguồn → một chart một câu, trục từ 0, data label + annotation + dòng nguồn ngay dưới chart.
4. Số gõ tay, không công thức, không kiểm chứng → mọi tổng là `=SUM`/`SUMIFS`, freeze panes, print setup, recalc bằng LibreOffice.
5. Trang trí thừa: gradient, emoji, shadow, animation mọi vật thể → bỏ thứ không giúp hiểu nhanh hơn; animation chỉ dẫn mắt theo thứ tự nói.

## Ví dụ khó áp dụng các nguyên tắc này

`examples/build_hard_example.py` — báo cáo 4 quý 2026, doanh thu tăng nhưng lợi nhuận giảm.
Nó cố tình khó ở chỗ: kết luận không hiển nhiên (kênh lãi nhất lại là kênh đang đắt dần),
kịch bản thắng chỉ hơn kịch bản nhì 2 triệu (1.490 so với 1.488), nên đề xuất phải đi kèm
điều kiện dừng. Critic bằng code không chấm được những điểm này — chúng được ghi rõ trong
file, trong phụ lục giới hạn, và trong backlog mở, thay vì giấu đi cho đẹp.

## Backlog mở sau ví dụ khó (không tự đóng bằng code)

- Chưa trình chiếu click thật trong PowerPoint; animation mới đạt cấu trúc + LibreOffice render.
- Chưa có xếp hạng của người thật về độ rõ ràng và độ tin cậy của deck.
- Slide quyết định và phụ lục vẫn dày chữ (gần giới hạn 80 từ/slide) — lần cải tiến sau nên tách phụ lục.
- Dữ liệu là tổng hợp nội bộ, không tách chi phí cố định / biến đổi, không có dữ liệu thị trường ngoài.
