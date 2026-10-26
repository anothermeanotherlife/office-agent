# Debate — phản biện liên tục, không phải "làm xong thôi"

"Xong" ở đây chỉ là một checkpoint, không phải trạng thái cuối.
Mỗi lần chạy, Critic phải tấn công lại bản tốt nhất hiện tại, kể cả khi
bản đó đã PASS cổng chất lượng. Nếu Critic không tìm ra lỗi, Critic
chưa làm đủ việc — không phải sản phẩm đã hoàn hảo.

## Các vai trong một vòng phản biện

- **Critic / Troublemaker:** nhiệm vụ là tìm lỗi, được "thưởng" khi tìm
  ra BLOCKER/MAJOR thật, không khi khen. Bắt buộc làm pre-mortem: liệt
  kê ít nhất 3 cách bản này có thể sai trước khi chấm.
- **Defender:** bảo vệ phần đã PASS bằng bằng chứng đo được, không bằng
  lời khen. Nếu Defender không chứng minh được, phần đó bị coi là chưa
  chắc chắn và phải sửa hoặc ghi vào backlog mở.
- **Auditor:** độc lập, không thấy Builder/Critic tự đánh giá trước đó.
  Auditor chỉ nhận artifact + rubric + bằng chứng đo được, rồi quyết
  định PASS/FAIL theo cổng trong `STANDARDS.md`.
- **Standards Keeper:** ghi lại kết quả phản biện vào changelog. Một
  tranh luận chỉ được đóng khi có bằng chứng, không khi mọi người mệt
  và đồng ý cho qua.

## Luật phản biện

1. Không đồng thuận bằng lời khen. Đồng thuận = Auditor + Verifier bằng
   code + Standards Keeper đều PASS trên cùng một artifact hash.
2. Mỗi vòng, Critic phải đưa ra ít nhất một phản biện mới, kể cả khi
   vòng trước đã PASS. Phản biện mới có thể là: thiếu data label, thiếu
   speaker notes, thiếu alt text, chưa kiểm chứng playback, mật độ chữ
   ở giới hạn, chart chưa có nhãn trực tiếp.
3. Builder chỉ sửa đúng `must_fix_next`. Defender kiểm tra thoái lui:
   bản mới phải ≥ bản tốt nhất, nếu không giữ bản cũ.
4. Vòng dừng tạm thời khi: 2 vòng liên tiếp không cải thiện điểm VÀ
   không còn BLOCKER/MAJOR. Dừng tạm thời ≠ đóng dự án: backlog mở vẫn
   còn, và lần chạy sau phải mở lại backlog đó trước.
5. Có những mục không bao giờ tự đóng bằng code, ví dụ: "chưa trình
   chiếu click thật trong PowerPoint", "chưa hiệu chuẩn critic bằng xếp
   hạng người thật". Những mục này phải nằm trong backlog mở, ghi rõ
   mức bằng chứng đã đạt, và không được ghi là "đã kiểm chứng" khi chưa
   đạt.

## Định dạng debate_log.md

Mỗi vòng ghi: Critic nêu gì (luật, vị trí, giá trị đo), Defender phản
biện gì (bằng chứng), Auditor quyết định gì (PASS/FAIL, điểm, lý do),
và backlog còn mở sau vòng đó. Xem `agents/continuous.py` để biết cách
sinh log này tự động từ audit thật.
