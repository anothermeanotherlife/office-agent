# Nghiên cứu PPTX cho AI agent
Nguồn: nhóm nghiên cứu PPTX, 2026-10-26. Repo mẫu: https://github.com/togcoder/skill-for-pptx

## skill-for-pptx: cái hay
- PPTX Motion Lab, AI Motion Director, không phải trình tạo slide thông thường.
- 3 đường sản phẩm: Motion Director (giữ nguyên slide, gắn chuyển động native), Morph Studio (Morph bằng tên !!), Slide Forge (từ spec ngữ nghĩa 12 layout, đo bố cục/cỡ chữ thật).
- Một nguồn sự thật duy nhất: knowledge/design_system.json; generator và checker cùng đọc một bộ token.
- Mô hình nhịp click: Slide -> Click Beat -> Stage -> Effects. Click là ranh giới người thuyết trình. Trong một beat, With Previous chạy cùng lúc, After Previous chạy nối tiếp.
- Thang bằng chứng: cấu trúc -> mô phỏng -> LibreOffice -> PowerPoint đọc -> PowerPoint render CreateVideo -> trình chiếu tương tác thật bằng click.
- 198 preset ppt:<name>, 289 test trên nhánh nghiên cứu T034, design_critic.py chấm theo luật kèm giá trị đo, ngưỡng, cách sửa.

## skill-for-pptx: cái thiếu / rủi ro
- Chưa có bằng chứng trình chiếu tương tác nào; README/STATUS thừa nhận một phần đã được PowerPoint 16.0 đọc/render nhưng chưa deck nào trình chiếu click-by-click thật. GIF/storyboard chỉ là mô phỏng.
- main chậm hơn nhánh nghiên cứu: main ở T022/T023, nhánh nghiên cứu tới T034. Đừng lấy main làm chuẩn duy nhất.
- Morph chỉ có bằng chứng Fade fallback ở LibreOffice; LibreOffice không có Morph, cần kiểm chứng riêng trên PowerPoint.
- Trigger bấm shape mới ở mức cấu trúc; CreateVideo không bấm được nên không kiểm chứng trigger bằng render video.
- 3 script bị cách ly vì viết XML sai schema hoặc sai trên deck thật: morph_playbook.py, cognitive_pacing.py, interactive_triggers.py. Không tái sử dụng.
- Thí nghiệm cũ E001-T018 đã đóng băng để tái lập, không xây tiếp.
- Lỗi còn mở: token kết thúc sai vị trí trước Morph (MOT-SETTLE), Forge dùng 13-15 cỡ chữ trong khi chuẩn 3-6 cỡ/deck, critic chưa hiệu chuẩn bằng xếp hạng người thật.
- Repo chỉ làm PPTX, không dạy Word/Excel. Repo đa Office cần skill riêng python-docx và openpyxl, chỉ chia sẻ chung spec, design token và cổng kiểm chứng.
- Trang GitHub hiện ghi Public, 0 star, 81 commit, tạo 2026-10-04; README cũ còn ghi private, kiểm tra lại quyền thực tế.

## python-pptx
- prs.slide_layouts[0..8] theo quy ước theme chuẩn nhưng thứ tự không bắt buộc; đọc tên layout và placeholder thực tế, không hard-code index mù.
- Sửa đồng nhất ở master/layout/template, không sửa lẻ từng slide.
- Cách bền nhất: bắt đầu từ file .pptx template đã thiết kế sẵn trong PowerPoint, rồi thay nội dung bằng python-pptx.
- Theme ở ppt/theme/theme1.xml: a:clrScheme màu, a:fontScheme font. Đổi theme là đổi cả deck; series biểu đồ lấy màu Accent 1-6 theo thứ tự.
- Font: đặt theo run, cả Latin, East Asian, Complex Script nếu có tiếng Việt; kiểm tra font có sẵn trên máy trình chiếu.
- Hình: add_picture, giữ tỉ lệ, dùng Pillow đo kích thước/DPI thật; python-pptx không tự nén ảnh.
- Bảng: add_table trả về GraphicFrame, truy cập qua .table; tạo đúng số hàng/cột ngay từ đầu, tránh chèn/xóa hàng bằng deepcopy XML gây trùng ID.
- Biểu đồ: add_chart với CategoryChartData / XyChartData / BubbleChartData; pie chỉ một series; tô màu từng cột bằng series.points[i].
- Tọa độ EMU: 914400 EMU = 1 inch.
- Text: python-pptx không đo text thật như PowerPoint render; word_wrap và auto_size không đảm bảo không tràn. Giới hạn từ, cỡ chữ, chiều rộng, kiểm chứng bằng render.

## Animation và transition: giới hạn lớn nhất
- python-pptx không có API cao cấp cho animation hay transition; phải tự viết XML <p:timing>.
- Transition là phần tử riêng <p:transition> của slide, ví dụ <p:transition spd="medium"><p:fade/></p:transition>, có spd, dur, advClick, advTm.
- Morph là transition ở slide đích: <p:transition><p:morph type="byObject"/></p:transition>.

## Cấu trúc XML timing
- p:timing > p:tnLst > p:par > p:cTn nodeType="mainSeq" > childTnLst, mỗi click là p:par nodeType="clickEffect" với stCondLst cond delay="indefinite" evt="onClick".
- Effect: animEffect preset (entr/exit, presetID, presetClass, presetSubtype), anim (attrName), animClr, animRot, animScale, animMotion (origin, path, rAng, ptsTypes, dùng #ppt_x, #ppt_y, #ppt_w, #ppt_h), set.
- Target: p:tgtEl > p:spTgt spid="<id shape>" khớp p:cNvPr id, không phải tên hiển thị. Text theo đoạn: txEl > paraRg st/end; theo ký tự: charRg.
- Trigger bấm shape: nằm trong interactiveSeq, điều kiện evt="onClick" trỏ tới shape trigger, không chỉ đổi nodeType.

## Chèn XML an toàn
1. Tạo slide bằng python-pptx, lưu, đọc lại shape ID thực tế, đừng đoán ID.
2. Parse slide XML bằng lxml, đúng namespace p và a.
3. Lấy fragment timing từ file vàng: tạo deck 1 slide trong PowerPoint, gắn đúng effect, lưu, giải nén, copy fragment, chỉ thay spid, dur, delay, preset.
4. Nếu slide đã có <p:timing>, thay toàn bộ, không append timing thứ hai.
5. Đúng thứ tự schema slide: cSld, clrMapOvr (nếu có), transition, timing, extLst.
6. cTn id duy nhất, tăng dần trong từng slide.
7. Không sửa file bằng thay chuỗi thô trong ZIP.
8. Sau khi chèn, mở lại bằng python-pptx, parse lại bằng lxml, rồi render kiểm chứng.
9. Aspose/Slides hoặc Spire có API timeline thật nhưng là thư viện thương mại, không phải python-pptx thuần.

## Mẫu animation tốt
- Entrance: Fade, Fly In, Wipe, Zoom nhẹ, Appear; một ngôn ngữ entrance thống nhất cả deck.
- Emphasis: Pulse, Grow/Shrink nhẹ, đổi màu, highlight số liệu/call-to-action.
- Exit: Fade, Fly Out, Wipe, ngắn hơn entrance, chỉ khi cần dọn sân khấu.
- Motion path: đường thẳng, cong đơn giản, theo trục chính; tránh ngoằn ngoèo, ra khỏi slide không quay lại.
- Thứ tự theo thứ tự đọc và kể chuyện; tiêu đề thường không nên bị animation che; luật "draft không target title" đáng giữ.
- Duration/delay token skill-for-pptx: entrance 250-700ms (mặc định 500, giới hạn 150-1000), exit 200-500 (350), emphasis 250-600 (450), move 400-1200 (700, tăng theo quãng đường travel_ms), Morph 1000-1600 (1400), stagger 60-150 (120), ambient loop 6-60s (12, không đặt lên tiêu điểm), easing PowerPoint chỉ accel/decel (enter 0/0.90, exit 0.35/0, move/emphasis 0.15/0.70).
- Click: tốt ≤6 click/slide, giới hạn ≤8; mỗi click tốt ≤4 vật thể độc lập, ≤8 effect; cả deck tốt 3-10 loại effect, giới hạn 2-14. Mỗi click kết thúc ở trạng thái đáng dừng lại nói thì phải là click riêng.
- Morph: nhân đôi slide, giữ tên !! cho cặp shape, chỉ thay vị trí/kích thước/màu/hình dạng cần morph.
- Chart build: an toàn nhất là animate cả biểu đồ như một shape, hoặc tách nhiều slide/stage; build theo series/điểm khi viết XML tay là mong manh, chỉ làm khi có file vàng và kiểm chứng PowerPoint thật.

## Lỗi thường gặp
- Trùng cTn id, spid không trỏ tới shape nào sau khi lưu lại.
- Hai <p:timing>, timing sai thứ tự schema, thiếu namespace.
- presetID + presetClass + presetSubtype không khớp -> PowerPoint hiện "Custom" hoặc bỏ qua.
- Dùng delay dài thay click; gộp cả slide vào một click đầu.
- Exit trước entrance, animate cùng thuộc tính bằng hai effect cùng lúc.
- Animate group không nhất quán target group/con.
- Helper shape lộ ở frame cuối, tương phản thấp (ví dụ 1.14:1).
- Morph ghép sai vì tên !! khác, z-order đổi, shape thêm/xóa.
- Text tràn vì chỉ đo bằng python-pptx; font không có trên máy trình chiếu; màu RGB hard-code.
- LibreOffice mở được nhưng PowerPoint sửa file, hoặc ngược lại; LibreOffice không có Morph và chỉ chơi subset effect.
- Trigger không chạy vì chỉ sửa mainSeq, không làm đúng interactiveSeq.

## Tiêu chuẩn kiểm chứng bằng code
Cổng cứng (fail là không giao):
1. File ZIP hợp lệ, mọi XML parse được bằng lxml, không repair prompt.
2. Mở lại bằng python-pptx, số slide, thứ tự, nội dung không mất khi chỉ gắn animation.
3. Mọi spid trong timing tồn tại; cTn id duy nhất; tối đa một <p:timing>/slide; transition trước timing đúng schema.
4. Mọi shape cuối cùng trong khổ slide, trừ trạng thái bắt đầu có chủ đích của Fly In đã kiểm chứng.
5. Không effect mồ côi, không shape vừa exit mà không quay lại nhưng vẫn nhắc tới ở slide sau.
6. LibreOffice headless convert sang PDF thành công: soffice --headless --convert-to pdf file.pptx.
7. Nếu có PowerPoint Windows: ghi hash file, phiên bản, OS, font, màn hình; Animation Pane có tên effect thật, không phải "Custom"; render CreateVideo nếu cần.

Dải tốt:
- Contrast 4.5:1, chữ lớn 3:1 (≥18pt, hoặc 14pt bold); đồ họa/helper 3:1.
- Cỡ chữ: body ≥18pt tốt, ≥12pt giới hạn; tỉ lệ title/body 1.5-2.6; số cỡ chữ/deck 3-6 tốt, 2-9 giới hạn.
- Mật độ chữ: ≤45 từ/slide nội dung tốt, ≤80 giới hạn; đếm từ tiếng Việt bằng khoảng trắng sẽ đếm quá tay, dùng token tiếng Việt hoặc đếm ký tự kèm từ.
- Click/mover/effect/duration theo bảng trên.
- Bố cục tĩnh: nền chiếm 60-95% ở frame ổn định tốt, độ lệch cân bằng ≤0.15 tốt.
- Thứ tự timing khớp kể chuyện và thứ tự đọc; mỗi click kết thúc ở trạng thái ổn định, có ý nghĩa.
- Text không tràn: đo bằng render LibreOffice/PowerPoint hoặc Pillow/fontTools, không chỉ tin text_frame.
- Điểm critic là đo lường, không phải bằng chứng playback.

Định nghĩa "đồng ý đây là sản phẩm tốt nhất": không phải mọi agent cùng nói tốt, mà là agent nghiên cứu, lọc, thử, audit, tiêu chuẩn đều độc lập chạy cùng một bộ test, tất cả cổng cứng pass, không còn warning chưa giải thích, flagship deck đạt ít nhất PowerPoint render (rung 5) hoặc tốt hơn là playback tương tác (rung 6), mới đóng gói. Thêm quy tắc dừng: tối đa số vòng sửa, đóng băng spec, ghi changelog; nếu không vòng "làm tới khi hoàn hảo" sẽ lặp vô hạn.

## Nguồn
- https://github.com/togcoder/skill-for-pptx
- https://github.com/togcoder/skill-for-pptx/blob/main/docs/STATUS.md
- https://github.com/togcoder/skill-for-pptx/blob/main/docs/DESIGN_SYSTEM.md
- https://github.com/togcoder/skill-for-pptx/blob/main/docs/VERIFICATION.md
- https://github.com/togcoder/skill-for-pptx/blob/main/docs/CLICK_BEAT_CHOREOGRAPHY.md
- https://python-pptx.readthedocs.io/en/latest/user/slides.html
- https://python-pptx.readthedocs.io/en/latest/user/charts.html
- https://python-pptx.readthedocs.io/en/latest/user/table.html
- https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
- https://support.microsoft.com/en-us/office/set-the-start-time-and-speed-of-an-animation-effect-bf8c1cb4-c827-48b6-b756-8c1a3e681a60
