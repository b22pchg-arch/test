Bản V45.30: sửa treo khi dùng nhanh bộ câu hỏi và giữ các cập nhật cấu hình liên thiết bị.

Thay đổi chính:
- Bỏ chữ "Tiến độ:" và "Thời gian:" khỏi thanh dưới khi làm bài/ôn tập để tránh đè thông tin trên màn hình nhỏ.
- Mục hướng dẫn trong HTML đã bổ sung phần "Lưu cấu hình và chuyển sang thiết bị khác".
- Cấu hình liên thiết bị lưu bộ đang dùng, số câu theo từng bộ, thời gian thi/ôn, kiểu giải thích, tự gộp theo tỷ lệ và tiến độ ôn tập.
- Gói chuyển máy chứa ngân hàng câu hỏi, danh sách bộ câu hỏi, cấu hình, bộ đang dùng và tiến độ ôn tập để khôi phục đầy đủ trên máy khác.
- PWA đã đổi cache sang V45.30.

- V45.30 sửa 3 nút Dùng tất cả bộ / Dùng bộ có nhập số câu / Bỏ chọn bộ đang dùng: không render lặp và không lưu cấu hình lặp gây treo.

- V45.30 sửa riêng checkbox tự gộp theo tỷ lệ: bật/tắt độc lập, không phát sinh vòng lặp lưu cấu hình/render gây treo.

V45.30:
- Sửa triệt để checkbox Tự gộp theo tỷ lệ để không kích hoạt vòng lặp lưu cấu hình/render.
- Khi tự gộp, số câu được làm tròn theo phương pháp phần dư lớn nhất để tổng khớp số câu yêu cầu.
