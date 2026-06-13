Trắc nghiệm Excel Offline PWA V26

Cách dùng:
1) Giải nén toàn bộ thư mục ZIP, không mở trực tiếp file trong ZIP.
2) Đảm bảo index.html và xlsx.full.min.js nằm cùng thư mục.
3) Mở bằng localhost để cài PWA:
   python -m http.server 8080
   sau đó truy cập http://localhost:8080
4) Bấm Ép cập nhật PWA nếu đã từng dùng bản cũ.

V26 sửa:
- Khi nhập số câu ở từng bộ, ô Số câu hỏi / Số câu mỗi lượt ôn tự điền tổng tương ứng.
- Nút Đã học ở thanh dưới chỉ xác nhận nhanh cho các câu trả lời đúng của phiên ôn tập hiện tại.
- Giữ SheetJS trong ZIP PWA và giữ logic V25: đề ngẫu nhiên theo từng bộ, ôn tập lần lượt theo từng bộ.
