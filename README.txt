Trắc nghiệm Excel Offline PWA V25

Cách dùng:
1) Giải nén toàn bộ thư mục ZIP, không mở trực tiếp file trong ZIP.
2) Đảm bảo index.html và xlsx.full.min.js nằm cùng thư mục.
3) Mở bằng localhost để cài PWA:
   python -m http.server 8080
   sau đó truy cập http://localhost:8080
4) Bấm Ép cập nhật PWA nếu đã từng dùng bản cũ.

V25 sửa:
- Bổ sung xlsx.full.min.js vào ZIP PWA và cache offline.
- Sửa lỗi xem trước ngân hàng không hiện do hàm nhãn phương án bị thiếu.
- Giữ logic V24: đề thi lấy ngẫu nhiên theo số lượng từng bộ, ôn tập lấy lần lượt theo từng bộ.
