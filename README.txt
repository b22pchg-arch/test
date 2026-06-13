Trắc nghiệm Excel Offline PWA V7

- Mở index.html qua localhost/HTTPS để dùng PWA và service worker.
- Dữ liệu 160 câu đã nhúng sẵn trong index.html.
- Bản V7 giữ giao diện toàn màn hình khi làm bài/khi xem kết quả.
- Phần giải thích lấy phương án sai làm trọng tâm và so khớp chuỗi từ đầu đến cuối:
  + nếu khác ngay từ đầu: hiển thị phương án đúng bắt đầu bằng từ/cụm nào và phương án sai không có phần mở đầu đó;
  + nếu giống một đoạn rồi mới khác: hiển thị đúng từ/cụm đầu tiên bị lệch;
  + vẫn tính % giống, % bao phủ ý đúng, phần thiếu và phần sai/thêm.
- Vẫn đọc Excel khác bằng SheetJS khi cần.
