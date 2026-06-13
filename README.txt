TRẮC NGHIỆM EXCEL OFFLINE PWA V18

Bản V18 khôi phục điều kiện cài PWA đúng chuẩn:
- index.html đã có <link rel="manifest">
- manifest.webmanifest có id, scope, start_url, display=standalone
- có icon PNG 192x192 và 512x512 phục vụ installability
- sw.js cache đầy đủ index/app/xlsx/manifest/icon
- có nút “Cài PWA” và thông báo chẩn đoán tình trạng cài đặt

Cách chạy để hiện “Cài đặt ứng dụng” thay vì chỉ “Tạo lối tắt”:
1) Giải nén ZIP.
2) Mở terminal trong thư mục đã giải nén.
3) Chạy: python -m http.server 8080
4) Mở Chrome/Edge: http://localhost:8080
5) Bấm “Ép cập nhật PWA”, sau đó bấm “Cài PWA”.

Lưu ý: nếu mở trực tiếp index.html bằng file://, nhiều trình duyệt chỉ cho “Tạo lối tắt”, không cài PWA đầy đủ vì service worker/manifest không hoạt động như khi chạy qua localhost/HTTPS.
