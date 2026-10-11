# toolgopfilepptbylekimlam - Công cụ gộp file PowerPoint (.pptx) trực tuyến theo thứ tự

Bản quyền © 2026 **Lê Kim Lâm** (`lekimlam2707@gmail.com`). Mọi quyền được bảo lưu.

Ứng dụng web giúp người dùng tải lên nhiều bài thuyết trình PowerPoint (`.pptx`), sắp xếp thứ tự ưu tiên trực quan bằng kéo thả, xem trước nội dung từng slide và ghép lại thành 1 file duy nhất với đầy đủ bố cục, màu sắc và hình ảnh gốc.

---

## ⚡ Ưu điểm nổi bật (Kiến trúc In-Browser)
- **100% Xử lý trực tiếp trên trình duyệt (Client-Side OpenXML)**: Toàn bộ quá trình ghép slide diễn ra trong bộ nhớ trình duyệt nhờ thư viện `JSZip`.
- **Không bao giờ bị lỗi HTTP 500 hay HTTP 413**: Do không gửi tệp qua máy chủ trung gian, không bị giới hạn dung lượng request (Payload Too Large) hay timeout của máy chủ.
- **Bảo mật tuyệt đối**: Dữ liệu slide và tài liệu nhạy cảm không rời khỏi máy tính của người dùng.
- **Chạy được trên mọi nền tảng**: Tương thích 100% với **Vercel**, **GitHub Pages**, Netlify, Cloudflare Pages...

---

## 🚀 Hướng dẫn Deploy lên Vercel (Miễn phí & Tự động 100%)

1. Đẩy mã nguồn lên **GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Update toolgopfilepptbylekimlam client-side engine"
   git branch -M main
   git remote add origin https://github.com/<tai-khoan-cua-ban>/<ten-repo>.git
   git push -u origin main
   ```
2. Mở [vercel.com](https://vercel.com) và bấm **"Add New..."** → **"Project"**.
3. Chọn Repository vừa tạo và bấm **Deploy**.
4. Vercel tự động build và cung cấp link web chạy vĩnh viễn không lo lỗi máy chủ!

---

## 🚀 Hướng dẫn Deploy lên GitHub Pages
Vì ứng dụng hiện tại chạy 100% phía trình duyệt (Client-Side SPA), bạn có thể chạy trên **GitHub Pages**:
1. Vào repository trên GitHub → **Settings** → **Pages**.
2. Tại mục **Build and deployment** → Chọn **GitHub Actions** → Chọn workflow **Vite** (hoặc Static HTML).
3. Trang web sẽ hoạt động trực tiếp tại: `https://<ten-user>.github.io/<ten-repo>/`.

---

## 💻 Chạy thử nghiệm trên máy tính
```bash
npm install
npm run dev
```
Truy cập: `http://localhost:3000`.
