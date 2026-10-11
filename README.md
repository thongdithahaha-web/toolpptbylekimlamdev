# PPTMerge - Công cụ gộp file PowerPoint (.pptx) trực tuyến theo thứ tự

Ứng dụng web giúp người dùng tải lên nhiều bài thuyết trình PowerPoint (`.pptx`), sắp xếp thứ tự trực quan bằng kéo thả, xem trước nội dung từng slide và ghép lại thành 1 file duy nhất với đầy đủ bố cục, màu sắc và hình ảnh gốc.

---

## 🚀 Hướng dẫn Deploy (Triển khai lên máy chủ)

### 1. Triển khai lên Vercel (Khuyên dùng - Miễn phí & Nhanh nhất)
Ứng dụng đã được cấu hình sẵn file `vercel.json` để chạy cả **Frontend (React)** và **Backend API (Node.js)** trên Vercel:

1. Đẩy mã nguồn lên **GitHub** của bạn:
   ```bash
   git init
   git add .
   git commit -m "Initial commit PPTMerge"
   git branch -M main
   git remote add origin https://github.com/<tai-khoan-cua-ban>/<ten-repo>.git
   git push -u origin main
   ```
2. Truy cập [vercel.com](https://vercel.com) và đăng nhập bằng GitHub.
3. Bấm **"Add New..."** → **"Project"** → Chọn Repository vừa đẩy lên.
4. Ở phần cài đặt:
   - **Framework Preset**: Vite
   - **Build Command**: `npm run build`
   - **Output Directory**: `dist`
5. Bấm **Deploy**. Vercel sẽ tự động cấu hình và cung cấp tên miền miễn phí (VD: `ppt-merge.vercel.app`).

---

### 2. Về GitHub Pages
> ⚠️ **Lưu ý quan trọng về GitHub Pages**:
> - **GitHub Pages** chỉ hỗ trợ lưu trữ các trang web **tĩnh (Static HTML/CSS/JS)**, **KHÔNG hỗ trợ chạy máy chủ Node.js/Express**.
> - Vì chức năng gộp PowerPoint cần xử lý cấu trúc tệp OpenXML trên máy chủ, bạn **không thể chạy toàn bộ ứng dụng chỉ trên GitHub Pages**.
> - **Cách tốt nhất**: Lưu mã nguồn trên GitHub, và kết nối với **Vercel** hoặc **Render** để deploy tự động mỗi khi có thay đổi code trên GitHub!

---

### 3. Triển khai lên Render.com (Hoàn toàn miễn phí)
Nếu bạn muốn chạy dạng dịch vụ Web Service độc lập:
1. Đăng ký tài khoản tại [render.com](https://render.com).
2. Chọn **New** → **Web Service** → Kết nối repository GitHub của bạn.
3. Thiết lập thông số:
   - **Environment**: `Node`
   - **Build Command**: `npm install && npm run build`
   - **Start Command**: `npm run start` (hoặc `npx tsx server.ts`)
4. Bấm **Create Web Service**.

---

### 4. Chạy cục bộ trên máy tính (Local Development)
```bash
# 1. Cài đặt các gói phụ thuộc
npm install

# 2. Khởi chạy máy chủ phát triển (cổng 3000)
npm run dev
```
Truy cập trình duyệt tại `http://localhost:3000`.

---

## 🛠️ Công nghệ sử dụng
- **Frontend**: React 19, TypeScript, Tailwind CSS, Lucide Icons, JSZip.
- **Backend**: Node.js, Express, Multer, PPTX-Automizer, PPTXGenJS.
- **Tính năng nổi bật**:
  - Hỗ trợ tải phân mảnh chống lỗi HTTP 413 khi tải file nặng.
  - Xem trước tiêu đề slide và lọc slide cần ghép.
  - Sắp xếp vị trí slide tự do theo nhu cầu.
