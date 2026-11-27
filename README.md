# Tích Hợp Tổng Hợp Giáo Án AI (Bản Độc Lập)

Thư mục độc lập sẵn sàng đưa lên dự án mới (Vercel, Netlify, GitHub Pages hoặc chạy local bằng Live Server / `npx serve`).

## Cấu trúc tệp
- `index.html`: Giao diện chính + hộp thoại nhập Gemini API Key (kèm link lấy key tại `https://aistudio.google.com/apikey`).
- `app.bundle.js`: Mã nguồn đã đóng gói sẵn, chạy trực tiếp không cần bước build.
- `app.js`: Mã nguồn React gốc (nếu cần chỉnh sửa và đóng gói lại).
- `assets/`:
  - `docxModel.js`: Đọc/ghi `.docx` (giữ nguyên định dạng, bảng, ảnh, công thức), đọc PPCT `.xlsx`, xuất giáo án mới `.docx`.
  - `frameworksData.js`: Dữ liệu chuẩn QĐ 2422/QĐ-BGDĐT (282 YCCĐ lớp 1–12), CV 6206/BGDĐT-GDPT (`ATM1`–`ATM8`), TT 02/2025/TT-BGDĐT (24 năng lực số) và sơ đồ liên kết ATM–NLS.
  - `jszip.min.js`, `pdf.min.js`, `pdf.worker.min.js`: Thư viện xử lý Word/Excel/PDF chạy 100% trên trình duyệt.

## Chạy thử nhanh
Mở thư mục này bằng một web server tĩnh bất kỳ (ví dụ `npx serve .` hoặc `python -m http.server 8080`) rồi truy cập `http://localhost:8080`.
