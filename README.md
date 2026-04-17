# Credit Card Reconciliation System (v3.0)

Hệ thống đối soát chi phí thẻ doanh nghiệp tự động sử dụng Google Apps Script, tích hợp Cloud Vision OCR và Gemini AI.

## 🌟 Tổng quan
Hệ thống giúp tự động hóa quy trình đối soát hóa đơn (receipts) với sao kê ngân hàng (bank statements). Thay vì nhập liệu thủ công, người dùng chỉ cần tải ảnh hóa đơn lên Drive, hệ thống sẽ tự động quét, trích xuất thông tin và tìm dòng giao dịch tương ứng trên bảng tính.

## 🛠 Tech Stack
- **Platform:** Google Apps Script (V8)
- **Database:** Google Sheets
- **Storage:** Google Drive
- **OCR:** Google Cloud Vision API
- **AI Engine:** Google Gemini AI (1.5 Flash / 2.0 Flash)

## 📁 Cấu trúc thư mục (Modular)
- `01_Core/`: Chứa các hằng số, tiện ích và logic đọc cấu hình.
- `02_Sheets/`: Các thao tác với Spreadsheet (formatting, tạo sheet).
- `03_Services/`: Logic nghiệp vụ chính (Matching, OCR, Drive Operations).
- `04_UI/`: Quản lý Menu, Sidebar và Dialog.
- `05_Testing/`: Các bộ test case để kiểm tra hệ thống.

## 🚀 Hướng dẫn nhanh
1. **Cấu hình:** Điền thông tin API Key và Drive Folder ID vào sheet `Config`.
2. **Khởi tạo:** Chạy `Initialize Main Sheet` từ menu.
3. **Sử dụng:** 
   - Dán dữ liệu sao kê vào sheet chính.
   - Upload hóa đơn qua Sidebar.
   - Nhấn `Scan & Reconcile` để AI thực hiện đối soát.

## 📄 Tài liệu liên quan
- [CLAUDE.md](CLAUDE.md): Tài liệu dành cho AI (Context/Rules).
- [SOP.md](SOP.md): Hướng dẫn vận hành chi tiết.
- [CHANGELOG.md](CHANGELOG.md): Lịch sử cập nhật.
