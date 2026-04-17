# Changelog

Tất cả các thay đổi đáng chú ý đối với dự án này sẽ được ghi lại trong file này.

## [v3.0] — 2026-04-17
### Added
- **Cấu trúc Modular:** Tách code thành các thư mục `01_Core`, `02_Sheets`, `03_Services`, `04_UI`, `05_Testing`.
- **Drive-First Workflow:** Tối ưu hóa việc tải lên và quản lý file trực tiếp trên Drive.
- **Improved Matching:** Tăng độ trễ ngày lên ±4 ngày và tinh chỉnh trọng số AI.
- **Auto-Account Code:** Gemini gợi ý mã tài khoản kế toán dựa trên tên cửa hàng.

### Technical
- Chuyển sang sử dụng **Gemini 2.0 Flash** làm mặc định.
- Tối ưu hóa batching: 1 call Gemini cho toàn bộ ảnh trong 1 lần scan.
- Thêm bộ test suite tự động trong `Tests.js`.

### Lessons Learned
- Việc sử dụng Cloud Vision (OCR) riêng biệt với Gemini (Parse) giúp tiết kiệm token và tăng độ chính xác của text thô đáng kể so với việc để Gemini tự đọc ảnh trực tiếp (Native Multimodal) khi xử lý số lượng lớn.

---

## [v2.1] — 2026-04-15
### Added
- Tích hợp Cloud Vision API cho OCR.
- Thêm sheet `Config` để quản lý API Keys.

## [v1.0] — 2026-04-10
- Phiên bản MVP đầu tiên: Nhập sao kê và đối soát thủ công.
