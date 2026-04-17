# CLAUDE.md — Credit Card Reconciliation App

## TÓM TẮT 30 GIÂY
Hệ thống đối soát hóa đơn tự động. 
- **Input:** Ảnh hóa đơn (.jpg, .png) + Sao kê ngân hàng (Sheet).
- **Xử lý:** Cloud Vision (OCR) -> Gemini (Parse) -> GAS Logic (Match theo Amount, Date, Currency).
- **Output:** Tự động điền Account Code, Summary và gán Invoice Reference vào sheet.

## TECH STACK
| Component | Technology |
|---|---|
| Platform | Google Apps Script (V8) |
| OCR | Google Cloud Vision API |
| LLM | Gemini 1.5 Flash / 2.0 Flash |
| Storage | Google Drive |
| Auth | Service Account (JWT) |

## CẤU TRÚC CODE
- `Mã.js`: Entry point & Folder Map.
- `01_Core/`: `Constants.js` (Config), `Utils.js` (Date/Amount parsing), `Config.js` (Sheet lookup).
- `02_Sheets/`: `SheetOps.js` (Formatting, Conditional Rules).
- `03_Services/`: 
    - `Matching.js`: Logic đối soát chính (Score-based).
    - `DriveOps.js`: Quản lý file, folder và Auth Vision API.
- `04_UI/`: `UI.js` (Sidebar, Upload Dialog).
- `05_Testing/`: `Tests.js` (Diagnostic tests).

## CRITICAL CONSTRAINTS ⚠️
- ⚠️ **Match Keys:** Phải khớp tuyệt đối về **Currency** và **Amount**. **Date** cho phép lệch ±4 ngày.
- ⚠️ **OCR Pattern:** Luôn dùng `callCloudVision` cho từng ảnh trước, sau đó mới gộp vào 1 call `callGeminiBatch` duy nhất cho toàn bộ sheet/batch để tối ưu cost và speed.
- ⚠️ **Image Type:** Chỉ hỗ trợ ảnh chèn "Over cells" (OverGridImage). Ảnh "In cell" sẽ bị bỏ qua.
- ⚠️ **Permissions:** Yêu cầu GCP Project đã enable Cloud Vision API và có Service Account JSON hợp lệ.

## COMMON TASKS
### 1. Thêm một trường dữ liệu mới (ví dụ: VAT Amount)
1. Cập nhật `COL` constants trong `01_Core/Constants.js`.
2. Sửa `buildBatchInvoicePrompt` trong `03_Services/Matching.js` để AI trích xuất thêm trường này.
3. Cập nhật `applyMatch` để ghi dữ liệu vào cột mới.

### 2. Thay đổi AI Model
1. Sửa `AI_MODEL` trong `01_Core/Constants.js`.
2. Đảm bảo model name đúng theo [Google AI Studio](https://aistudio.google.com/).

## DATA SCHEMAS
### Gemini Output JSON
```json
[
  {
    "imageIndex": 1,
    "date": "2026-04-17",
    "amount": 150000,
    "currency": "VND",
    "merchant": "Starbucks",
    "summary": "Coffee and meeting"
  }
]
```

## QUICK REFERENCE
- **Log Location:** `Apps Script Editor -> View -> Executions`.
- **Test Command:** Menu `💳 Corporate Card -> 🧪 Run All Tests`.
- **Config Table:** Sheet `Config`, bảng `API Credentials`.
