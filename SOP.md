# Standard Operating Procedure (SOP)
## Hệ thống Đối soát Giao dịch Thẻ Doanh nghiệp (v3.0)

Tài liệu này hướng dẫn quy trình vận hành dành cho người dùng cuối và quản trị viên.

---

### 1. Thiết lập ban đầu (Admin)
1. **Khởi tạo Config:** `Menu -> Initialize Config Template`.
2. **Cài đặt API:**
   - Dán **Gemini API Key** (từ AI Studio).
   - Dán **Service Account JSON** (từ GCP Console - Cloud Vision).
   - Điền **Drive Folder ID** (ID thư mục chứa ảnh).
3. **Khởi tạo Sheet chính:** `Menu -> Initialize Main Sheet`.

### 2. Quy trình hàng tháng (User)
#### Bước 1: Nhập sao kê
- Copy dữ liệu từ file Excel ngân hàng vào cột A-P của sheet `Corporate card details`.
- Đảm bảo cột **Date** và **Amount** đúng định dạng.

#### Bước 2: Tải hóa đơn
- Mở `Menu -> Process Invoices (Upload/Scan)`.
- Kéo thả ảnh hoặc dán từ Clipboard.
- Nhấn **Upload All to Drive**.

#### Bước 3: Đối soát AI
- Tại Sidebar, nhấn **Scan Drive Folder & Process**.
- Chờ hệ thống quét (Phase A) -> AI phân tích (Phase B/C) -> Đối soát (Phase D).
- **Kết quả:**
  - `✅ Matched`: Khớp hoàn toàn.
  - `⚠️ Needs Review`: Có từ 2 giao dịch giống nhau trở lên cho 1 hóa đơn.
  - `⏳ Pending`: Không tìm thấy dòng sao kê tương ứng.

#### Bước 4: Hoàn thiện dữ liệu
- Chọn **Account Code** (Cột Q) -> Cột R tự động nhảy tên tài khoản.
- Review cột **Summary** do AI viết.
- Với dòng `⚠️ Needs Review`, mở Sidebar để chọn dòng khớp chính xác.

### 3. Xử lý sự cố
- **AI không đọc được ảnh:** Kiểm tra độ phân giải, đảm bảo ảnh không quá mờ. Các file lỗi sẽ nằm trong thư mục `Error/` trên Drive.
- **Không khớp ngày:** Hệ thống cho phép lệch ±4 ngày. Nếu hóa đơn thanh toán quá trễ so với ngày ghi sổ, cần đối soát thủ công.
- **Lỗi 429:** Hết quota API. Hãy chờ 1 phút và thử lại.

---

### 4. Quản lý File trên Drive
- Ảnh sau khi khớp sẽ được di chuyển vào folder `DONE-MM-YYYY`.
- Tên file được đổi thành `[Done-DD-MM]...` để dễ tra cứu.
