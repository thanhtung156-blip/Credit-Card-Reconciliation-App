# Standard Operating Procedure (SOP)
## Hệ thống Đối soát Giao dịch Thẻ Doanh nghiệp (v1.0)

Tài liệu này hướng dẫn quy trình vận hành hàng ngày dành cho Bộ phận Kế toán/Tài chính để thực hiện đối soát hóa đơn với sao kê ngân hàng.

---

### 1. Chuẩn bị (Chỉ thực hiện lần đầu)
1. **Khởi tạo Cấu hình**: Vào menu `💳 Corporate Card` -> `⚙️ Initialize Config Template`.
2. **Điền thông tin**: Tại sheet **Config**, điền đủ API Key (Gemini), Service Account JSON (Cloud Vision), Drive Folder ID và chọn Model AI (mặc định là `gemini-2.0-flash`).
3. **Khởi tạo bảng tính**: Vào menu `💳 Corporate Card` -> `⚙️ Initialize Main Sheet` để chuẩn bị các cột từ Q đến X.

---

### 2. Quy trình làm việc hàng ngày / hàng tháng

#### Bước 1: Nhập liệu Sao kê Ngân hàng
* Copy dữ liệu từ file Excel của ngân hàng và dán vào sheet **Corporate card details** (từ cột A đến P).
* Đảm bảo cột Ngày tháng (C) và Số tiền (J) chính xác.

#### Bước 2: Tải lên Hóa đơn (Receipts)
1. Vào menu `💳 Corporate Card` -> `📷 Process Invoices (Upload/Scan)`.
2. Vùng **1. Upload & Paste**:
   * Kéo thả file ảnh vào hoặc nhấn **Ctrl+V** để dán ảnh trực tiếp từ clipboard.
   * Nhấn **Upload All to Drive** để lưu hóa đơn vào thư mục gốc trên Google Drive.
3. Cửa sổ sẽ tự động đóng sau khi tải lên xong. Tiến trình tải lên sẽ hiện ở Sidebar bên phải.

#### Bước 3: Quét và Đối soát tự động (Reconciliation)
1. Mở lại menu `📷 Process Invoices (Upload/Scan)`.
2. Vùng **2. Scan & Reconcile**: Nhấn nút **Scan Drive Folder & Process**.
3. **Theo dõi tiến trình**: Quan sát Panel bên phải (Sidebar) để xem AI đang xử lý hóa đơn nào.
4. **Kết quả**:
   * Các dòng khớp sẽ có màu xanh và trạng thái `✅ Matched`.
   * Các dòng cần kiểm tra lại sẽ có màu vàng `⚠️ Needs Review`.
   * Hệ thống tự động điền **Account Code, I/O Code, Cost Center** cho các dòng đã khớp.
   * File ảnh trên Drive sẽ được tự động đổi tên thành `[Done-DD-MM]...` và di chuyển vào thư mục `DONE-MM-YYYY`.
   * Các file không khớp (Unmatched) sẽ được để lại thư mục gốc để bạn kiểm tra.

#### Bước 4: Kiểm tra và Hoàn thiện (Audit)
*   **Kiểm tra lý do Unmatch**: Nếu một hóa đơn không tự động khớp, hãy vào sheet **OCR Audit**. Kiểm tra cột **"Match Reason"** để biết sai lệch ở đâu.
*   **Quản lý ảnh**: Các ảnh đã khớp sẽ có tiền tố `[Done-DD-MM]` trên Drive và nằm trong folder `DONE-MM-YYYY`.
*   **Điền mã tài khoản**: AI đã tự động đề xuất Account Code. Bạn chỉ cần chọn lại nếu AI gợi ý chưa chính xác. Tên tài khoản (Cột R) sẽ tự động hiển thị.
*   **Sheet Ngày tháng**: Hệ thống tự động tạo các sheet dạng `DD-MM` (ví dụ `02-04`). Các sheet này được để trắng hoàn toàn để bạn có thể chủ động chèn ảnh hóa đơn vào sau này.

#### Bước 5: Xuất dữ liệu
* Nhấn `💳 Corporate Card` -> `✅ Validate before export` để kiểm tra lần cuối các trường thông tin còn thiếu.
* Dữ liệu đã sẵn sàng để đối chiếu với sổ cái hoặc phần mềm kế toán.

---

### 3. Các lưu ý quan trọng
* **Ảnh chụp**: Hóa đơn cần rõ nét, không bị lóa sáng để AI đọc chính xác số tiền và ngày tháng.
* **Sai số**: Hệ thống cho phép sai lệch ngày tháng tối đa 4 ngày và sai số tiền nhỏ (phí cà thẻ).
* **Trạng thái**: Chỉ các dòng có trạng thái trống hoặc `⏳ Pending` mới được hệ thống đưa vào đối soát.
