# Quy tắc Đặc vụ AI: Hệ thống Bảo mật 3 Lớp (Code > Data > DB)

Tệp này quy định các chuẩn mực bảo mật bắt buộc đối với tất cả hoạt động sinh mã, phát triển và cấu hình trong dự án. Tất cả đặc vụ AI tham gia vào dự án phải tuân thủ nghiêm ngặt các quy tắc này để bảo vệ ứng dụng khỏi các lỗ hổng rò rỉ dữ liệu, tấn công trục lợi chi phí (Denial of Wallet), hoặc chiếm quyền điều khiển.

---

## 1. LỚP CODE (Code Layer Security)
Lớp Code là hàng phòng thủ đầu tiên. Mã nguồn phải có tính tự vệ cao, kiểm soát chặt chẽ luồng thực thi và dữ liệu cấu hình.

### 1.1 Quản lý Khóa Bí mật và Biến Môi trường (Secret Management)
*   **Tuyệt đối không để lộ Key:** Không bao giờ hardcode API Keys, Token, Service Account Credentials hoặc client secrets vào bất kỳ tệp dữ liệu tĩnh hay mã nguồn client-side nào.
*   **Sử dụng File Cấu hình Biên:** Tất cả biến bí mật phải được khai báo trong `.env.example` và truy xuất thông qua `process.env` phía máy chủ (Server-side).
*   **Lazy Initialization:** Khởi tạo các SDK có chứa khóa bí mật (như Firebase Admin SDK hay Google Auth OAuth2 client) theo dạng "khởi tạo lười" (Lazy/On-demand). Chỉ khởi tạo khi bắt đầu có yêu cầu nghiệp vụ thực tế, kèm theo kiểm tra tính hiện hữu của khóa nhằm tránh làm crash ứng dụng ngay khi startup nếu thiếu cấu hình.
    ```typescript
    // Ví dụ khởi tạo lười để tránh crash ứng dụng khi thiếu key môi trường:
    let oauth2Client: any = null;
    export function getOAuth2Client() {
      if (!oauth2Client) {
        const clientId = process.env.GOOGLE_CLIENT_ID;
        const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
        if (!clientId || !clientSecret) {
          throw new Error("Vui lòng cấu hình đầy đủ GOOGLE_CLIENT_ID và GOOGLE_CLIENT_SECRET!");
        }
        oauth2Client = new google.auth.OAuth2(clientId, clientSecret, ...);
      }
      return oauth2Client;
    }
    ```

### 1.2 An toàn Kiểu Dữ liệu và Kiểm thử (Type Safety & Static Analysis)
*   **TypeScript Strict Mode:** Tránh tối đa việc sử dụng kiểu `any` hoặc `as any` ép kiểu một cách bừa bãi. Sử dụng `unknown` và ép kiểu tường minh hoặc Type Guards khi xử lý payloads không rõ nguồn gốc.
*   **Type Imports:** Đặt tất cả các lệnh `import` lên đầu trang, sử dụng named imports và phân biệt rõ ràng imports loại dữ liệu (`import type`).

### 1.3 Quản lý Authentication và Session (Cookie & Tokens)
*   **Chỉ lưu token Server-side:** Khi tích hợp các cơ chế OAuth2 (Google API, Microsoft v.v.), hãy lưu giữ refresh tokens và access tokens ở database bảo mật hoặc Cookie bảo mật (`httpOnly`, `secure`, `sameSite: "none"` hoặc `"lax"`).
*   **Kiểm tra tính hợp lệ trước khi thực hiện hành động:** Các `/api/*` endpoint tuyệt đối phải có middleware xác thực người dùng gửi yêu cầu.

### 1.4 Cách ly Lỗi và Tránh Lộ Trích lục Hệ thống (Detailed Error Isolation)
*   **Không trả về Raw Traces:** Không bao giờ trả trực tiếp mã lỗi hệ thống, stack traces hoặc thông tin cấu trúc cơ sở dữ liệu nội bộ về phía Client. Hãy ghi nhật ký lỗi chi tiết trên Server và trả về Client các mã thông báo lỗi chung, bảo mật (Generic Error Message) kèm ID lỗi ẩn.

---

## 2. LỚP DỮ LIỆU (Data Layer Security)
Lớp Dữ liệu xử lý cách vận chuyển, lọc và ánh xạ dữ liệu giữa Client và Server nhằm chống độc hại hóa dữ liệu đầu vào.

### 2.1 Lọc Sạch và Giới hạn Payload (Input Sanitization & Payload Limits)
*   **Validate Toàn diện Payload:** Trước khi chuyển giao dữ liệu vào controller lưu trữ, dữ liệu đầu vào phải được xác thực chính xác thông qua helper hoặc thư viện kiểm duyệt kiểu.
*   **Khai báo giới hạn tệp và kích cỡ request body:** Chỉ định rõ ràng giới hạn dung lượng xử lý payload tại Server (ví dụ: `express.json({ limit: "5mb" })`).
*   **Xử lý chuỗi và ngăn chặn XSS:** Làm sạch các đầu vào dạng văn bản dài giàu HTML trước khi trả ra giao diện, không lạm dụng `dangerouslySetInnerHTML`.

### 2.2 Ánh xạ Dữ liệu Bảo vệ (Data Transfer Object - DTO Mapping)
*   **Lọc trường thông tin nhạy cảm:** Không truyền tải nguyên bản cấu trúc đối tượng dữ liệu trong cơ sở dữ liệu về phía Client. Phải ánh xạ qua một lớp DTO để loại bỏ các trường hệ thống không cần thiết hoặc nhạy cảm (như salt mật khẩu, dữ liệu cấu hình, logs nội bộ, chi tiết phân quyền hệ thống).

### 2.3 Cách ly và Tách biệt Thông tin Nhạy cảm (PII Isolation)
*   **Tách biệt PII:** Các thông tin định danh cá nhân (PII) như số điện thoại, email riêng, địa chỉ nhà phải được lưu giữ ở các cấu trúc thực thể riêng biệt hoặc phân tách quyền đọc ghi cực kỳ nghiêm ngặt (Ví dụ: tách biệt thành bảng `users/{userId}/public` và `users/{userId}/private`).

---

## 3. LỚP DATABASE (DB Layer Security)
Đảm bảo an toàn tuyệt đối tại nơi dữ liệu được lưu trữ. Kết hợp cả bảo mật Firestore (NoSQL) và SQLite (SQL).

### 3.1 Phòng ngừa Tấn công SQL Injection (SQLite/better-sqlite3)
*   **Luôn dùng Parameterized Queries:** Tuyệt đối không dùng phép cộng chuỗi trực tiếp để ghép câu lệnh SQL cùng tham số từ người dùng. Sử dụng prepared statements với các ký tự thay thế `?` hoặc `@param`.
    ```typescript
    // ĐÚNG ✅
    const stmt = db.prepare('SELECT * FROM staff WHERE id = ?');
    const result = stmt.get(staffId);

    // SAI ❌ (Nguy hiểm, dính SQL Injection)
    const result = db.exec(`SELECT * FROM staff WHERE id = '${staffId}'`);
    ```

### 3.2 Tám Trụ Cột Bảo Mật Firebase Firestore (Firestore Hardening)
Khi tích hợp Firestore Client SDK, file `firestore.rules` bắt buộc phải được thiết kế và kiểm thử dựa trên mô hình Zero-Trust.

1.  **Cổng Chính Master Gate (Xác thực Liên đới):**
    *   Mọi bộ quy tắc của sub-collection (ví dụ: `/tasks/`) phải kiểm tra quyền sở hữu hoặc thành viên dựa trên tài liệu cha (ví dụ: `/projects/`) bằng phương thức `get()`.
    *   Không được lưu trữ các danh sách người tham gia quá lớn không giới hạn trong Arrays. Phải chuyển chúng thành góc sub-collection để tối ưu hóa và chống rách tệp dữ liệu.
2.  **Sơ Đồ Xác Thực Chống Kẽ hở Ghi (Validation Blueprints):**
    *   Tách biệt toàn bộ logic xác thực của một đối tượng ra một hàm helper riêng biệt dạng `isValid[Entity](data)`.
    *   Hàm kiểm duyệt này phải được áp dụng đồng thời cho cả quyền `create` và `update`.
    *   Đối với quyền khởi tạo (`create`), kiểm tra nghiêm ngặt số lượng khóa và sự tồn tại của khóa bắt buộc: `data.keys().hasAll(['req1', 'req2']) && data.keys().size() == 2`.
    *   Xác minh trường nhận danh tác giả luôn bằng đúng `request.auth.uid`.
3.  **Khóa Chặt Biến Đường Dẫn (Path Variable Hardening):**
    *   Áp dụng hàm kiểm duyệt ID hợp lệ `isValidId(id)` đối với tất cả các biến đường dẫn ID động nhận vào (nhung `{projectId}`, `{taskId}`) trong các tác vụ đơn tài liệu (`get`, `create`, `update`, `delete`). Tránh cho kẻ tấn công chèn các chuỗi ký tự rác phá hoại hệ thống.
4.  **Phát Phân Tầng Quyền Hạn Ghi/Sửa (Tiered Identity Logic):**
    *   Phân chia các nhóm trường dữ liệu được sửa thành các hành động (Actions) cụ thể.
    *   Cho phép người dùng thông thường chỉ thay đổi các trường nhất định (như `status`) bằng cách giới hạn chính xác: `incoming().diff(existing()).affectedKeys().hasOnly(['status'])`.
5.  **Bảo vệ Toàn diện Mảng Dữ liệu (Total Array Guarding):**
    *   Mọi mảng dữ liệu gửi lên phải được kiểm soát chặt chẽ về số lượng thông qua `.size() <= MAX` để tránh phình to dung lượng tài liệu vô tội vạ.
6.  **Cô lập PII Công khai/Riêng tư:**
    *   Tuyệt đối cấm viết luật đọc bừa bãi kiểu `allow read: if isSignedIn();` trên bộ sưu tập chứa email hay thông tin cá nhân. Phải áp dụng luật sở hữu cá nhân `resource.data.userId == request.auth.uid`.
7.  **Đảm Bảo Tính Nguyên Tố (Atomicity Guarantee):**
    *   Các giao dịch đồng bộ trạng thái giữa các tài liệu liên đới phải sử dụng `existsAfter(...)` hoặc `getAfter(...)` để bảo đảm cấu trúc giao dịch được thực hiện hoặc hủy bỏ hoàn toàn đồng thời.
8.  **Xác Thực Danh Sách Luật Chặt Chẽ (Secure List Queries):**
    *   Mọi luật `allow list` bắt buộc phải kiểm duyệt điều kiện dựa vào `resource.data` (ví dụ: `resource.data.ownerId == request.auth.uid`). Không bao giờ trông cậy vào việc Client tự lọc bằng câu lệnh truy vấn `.where()`.
    *   **Tối ưu hóa tránh rách ví tiền (Denial of Wallet):** Tuyệt đối không chạy các hàm `get()` hay `exists()` tốn phí bên trong khối `allow list` để bảo vệ tài khoản khỏi các đợt quét dữ liệu O(n) độc hại.

### 3.3 Quy Trình Kiểm Thử Và Thử Nghiệm Red Team
*   Hải tặc Shadow Update: Kiểm thử gửi payload chứa các trường được phép kèm theo một "Ghost Field" (trường giả mạo như `role: 'admin'`). Luật quy tắc bảo mật phải từ chối yêu cầu này ngay lập tức nhờ bộ lọc kiểm tra nghiêm ngặt `affectedKeys().hasOnly()`.
*   Temporal Integrity: Các mốc thời gian như `createdAt`, `updatedAt` phải được đối sánh trực tiếp với thời gian thực tế của hệ thống lưu trữ `request.time` thay vì tin tưởng thời gian gửi từ Client: `incoming().updatedAt == request.time`.
*   Immutability: Các trường dữ liệu khởi tạo cố định (như `projectId`, `createdAt`, `ownerId`) phải được khóa cứng không cho cập nhật: `incoming().ownerId == existing().ownerId`.

---

*Lưu ý: Mọi thay đổi nghiệp vụ liên quan đến Cấu trúc Dữ liệu hoặc Cơ sở dữ liệu của dự án phải chạy lệnh `npm run lint` hoặc `compile_applet` để bảo đảm hệ thống hoạt động chính xác trước khi vận hành.*
