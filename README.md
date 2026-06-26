# Gmail SMTP DOCX Mailer

Web Node.js đơn giản để gửi nội dung từ file `EU RATE BÁO GIÁ.docx` tới danh sách trong `HAMBURG.xlsx`. App đọc tất cả trang tính trong file Excel, gửi từng người nhận riêng lẻ, mỗi email cách nhau 2 phút theo cấu hình `SEND_INTERVAL_MS`.

## Cài đặt

```bash
npm install
cp .env.example .env
npm run dev
```

Mở trình duyệt tại:

```text
http://localhost:4000
```

## Cấu hình `.env`

```env
PORT=4000
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_SECURE=true
SMTP_USER=your-gmail@gmail.com
SMTP_PASS=your-google-app-password
MAIL_FROM="Your Name <your-gmail@gmail.com>"
MAIL_SUBJECT=EU RATE BAO GIA
DOCX_PATH=./EU RATE BÁO GIÁ.docx
EXCEL_PATH=./HAMBURG.xlsx
SEND_INTERVAL_MS=120000
SEND_FIRST_IMMEDIATELY=false
```

Ghi chú:

- Gmail bắt buộc dùng `App Password` cho SMTP nếu tài khoản bật 2-Step Verification. Không dùng mật khẩu đăng nhập Gmail thường.
- `SMTP_HOST=smtp.gmail.com`, `SMTP_PORT=465`, `SMTP_SECURE=true` là cấu hình SSL ổn định cho Nodemailer.
- `MAIL_FROM` nên trùng với `SMTP_USER` để Gmail không từ chối hoặc đổi sender.
- Danh sách người nhận lấy từ tất cả trang tính trong `EXCEL_PATH`: cột 1 là `FULL NAME`, cột 5 là `EMAIL`.
- `SEND_INTERVAL_MS=120000` là 2 phút. Khi test nhanh có thể đổi thành `10000`.
- `SEND_FIRST_IMMEDIATELY=false` nghĩa là email đầu tiên cũng chờ 2 phút. Đổi thành `true` nếu muốn gửi email đầu tiên ngay.

## Cách app xử lý file DOCX

- DOCX được chuyển sang HTML để hiển thị trong thân email.
- Ảnh trong DOCX được nhúng inline bằng CID để mail client hiển thị trong email.
- File DOCX gốc không được đính kèm; email chỉ gửi nội dung HTML và ảnh inline.

## API nhanh

- `POST /api/send-campaign`: tạo job gửi mail.
- `GET /api/jobs/current`: xem job đang chạy.
- `DELETE /api/jobs/:id`: hủy job.
