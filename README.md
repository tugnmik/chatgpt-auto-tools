# ⚡ ChatGPT Auto Tools

Công cụ tự động hóa đăng ký và quản lý tài khoản ChatGPT với giao diện đồ họa hiện đại.

![Version](https://img.shields.io/badge/version-2.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

---

## 📋 Tính năng chính

### 1. 🚀 Đăng ký tài khoản tự động (Auto Registration)
- Tự động tạo email tạm thời qua API tinyhost.shop
- Đăng ký tài khoản ChatGPT hoàn toàn tự động
- Hỗ trợ **multithreading** - đăng ký nhiều tài khoản đồng thời
- Tự động nhận và nhập mã OTP từ email
- Tùy chọn lấy checkout link (Plus/Business)
- Lưu thông tin tài khoản vào file Excel

### 2. 🔐 Bật MFA tự động (MFA Enrollment)
- Tự động bật xác thực 2 yếu tố (TOTP)
- Trích xuất secret key từ QR code
- Lưu TOTP secret vào Excel để sử dụng sau
- Hỗ trợ xử lý xác thực qua email/password

### 3. 💳 Lấy Checkout Link (Checkout Capture)
- Lấy link thanh toán ChatGPT Plus
- Lấy link thanh toán ChatGPT Business
- Hỗ trợ chọn nhiều tài khoản cùng lúc
- Tự động lưu link vào Excel

---

## 🛠️ Yêu cầu hệ thống

### Phần mềm
- **Python** 3.8 trở lên
- **Google Chrome** (phiên bản mới nhất)

### Thư viện Python
```bash
pip install requests
pip install undetected-chromedriver
pip install selenium
pip install colorama
pip install pyotp
pip install openpyxl
pip install customtkinter
```

Hoặc cài đặt tất cả cùng lúc:
```bash
pip install requests undetected-chromedriver selenium colorama pyotp openpyxl customtkinter
```

---

## 🚀 Cách sử dụng

### Khởi chạy ứng dụng
```bash
python chatgpt_auto_gui.pyw
```

Hoặc double-click file `chatgpt_auto_gui.pyw` (Windows)

### Chế độ OAuth2 Email (Tùy chọn)

Nếu bạn muốn sử dụng tài khoản Outlook/Hotmail qua OAuth2 thay vì TinyHost:

1. **Chuẩn bị file template**:
   Chạy lệnh sau để tạo file `oauth2.xlsx`:
   ```bash
   python create_oauth2_template.py
   ```

2. **Điền thông tin tài khoản**:
   Mở file `oauth2.xlsx` vừa tạo và điền thông tin vào các cột:
   - Cột A: Định dạng `email|password|refresh_token|client_id`
   - Cột B: `Status` (Để trống, tool sẽ tự điền "registered" khi thành công)

3. **Sử dụng trong GUI**:
   - Tại Tab Registration > Advanced Options
   - Chọn **Email Mode**: `OAuth2`
   - Nhấn nút 🔄 để load danh sách tài khoản

### Tab 1: Registration (Đăng ký)

1. **Số lượng tài khoản**: Nhập số tài khoản muốn đăng ký
2. **Mật khẩu**: Đặt mật khẩu chung cho tất cả tài khoản
3. **Network Mode**:
   - `Fast`: Mạng ổn định, tốc độ cao
   - `VPN/Slow`: Mạng không ổn định, timeout dài hơn
4. **Get Checkout Link**: Bật để lấy link thanh toán sau khi đăng ký
5. **Multithread Mode**: Bật để đăng ký nhiều tài khoản đồng thời
   - Chọn số threads (1-10)
   - Delay giữa các thread (ms)
6. Nhấn **▶ Start Registration** để bắt đầu

### Tab 2: MFA Enrollment (Bật MFA)

1. **Chọn file Excel**: File chứa danh sách tài khoản đã đăng ký
2. **Multithread Mode**: Bật để xử lý nhiều tài khoản đồng thời
3. Nhấn **▶ Start MFA** để bắt đầu

### Tab 3: Checkout Capture (Lấy link thanh toán)

1. **Load Accounts**: Tải danh sách tài khoản từ file Excel
2. **Chọn tài khoản**: Tick chọn các tài khoản cần lấy link
3. **Checkout Type**: Plus, Business, hoặc Both
4. Nhấn **▶ Start Capture** để bắt đầu

---

## 📁 Cấu trúc file

```
auto gpt/
├── chatgpt_auto_gui.pyw     # File chính
├── README.md                 # Hướng dẫn sử dụng
└── chatgpt_accounts_*.xlsx   # File lưu tài khoản (tự động tạo)
```

### Cấu trúc file Excel đầu ra

| Cột | Mô tả |
|-----|-------|
| Email | Địa chỉ email |
| Password | Mật khẩu |
| Cookie | Session cookie (JSON) |
| TOTP Secret | Secret key cho 2FA |
| Plus Checkout | Link thanh toán Plus |
| Business Checkout | Link thanh toán Business |
| Status | Trạng thái tài khoản |
| Created At | Thời gian tạo |

---

## ⚙️ Cấu hình

### Thay đổi phiên bản Chrome
Nếu Chrome của bạn khác phiên bản mặc định, sửa dòng:
```python
CHROME_VERSION_MAIN = 142  # Đổi thành phiên bản Chrome của bạn
```

### Thay đổi mật khẩu mặc định
Có thể thay đổi trong GUI hoặc sửa trực tiếp:
```python
DEFAULT_PASSWORD = "Matkhau123!@#"
```

---

## 🎨 Giao diện

- **Dark Mode** mặc định
- Hiệu ứng animation mượt mà
- Status bar với màu sắc trạng thái
- Console log với syntax highlighting
- Thống kê realtime (Success/Failed)

---

## ⚠️ Lưu ý quan trọng

1. **Sử dụng VPN** nếu IP của bạn bị giới hạn
2. **Không lạm dụng** - Có thể bị ban IP
3. **Kiểm tra Chrome version** trước khi chạy
4. **Backup file Excel** thường xuyên
5. Tool chỉ dành cho mục đích **học tập và nghiên cứu**

---

## 🐛 Xử lý lỗi thường gặp

| Lỗi | Giải pháp |
|-----|-----------|
| ChromeDriver version mismatch | Cập nhật `CHROME_VERSION_MAIN` |
| Operation timed out | Chuyển sang Network Mode: VPN/Slow |
| Email not received | Tool sẽ tự động resend OTP |
| Session expired | Đăng ký lại tài khoản |

---

## 📝 Changelog

### v2.0
- Giao diện GUI hiện đại với CustomTkinter
- Hỗ trợ multithreading
- Thêm module Checkout Capture
- Animation và hiệu ứng UI
- Cải thiện xử lý lỗi và retry logic

---

## 📄 License

Dự án này chỉ dành cho mục đích học tập và nghiên cứu. Tác giả không chịu trách nhiệm cho bất kỳ việc sử dụng sai mục đích nào.

---

## 👤 Tác giả

**tungd** - *Developer*

---

⭐ Nếu thấy hữu ích, hãy cho một star nhé!
