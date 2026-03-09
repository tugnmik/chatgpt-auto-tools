# ⚡ ChatGPT Auto Tools

Tool GUI đa luồng để tự động đăng ký tài khoản ChatGPT và lấy checkout link hàng loạt.

---

## 🚀 Cài đặt

**Yêu cầu:** Python 3.9+

```bash
pip install patchright customtkinter openpyxl colorama tls-client pyotp requests
python -m patchright install chromium
```

Chạy tool:

```bash
python chatgpt_auto_gui.pyw
```

---

## Tab 🚀 Registration – Đăng ký tài khoản tự động

Đăng ký hàng loạt tài khoản ChatGPT theo flow API-First (không click UI, nhanh hơn Selenium).

**Tính năng:**
- **Multithreading** – Nhiều luồng đăng ký song song (cấu hình số luồng tùy ý)
- **2 chế độ email:**
  - `TinyHost` – Tự tạo email tạm thời ngẫu nhiên qua `tinyhost.shop`
  - `OAuth2` – Dùng Gmail thật từ file `oauth2.xlsx` (DongVan API)
- **Tùy chọn sau đăng ký:**
  - `Get Checkout Link` – Tự động lấy link Plus / Business / cả hai ngay sau khi tạo xong
  - `Enable 2FA` – Tự động bật TOTP, lưu secret key vào Excel
- **Proxy support** – Hỗ trợ 3 format: `user:pass@host:port`, `host:port:user:pass`, `user:pass:host:port`
- Tên, ngày sinh ngẫu nhiên cho mỗi tài khoản
- Retry tự động 2 lần nếu thất bại

**Quy trình đăng ký:**
1. Lấy CSRF token
2. POST signin → lấy auth URL
3. Đăng ký qua `/api/accounts/user/register`
4. Chờ OTP email và validate
5. Tạo profile (tên + ngày sinh)
6. Xác nhận qua `/backend-api/accounts/check`
7. Lấy session token
8. *(Tùy chọn)* Setup 2FA + lấy checkout link

---

## Tab 💳 Checkout Capture – Lấy checkout link hàng loạt

Lấy checkout link từ các tài khoản đã lưu trong `chatgpt.xlsx` mà chưa có link.

**Tính năng:**
- Đọc Excel, tự phát hiện tài khoản chưa có link
- Hỗ trợ format mới (Session JSON) và format cũ (plain accessToken)
- Chọn loại link: **Plus**, **Business**, hoặc **cả hai**
- Chạy đa luồng, lưu kết quả trực tiếp vào Excel
- Random TLS fingerprint (Chrome/Firefox) chống bot

---

## 📁 Cấu trúc file dữ liệu

#### `chatgpt.xlsx` – Output tài khoản đã đăng ký
| Cột | Nội dung |
|-----|----------|
| A | `email:password` |
| B | Session JSON (từ `/api/auth/session`) |
| C | Plus Checkout URL |
| D | Business Checkout URL |
| E | 2FA Secret (TOTP) |

#### `oauth2.xlsx` – Input tài khoản Gmail (chế độ OAuth2)
| Cột A | `email|password|refresh_token|client_id` |
| Cột B | Status (`registered` = đã dùng, trống = chưa) |

#### `proxy_config.json` – Cấu hình proxy (tự tạo khi lưu trong GUI)
