# AI Reply (Claude) – Outlook add-in

Khung bên phải trong Outlook: Claude đọc email, soạn bản trả lời; bạn sửa trong khung,
bấm **Trả lời / Trả lời tất cả** để mở cửa sổ reply của Outlook có sẵn nội dung, kiểm tra lần cuối rồi **Send**.

## Thành phần
| File | Vai trò |
|---|---|
| `manifest.xml` | Khai báo add-in với Outlook (nút trên Ribbon, đường dẫn tới trang) |
| `taskpane.html / .css / .js` | Giao diện khung bên phải + gọi Claude API |
| `commands.html` | Trang kỹ thuật bắt buộc của manifest |
| `assets/` | Icon |

## Bước 1 – Đưa các file lên web (HTTPS)
Outlook chỉ nạp add-in từ địa chỉ **https**. Cách miễn phí: GitHub Pages.

1. Tạo repo mới trên GitHub, ví dụ `outlook-ai-reply`.
2. Upload toàn bộ file (giữ nguyên thư mục `assets`).
3. Repo → **Settings → Pages** → Source: *Deploy from a branch* → `main` / `(root)` → Save.
4. Sau 1–2 phút, mở thử `https://<tên-github>.github.io/outlook-ai-reply/taskpane.html` để chắc chắn trang đã chạy.

> Repo public chỉ lộ mã nguồn, **không** lộ API key. Key được lưu trong hộp thư của bạn, không nằm trong file nào.
> Có thể dùng bất kỳ hosting HTTPS nào khác (Azure Static Web Apps, IIS nội bộ có SSL...).

## Bước 2 – Sửa manifest
Mở `manifest.xml` bằng Notepad → **Ctrl + H**:
- Find: `https://YOUR-HOST`
- Replace: `https://<tên-github>.github.io/outlook-ai-reply` (không có dấu `/` ở cuối)
- **Replace All** → lưu file.

## Bước 3 – Cài vào Outlook
1. Mở trình duyệt, đăng nhập Outlook on the web bằng tài khoản công ty, vào: **https://aka.ms/olksideload**
2. Chọn **My add-ins → Custom Addins → Add a custom add-in → Add from File...** → chọn `manifest.xml` → **Install**.
3. Tắt hẳn Outlook 2016 rồi mở lại.

(Cách khác trong Outlook 2016: **File → Manage Add-ins**, nó cũng mở trang trên.)

## Bước 4 – Dùng
1. Chọn một email → tab **Home** → nhóm **AI Reply** → nút **Soạn trả lời AI**.
2. Lần đầu: nhập Claude API key ở ⚙ Cài đặt → **Lưu**. Có thể điền thêm "Thông tin về bạn" để AI viết đúng vai.
3. (Tuỳ chọn) gõ ý chính muốn trả lời → **✨ Soạn trả lời**.
4. Sửa trực tiếp trong khung, hoặc gõ yêu cầu ("ngắn hơn", "lịch sự hơn") → **Chỉnh**.
5. Bấm **Trả lời tất cả** / **Trả lời** → cửa sổ reply mở ra với nội dung sẵn → kiểm tra → **Send**.

## Lỗi thường gặp
| Hiện tượng | Xử lý |
|---|---|
| Không thấy nút trên Ribbon | Khởi động lại Outlook; kiểm tra Ribbon đang hiện đủ (Ctrl + F1); đúng là đang **đọc** email (không phải đang soạn) |
| Khung trắng / báo không tải được | URL trong manifest sai, hoặc trang GitHub Pages chưa chạy – mở thử link taskpane.html trên trình duyệt |
| "Không kết nối được tới api.anthropic.com" | Tường lửa/proxy công ty chặn – mở `api.anthropic.com` (HTTPS 443) |
| "API key không hợp lệ" | Mở ⚙, dán lại key |
| Mục Custom Addins không có | Admin M365 đã tắt cài add-in tự do. Admin có thể triển khai cho cả công ty: **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps** → chọn manifest |
| Thư trả lời có 2 lần tên ký | Xoá ô "Ký tên cuối thư" trong ⚙ (để Outlook tự chèn chữ ký) |

## Lưu ý
- Nội dung email được gửi tới Anthropic qua API để soạn trả lời. Cân nhắc với email mật theo chính sách công ty.
- Outlook 2016 không cho ghim khung: mỗi email mới cần bấm nút lại. Outlook bản mới (Microsoft 365) có biểu tượng ghim 📌 ở góc khung.
- Đổi model trong ⚙: Sonnet (mặc định, cân bằng), Haiku (nhanh, rẻ), Opus (chất lượng cao nhất).
