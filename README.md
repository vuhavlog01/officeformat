# V-FORMAT 2.2.9 (Bản Quyền Vĩnh Viễn & Toàn Bộ Mã Nguồn)

## 1. Cấu trúc thư mục tại `D:\PM\ADD_IN\v-format`

- **`App_VFormat_2.2.9\`**: 
  - Bản Add-in hoàn chỉnh đã được **bẻ khóa / kích hoạt bản quyền Vĩnh viễn (Lifetime Pro)**.
  - Không còn thông báo hay cửa sổ yêu cầu kích hoạt bản quyền.
  - Đã tích hợp sẵn toàn bộ mẫu văn bản (`templates/`), từ điển tiếng Việt (`SpellCheck/Data/`), cấu hình (`user_presets.json`).
  - **`CaiDat_VFormat.bat`**: Nhấp đúp để cài đặt và đăng ký V-Format vào Microsoft Word chỉ với 1 cú click (không cần quyền Admin).
  - **`GoCaiDat_VFormat.bat`**: Nhấp đúp nếu muốn gỡ bỏ V-Format khỏi Word.

- **`src\`**: 
  - Toàn bộ mã nguồn C# (.NET Framework 4.8 / SDK-style) của toàn bộ 5 module ứng dụng V-Format 2.2.9:
    1. `VFormat.Core`: Quy chuẩn thể thức (Nghị định 30, Hướng dẫn 05, Thông tư 68, HD29), kiểm tra chính tả tiếng Việt.
    2. `VFormat.WordEngine`: Tương tác với Microsoft Word COM Interop, định dạng bảng, căn lề, tiêu ngữ, chữ ký.
    3. `VFormat.License`: Cơ chế bản quyền (đã được cấu hình mở khóa vĩnh viễn vĩnh cửu).
    4. `VFormat.UI`: Giao diện WPF hiện đại (TaskPane, Hộp thoại chuẩn hóa, Kiểm tra chính tả, Thư viện mẫu).
    5. `VFormat.Addin`: Điểm khởi chạy VSTO / COM Add-in trong Word, quản lý Ribbon tab và Ribbon XML.
  - File Solution **`V-Format.slnx`**: Mở trực tiếp trong Visual Studio 2022/2026 hoặc VS Code để xem mã nguồn và build dự án.

- **`Goc_Installer_2.2.9\`**:
  - Lưu giữ file nén `.zip` và bộ cài đặt gốc `V-Format-2.2.9-Setup.exe` để đối chiếu khi cần.

- **`WordAddIn1\`**:
  - Dự án VSTO trước đây của bạn (được giữ nguyên vẹn).

---

## 2. Hướng dẫn sử dụng nhanh

1. Đóng Microsoft Word (nếu đang mở).
2. Vào thư mục `D:\PM\ADD_IN\v-format\App_VFormat_2.2.9\`.
3. Chạy file **`CaiDat_VFormat.bat`**.
4. Mở Word lên, bạn sẽ thấy ngay tab **V-Format** trên thanh Ribbon với đầy đủ tất cả các tính năng Pro hoạt động không giới hạn!
