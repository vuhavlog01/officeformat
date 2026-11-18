---
title: OfficeFormat Lịch Công Tác Tuần
emoji: 📅
colorFrom: red
colorTo: orange
sdk: docker
app_port: 7860
pinned: false
---

# OfficeFormat - Phần Mềm Chuẩn Hóa Văn Bản & Lập Lịch Công Tác Tuần

Bộ giải pháp công vụ văn phòng chuẩn Nghị định 30/2020/NĐ-CP và Hướng dẫn 05-HD/VPTW:
1. **Ứng dụng Web Lập Lịch Tuần:** Soạn thảo, xem trước A4 trực quan thời gian thực (Live Preview) và bóc tách / xuất file Word (`.docx` / `.doc`).
2. **Tiện ích Add-in Word (V-Format):** Tự động căn lề chuẩn quốc gia, định dạng Tiêu ngữ, Quốc hiệu, Bảng biểu, kiểm tra chính tả tiếng Việt.
3. **Kho Phần mềm Google Drive:** Đồng bộ trực tiếp bộ cài đặt và driver ký số từ Google Drive.

---

## 1. Ứng Dụng Web (Flask + Docker)
- **Tạo & xuất lịch tuần:** Tự động chia cột Buổi Sáng / Buổi Chiều, hợp nhất hàng (merge cells), nhận diện chữ đỏ sự kiện quan trọng.
- **Bóc tách Word cũ:** Hỗ trợ nạp lịch từ file Word cũ (`.doc` và `.docx`) thông qua LibreOffice Headless.
- **Xem trước trực quan A4:** Tỷ lệ chuẩn, xuất Word (.docx) chỉ với 1 click.

---

## 2. Tiện Ích Mở Rộng Microsoft Word (V-Format 2.2.9)
- **Bản quyền vĩnh viễn (Lifetime Pro):** Đầy đủ tính năng, không giới hạn.
- **Mã nguồn C# (.NET Framework 4.8 / SDK-style):**
  - `VFormat.Core`: Quy chuẩn thể thức (NĐ 30, HĐ 05, TT 68, HD29), từ điển kiểm tra chính tả.
  - `VFormat.WordEngine`: Tương tác Microsoft Word COM Interop, định dạng bảng, căn lề.
  - `VFormat.UI`: Giao diện WPF hiện đại (TaskPane, Hộp thoại chuẩn hóa, Thư viện mẫu).
  - `VFormat.Addin`: VSTO / COM Add-in trong Word, quản lý Ribbon tab.
- **Bộ cài đặt:** File cài tự động `OfficeFormat-1.0-Setup.exe` trong thư mục `downloads/`.
