import os
import sys
import subprocess

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

print("========================================================")
print("   ĐẨY TOÀN BỘ MÃ NGUỒN LÊN HUGGING FACE SPACES (DOCKER)")
print("========================================================\n")
print("Ưu điểm của Hugging Face Spaces:")
print(" - Cấu hình miễn phí cực mạnh: 16 GB RAM + 2 vCPU + 50 GB Disk")
print(" - Không bị giới hạn 512MB RAM như Render Free")
print(" - LibreOffice chạy chuyển đổi file .doc mượt mà không lo quá tải\n")

print("Bước 1: Vào https://huggingface.co/new-space")
print(" - Space name: lich-cong-tac (hoặc tùy chọn)")
print(" - Space SDK: Chọn DOCKER (Blank)")
print(" - License: MIT hoặc Apache 2.0")
print(" - Visibility: Public\n")

space_url = input("Dán đường link Space Git của bạn vào đây:\n(Ví dụ: https://huggingface.co/spaces/Rishima/lich-cong-tac)\n> ").strip()

if not space_url:
    print("\n[!] Bạn chưa nhập link. Vui lòng chạy lại file.")
    input("\nNhấn Enter để thoát...")
    sys.exit(1)

# Chuẩn hóa URL nếu chưa có .git
if not space_url.endswith(".git"):
    space_git_url = space_url + ".git"
else:
    space_git_url = space_url

print(f"\n[*] Đang cấu hình remote: {space_git_url}...")
subprocess.run(["git", "branch", "-M", "main"])
subprocess.run(["git", "remote", "remove", "hf"], capture_output=True)
subprocess.run(["git", "remote", "add", "hf", space_git_url], capture_output=True)

print("[*] Đang gom mã nguồn và đóng gói bản cập nhật...")
subprocess.run(["git", "add", "."])
subprocess.run(["git", "commit", "-m", "Deploy to Hugging Face Spaces (Docker)"])

print("\n[*] Đang đẩy mã nguồn lên Hugging Face Spaces (git push -f hf main)...")
print("[Lưu ý] Nếu Git hỏi đăng nhập:")
print(" - Username: Tên tài khoản Hugging Face của bạn")
print(" - Password: Access Token của bạn (Tạo tại: https://huggingface.co/settings/tokens - quyền Write)\n")

res_push = subprocess.run(["git", "push", "-f", "hf", "main"])

if res_push.returncode == 0:
    print("\n========================================================")
    print("🎉 ĐẨY CODE LÊN HUGGING FACE SPACES THÀNH CÔNG 100%!")
    print(f"Bây giờ bạn truy cập: {space_url}")
    print("Chờ 1-2 phút để Docker tự động Build và ứng dụng sẽ chạy vĩnh viễn!")
    print("========================================================")
else:
    print("\n[!] Quá trình đẩy code gặp lỗi.")
    print("Bạn có thể kiểm tra lại Access Token (Write) tại: https://huggingface.co/settings/tokens")

input("\nNhấn Enter để đóng cửa sổ...")
