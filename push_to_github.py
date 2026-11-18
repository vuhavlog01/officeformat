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
print("   ĐẨY TOÀN BỘ MÃ NGUỒN LÊN GITHUB ĐỂ DEPLOY RENDER.COM")
print("========================================================\n")

repo_url = input("Dán đường link GitHub Repository của bạn vào đây:\n(Ví dụ: https://github.com/username/officeformat-lich-tuan.git)\n> ").strip()

if not repo_url:
    print("\n[!] Bạn chưa nhập link. Vui lòng chạy lại file.")
    input("\nNhấn Enter để thoát...")
    sys.exit(1)

print(f"\n[*] Đang kết nối tới kho: {repo_url}...")
subprocess.run(["git", "branch", "-M", "main"])
subprocess.run(["git", "remote", "remove", "origin"], capture_output=True)
res_remote = subprocess.run(["git", "remote", "add", "origin", repo_url], capture_output=True)

print("[*] Đang gom mã nguồn và đóng gói bản cập nhật...")
subprocess.run(["git", "add", "."])
subprocess.run(["git", "commit", "-m", "Deploy to Render.com"])

print("\n[*] Đang đẩy mã nguồn lên GitHub (git push -u origin main)...")
res_push = subprocess.run(["git", "push", "-u", "origin", "main"])

if res_push.returncode == 0:
    print("\n========================================================")
    print("🎉 ĐẨY CODE LÊN GITHUB THÀNH CÔNG 100%!")
    print("Bây giờ bạn chỉ cần vào Render.com để liên kết và chạy tự động!")
    print("========================================================")
else:
    print("\n[!] Đẩy code chưa hoàn tất. Bạn có thể kiểm tra lại quyền truy cập hoặc Personal Access Token (PAT) của GitHub.")

input("\nNhấn Enter để đóng cửa sổ...")
