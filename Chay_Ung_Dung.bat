@echo off
chcp 65001 >nul
title ỨNG DỤNG LẬP & XUẤT LỊCH CÔNG TÁC TUẦN
color 0b
echo =======================================================================
echo          PHẦN MỀM LẬP & XUẤT LỊCH CÔNG TÁC TUẦN (CHUẨN ĐẢNG 2026)
echo =======================================================================
echo.
echo Đang khởi động máy chủ Web cục bộ...
cd /d "D:\PM\LichCongTacTuan"

:: Mở trình duyệt sau 2 giây
start "" http://localhost:5000

:: Chạy ứng dụng Flask
python server.py

pause
