@echo off
chcp 65001 >nul
title Cai Dat OfficeFormat 1.0.0
cd /d "%~dp0"
echo ============================================================
echo   DANG CAI DAT VA DANG KY OFFICEFORMAT 1.0.0 VAO WORD
echo ============================================================
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0CaiDat_VFormat.ps1"
echo.
pause
