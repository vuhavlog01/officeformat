@echo off
chcp 65001 >nul
title Go Cai Dat OfficeFormat
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0GoCaiDat_VFormat.ps1"
pause
