@echo off
chcp 65001 >nul
title Build & Cap nhat OfficeFormat 1.0.0
cd /d "%~dp0"

echo ============================================================
echo   DANG BUILD SOLUTION OFFICEFORMAT 1.0.0 (RELEASE)...
echo ============================================================
cd src
dotnet build -c Release
if %errorlevel% neq 0 (
    echo.
    echo [LOI] Build that bai! Vui long kiem tra lai code.
    pause
    exit /b %errorlevel%
)
cd ..

echo.
echo ============================================================
echo   DANG DONG WORD VA CAP NHAT DLL VAO THU MUC CHAY...
echo ============================================================
taskkill /f /im WINWORD.EXE 2>nul
timeout /t 1 /nobreak >nul

copy /y src\VFormat.Core\bin\Release\net48\VFormat.Core.dll App_VFormat_2.2.9\ >nul
copy /y src\VFormat.WordEngine\bin\Release\net48\VFormat.WordEngine.dll App_VFormat_2.2.9\ >nul
copy /y src\VFormat.License\bin\Release\net48\VFormat.License.dll App_VFormat_2.2.9\ >nul
copy /y src\VFormat.UI\bin\Release\net48\VFormat.UI.dll App_VFormat_2.2.9\ >nul
copy /y src\VFormat.Addin\bin\Release\net48\VFormat.Addin.dll App_VFormat_2.2.9\ >nul
copy /y src\VFormat.Addin\bin\Release\net48\PdfSharp*.dll App_VFormat_2.2.9\ >nul
copy /y src\VFormat.Addin\VFormat.Addin.Ribbon.customUI14.xml App_VFormat_2.2.9\ >nul

echo.
echo ============================================================
echo   BUILD VA CAP NHAT OFFICEFORMAT 1.0.0 THANH CONG!
echo ============================================================
echo Ban co the mo Microsoft Word de xem ket qua thay doi.
echo.
pause
