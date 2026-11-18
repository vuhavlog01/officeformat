# Script cai dat va dang ky OfficeFormat 1.0.0 vao Microsoft Word (Khong can quyen Admin)
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$addinDll = Join-Path $scriptDir 'VFormat.Addin.dll'

if (-not (Test-Path $addinDll)) {
    Write-Host "Khong tim thay file VFormat.Addin.dll tai: $addinDll" -ForegroundColor Red
    Exit 1
}

try {
    # Dong Microsoft Word neu dang mo de tranh xung dot file
    $wordProcesses = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
    if ($wordProcesses) {
        Write-Host "Dang dong Microsoft Word de cap nhat Add-in..." -ForegroundColor Yellow
        $wordProcesses | Stop-Process -Force
        Start-Sleep -Seconds 1
    }

    Write-Host "Dang nap assembly VFormat.Addin.dll..." -ForegroundColor Cyan
    $bytes = [System.IO.File]::ReadAllBytes($addinDll)
    $asm = [System.Reflection.Assembly]::LoadFrom($addinDll)
    $regType = $asm.GetType('VFormat.Addin.Registration.AddinRegistrar')
    $method = $regType.GetMethod('RegisterDualHive')
    $result = $method.Invoke($null, [object[]]@([string]$addinDll))

    if ($result) {
        Write-Host "============================================================" -ForegroundColor Green
        Write-Host "  CAI DAT VA DANG KY OFFICEFORMAT 1.0.0 THANH CONG!" -ForegroundColor Green
        Write-Host "  Trang thai ban quyen: BAN QUYEN PRO VINH VIEN (LIFETIME)" -ForegroundColor Green
        Write-Host "============================================================" -ForegroundColor Green
        Write-Host "Hay mo Microsoft Word de trai nghiem thanh cong cu OfficeFormat!" -ForegroundColor Yellow
    } else {
        Write-Host "Dang ky Add-in that bai. Vui long kiem tra lai quyen truy cap." -ForegroundColor Red
    }
} catch {
    Write-Host "Co loi xay ra trong qua trinh dang ky: $_" -ForegroundColor Red
    Exit 1
}
