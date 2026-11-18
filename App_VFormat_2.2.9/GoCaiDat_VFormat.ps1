# Script go cai dat OfficeFormat khoi Microsoft Word
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$addinDll = Join-Path $scriptDir 'VFormat.Addin.dll'

try {
    # Dong Microsoft Word neu dang mo
    $wordProcesses = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
    if ($wordProcesses) {
        Write-Host "Dang dong Microsoft Word..." -ForegroundColor Yellow
        $wordProcesses | Stop-Process -Force
        Start-Sleep -Seconds 1
    }

    if (Test-Path $addinDll) {
        $asm = [System.Reflection.Assembly]::LoadFrom($addinDll)
        $regType = $asm.GetType('VFormat.Addin.Registration.AddinRegistrar')
        $method = $regType.GetMethod('UnregisterPerUser')
        $result = $method.Invoke($null, $null)
    }

    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  DA GO BO OFFICEFORMAT KHOI MICROSOFT WORD THANH CONG!" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
} catch {
    Write-Host "Co loi xay ra: $_" -ForegroundColor Red
}
