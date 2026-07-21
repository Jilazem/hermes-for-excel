# HermesExcel — sideload kaydini kaldirir (yonetici gerekmez).
$ErrorActionPreference = "SilentlyContinue"

$repo = Split-Path -Parent $PSScriptRoot
$manifest = (Resolve-Path (Join-Path $repo "manifest.xml")).Path
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

if (Test-Path $key) {
  Remove-ItemProperty -Path $key -Name $manifest -ErrorAction SilentlyContinue
  Write-Host "Sideload kaydi kaldirildi: $manifest" -ForegroundColor Green
} else {
  Write-Host "Kayit bulunamadi." -ForegroundColor Yellow
}

# Eski SMB yontemi kaldiysa temizle
Remove-SmbShare -Name "HermesExcelCatalog" -Force -Confirm:$false -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{d3a1f4c2-8b6e-4f2a-9c1d-5e7a0b3c9d84}" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Excel'i yeniden baslatin." -ForegroundColor Green