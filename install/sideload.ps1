# HermesExcel — manifest'i Excel'e sideload eder (yonetici GEREKTIRMEZ).
# WEF "Developer" registry anahtarina manifest yolunu ekler; Excel bunu okur.
#
# Yerel (localhost) kurulum:   .\install\sideload.ps1
# Uzak (GX10) manifest:        .\install\sideload.ps1 -Manifest C:\yol\dist\manifest.xml
param([string]$Manifest)
$ErrorActionPreference = "Stop"

if (-not $Manifest) {
  $repo = Split-Path -Parent $PSScriptRoot
  $Manifest = Join-Path $repo "manifest.xml"
}
if (-not (Test-Path $Manifest)) { throw "manifest bulunamadi: $Manifest" }
$Manifest = (Resolve-Path $Manifest).Path

$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
New-ItemProperty -Path $key -Name $Manifest -Value $Manifest -PropertyType String -Force | Out-Null

Write-Host "Sideload kaydedildi (yonetici gerekmedi):" -ForegroundColor Green
Write-Host "  $Manifest" -ForegroundColor White
Write-Host ""
Write-Host "Sonraki adimlar:" -ForegroundColor Cyan
Write-Host "  1) Kopru calisiyor olsun (yerelde: npm start | GX10'da: node bridge/server.mjs)" -ForegroundColor White
Write-Host "  2) Excel'i tamamen kapatip yeniden acin." -ForegroundColor White
Write-Host "  3) Serit > Giris (Home) > 'Hermes' grubu > 'Hermes'i Ac'." -ForegroundColor White
Write-Host "     (Gorunmezse: Ekle > Eklentilerim > Gelistirici Eklentileri)" -ForegroundColor White
Write-Host "  4) Bir hucrede:  =HERMES.SOR(\"merhaba\")" -ForegroundColor White
Write-Host ""
Write-Host "Kaldirmak icin: install\uninstall.ps1" -ForegroundColor DarkGray