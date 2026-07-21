# HermesExcel — yerel HTTPS geliştirme sertifikasını kurar ve güvenilir yapar.
# Office eklentileri HTTPS zorunlu kılar; localhost için güvenilir bir CA gerekir.
$ErrorActionPreference = "Stop"
Write-Host "Office geliştirme sertifikaları kuruluyor (office-addin-dev-certs)..." -ForegroundColor Cyan
npx --yes office-addin-dev-certs install
Write-Host ""
Write-Host "Doğrulama:" -ForegroundColor Cyan
npx --yes office-addin-dev-certs verify
Write-Host ""
Write-Host "Sertifikalar ~/.office-addin-dev-certs klasorune kuruldu." -ForegroundColor Green
Write-Host "Kopru bunlari otomatik bulur. Simdi: npm start" -ForegroundColor Green
