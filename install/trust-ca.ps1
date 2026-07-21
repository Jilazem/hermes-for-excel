# HermesExcel (Win11) — GX10'da uretilen CA'yi guvenilir koke ekler.
# Boylece Excel WebView, https://<gx10>:8787 sertifikasina guvenir.
# Yonetici GEREKMEZ (CurrentUser\Root). Bir kez Windows guvenlik uyarisi cikar -> Evet.
#
# Kullanim:  .\install\trust-ca.ps1 -CaPath C:\yol\certs\ca.crt
param([Parameter(Mandatory=$true)][string]$CaPath)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $CaPath)) { throw "ca.crt bulunamadi: $CaPath" }
$CaPath = (Resolve-Path $CaPath).Path

# X509Store ile ekle — Import-Certificate'in gerektirdigi interaktif UI'yi atlar
# (kullanici CurrentUser deposuna ekliyor; yonetici gerekmez).
$c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CaPath)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($c)
$store.Close()
Write-Host "CA guvenilir koke eklendi:" -ForegroundColor Green
Write-Host ("  Subject : {0}" -f $c.Subject)
Write-Host ("  Gecerli : {0} -> {1}" -f $c.NotBefore, $c.NotAfter)
Write-Host ""
Write-Host "Kaldirmak icin: sertifikayi Cert:\CurrentUser\Root altindan silin (certmgr.msc)." -ForegroundColor DarkGray