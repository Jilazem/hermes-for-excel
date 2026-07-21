# Çalışan Mimari + Zorlu Dersler (Working Architecture & Gotchas)

Bu belge, HermesExcel'in **fiilen çalışan** kurulumunu ve kuruluma giden yolda
bulunan 6 kritik engeli anlatır. Aynı hataları tekrar yaşamamak için okuyun.

## Nihai mimari

```
┌─ Win11 (Excel istemcisi) ───────────────────┐        ┌─ GX10 (192.168.1.50, Linux) ─────────┐
│  Excel add-in (localhost'tan yüklenir)       │        │  bridge/server.mjs  :8799  (HTTPS)   │
│    =HERMES.SOR / .TAPU / .EMSAL ...          │        │    ↳ SAN'lı sertifika (≤398 gün)     │
│    + sohbet paneli                           │        │  Hermes API Server  :8646 (gateway)  │
│                                              │        │    ↳ ajan + MCP'ler (bilirkişi...)   │
│  bridge/proxy.mjs  (localhost:8799, dualstack)│        └───────────────────────────────────────┘
│    • add-in STATİK dosyalarını YEREL sunar    │  HTTPS (LAN)          ▲
│    • /api/*'i GX10'a iletir  ─────────────────┼──────────────────────┘
│    • Startup'ta otomatik (gizli süreç)        │
│                                               │
│  Masaüstü kısayolu "Hermes Excel"             │
│    → office-addin-debugging start (sideload)  │
└───────────────────────────────────────────────┘
```

**Neden proxy?** Office add-in WebView2'si bir **LAN IP'sine (192.168.1.50)
erişemez** — yalnız `localhost` veya public adres. Bu yüzden Win11'de localhost
proxy: add-in dosyalarını yerel sunar, gerçek API çağrılarını GX10'a iletir.

## Kurulum (özet)

1. **GX10:** `bridge/server.mjs` çalışır (systemd), `.env` → `HERMES_HOST=0.0.0.0`,
   `HERMES_PUBLIC_HOST=192.168.1.50`, `HERMES_TLS_CERT/KEY=certs/server.*`,
   `HERMES_PROVIDER=gateway`, `HERMES_GATEWAY_URL=http://127.0.0.1:8646/...`.
   Sertifika `install/configure-remote.mjs --host 192.168.1.50` ile (≤398 gün).
2. **Win11 — bir kez:**
   - `npm run cert` (localhost dev sertifikası — Office bunu ister).
   - GX10 CA'sına güven: **`certutil -addstore -f Root C:\HermesExcel\ca.crt`**
     (yönetici — `LocalMachine\Root`, sadece CurrentUser yetmiyor).
   - Add-in dosyalarını `C:\HermesExcel\www`'a, proxy'yi `C:\HermesExcel`'e kopyala.
   - Proxy launcher'ı Startup'a koy (`install/proxy-launcher.vbs` → ASCII yol!).
   - Masaüstü kısayolu: `install/baslat.bat`'i çağıran gizli VBS.
3. **Kullanım:** Masaüstündeki **"Hermes Excel"** kısayoluna çift tıkla →
   Excel Hermes yüklü açılır → `=HERMES.SOR("...")` ve panel çalışır.

## 6 KRİTİK ENGEL (hepsi çözüldü)

### 1. Sertifika geçerliliği > 398 gün → Chromium reddediyor
`ERR_CERT_VALIDITY_TOO_LONG`. `curl`/.NET bu kuralı uygulamaz, **WebView2
(Chromium) uygular**. Leaf sertifika **≤398 gün** olmalı. `configure-remote.mjs`
artık 397 gün kullanıyor. (CA uzun olabilir; leaf kısa.)

### 2. CA yalnız `CurrentUser\Root` → WebView2 güvenmiyor
Office WebView2 sandbox'ı **makine deposundaki** kök CA'ları ister. CA'yı
**`LocalMachine\Root`**'a ekleyin (yönetici):
`certutil -addstore -f Root ca.crt`. Sadece kullanıcı deposu yetmez.

### 3. Office add-in LAN IP'sine erişemez
WebView2 sandbox'ı `192.168.1.50` gibi yerel-ağ IP'lerine erişemez (public
add-in'ler — Claude — çalışır ama LAN olanlar "gerekli kaynağı indiremedi"
verir). Çözüm: **Win11 localhost proxy** (`bridge/proxy.mjs`).

### 4. `localhost` önce IPv6'ya (`::1`) çözülür → proxy dualstack olmalı
Proxy yalnız `127.0.0.1` (IPv4) dinlerse, WebView2'nin `https://localhost:8799`
isteği `[::1]`'e gider ve **hiç ulaşmaz** (log boş, panel "Yükleniyor"da takılır).
Proxy **hem `127.0.0.1` hem `::1`** dinler.

### 5. `CustomFunctions` ExtensionPoint yanlış yerde → `=HERMES.SOR` = `#AD?`
Manifest'te CustomFunctions **`<AllFormFactors>`** altında olmalı,
`<DesktopFormFactor>` altında DEĞİL. Aksi halde Office
`Skipped unrecognized XML element ... CustomFunctions` deyip atlar; fonksiyonlar
kaydolmaz (`#AD?` / `#NAME?`). Not: `office-addin-manifest validate` bunu
YAKALAMAZ; hata yalnız çalışma-zamanı logunda (`OfficeAddins.log.txt`) görünür.

### 6. Office custom-functions'ı agresif önbelleğe alır (Wef)
functions.js değişince Office eski sürümü Wef cache'inden kullanmaya devam eder.
Güncellemeyi zorlamak için: Excel kapalıyken
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` silin + yeniden sideload
(`office-addin-debugging start`). Ayrıca custom-functions runtime'ında
`self.location.origin` güvenilmez → `functions.js`'te köprü adresi **sabit**
(`https://localhost:8799`).

## Ek notlar

- **Sideload:** Normal Excel açılışı bu M365 sürümünde dev-sideload'u güvenilir
  yüklemedi; **`office-addin-debugging start`** akışı yüklüyor → masaüstü kısayolu
  bunu sarar. (Kalıcı "her açılışta" için: add-in'i public HTTPS'e host et.)
- **VBS launcher ASCII yol:** `G:\Drive'ım` içindeki Türkçe `ı` wscript'i bozar →
  proxy + dosyalar `C:\HermesExcel` (ASCII) altında.
- **Teşhis logları:** proxy → `%TEMP%\hermes-proxy.log`; Office →
  `%TEMP%\OfficeAddins.log.txt` (office-addin-debugging açar).
- **Güvenlik:** `.env` (API anahtarı, token) ve `certs/` (özel anahtarlar)
  git'e girmez (bkz. `.gitignore`). Google Drive senkronuysa buluta gider — dikkat.
