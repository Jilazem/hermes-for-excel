# Uzak dağıtım: Köprü GX10'da, Excel Win11'de

Köprü GX10 (Linux/ARM) üzerinde çalışır; Excel Win11'den ona LAN/VPN üzerinden
bağlanır. GX10 aynı zamanda Hermes ajanını (API Server :8642) barındırıyorsa
köprü ona `localhost` üzerinden konuşur — ideal.

```
Win11 Excel ──HTTPS──▶ GX10:8787 (köprü) ──localhost──▶ GX10:8642 (Hermes API Server)
```

## GX10 tarafı (Linux)

1. **Repoyu GX10'a getirin** (Google Drive senkronu, `git clone`, ya da `scp`/`rsync`).
2. **Adresi belirleyin** — Win11'in ulaşacağı ad/IP:
   - LAN IP: `ip -4 addr` veya `hostname -I`
   - Tailscale: `tailscale status` (MagicDNS adı, `*.ts.net`)
3. **Sertifika + manifest üret** (tek komut):
   ```bash
   # LAN IP:
   node install/configure-remote.mjs --host 192.168.1.50
   # veya mDNS ad:
   node install/configure-remote.mjs --host gx10.local
   # veya Tailscale (gerçek sertifika, Win11'de CA güven GEREKMEZ):
   node install/configure-remote.mjs --host gx10.tailXXXX.ts.net --tailscale
   ```
   Çıktı: `certs/server.crt`+`server.key`, (özel modda) `certs/ca.crt`, ve
   `dist/manifest.xml`.
4. **`.env`** (configure-remote.mjs çıktısındaki değerlerle) — özellikle:
   ```
   HERMES_HOST=0.0.0.0
   HERMES_PUBLIC_HOST=<adres>
   HERMES_TLS_CERT=./certs/server.crt
   HERMES_TLS_KEY=./certs/server.key
   HERMES_PROVIDER=gateway
   HERMES_GATEWAY_URL=http://127.0.0.1:8642/v1/chat/completions
   HERMES_API_KEY=<API_SERVER_KEY>
   ```
5. **Köprüyü başlat**:
   ```bash
   node bridge/server.mjs
   ```
   Kalıcı çalışması için systemd servisi:
   ```ini
   # /etc/systemd/system/hermes-excel.service
   [Unit]
   Description=HermesExcel bridge
   After=network.target
   [Service]
   WorkingDirectory=/opt/ExcelPluginHermes
   ExecStart=/usr/bin/node bridge/server.mjs
   Restart=always
   [Install]
   WantedBy=multi-user.target
   ```
   `sudo systemctl enable --now hermes-excel`
6. **Güvenlik duvarı**: GX10'da 8787/tcp'yi LAN'a açın
   (`sudo ufw allow from 192.168.1.0/24 to any port 8787 proto tcp`).

## Win11 tarafı (Excel istemcisi)

GX10'dan Win11'e iki dosyayı taşıyın: `certs/ca.crt` (Tailscale değilse) ve
`dist/manifest.xml`. (Google Drive senkronuysa zaten oradalar.)

```powershell
# 1) CA'ya güven (Tailscale ise ATLA) — bir kez "Evet" der
.\install\trust-ca.ps1 -CaPath <yol>\ca.crt
# 2) GX10 manifest'ini sideload et
.\install\sideload.ps1 -Manifest <yol>\manifest.xml
```

Sonra Excel'i tamamen kapatıp açın → **Giriş > Hermes > Hermes'i Aç**, hücrede
`=HERMES.SOR("...")`.

## Doğrulama

Win11'den: tarayıcıda `https://<adres>:8787/api/health` açılıyor ve sertifika
uyarısı vermiyorsa hazırsınız. Uyarı veriyorsa CA güven adımı eksiktir (ya da
SAN adresi Win11'in kullandığı adresle birebir aynı değildir).

## İpuçları

- **Adres birebir aynı olmalı**: sertifika SAN'ı, manifest URL'i ve Win11'in
  tarayıcıya yazdığı adres aynı olmalı (IP ise IP, ad ise ad).
- **Token**: LAN'a açtığınız için `.env`'de `HERMES_BRIDGE_TOKEN` verin;
  `addin/functions.js` ve `addin/taskpane.js` içindeki `TOKEN` sabitine de aynı
  değeri yazıp `dist/manifest.xml`'i yeniden üretin.
- Tailscale en sorunsuz yol: gerçek sertifika + sabit MagicDNS adı.
