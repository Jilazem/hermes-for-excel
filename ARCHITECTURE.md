# HermesExcel — Mimari (Architecture)

Excel ile Hermes ajanı arasında **iki yönlü** köprü kuran bir Office.js eklentisi.

> Amaç: Excel'de çalışırken Hermes'e ulaşmak. Bir hücreden
> `=HERMES.SOR("tapu bilgilerini getir")` yazınca Hermes, kendi MCP'lerini
> (bilirkişi orchestrator, emsal, parsel-konum/tapu, hesap-motoru, rapor
> pipeline...) kullanarak cevabı hücreye yazar; ayrıca yan panelden sohbet
> ederek Hermes'in çalışma kitabını **tam olarak kontrol etmesini** sağlar.

## Referans

İki GitHub projesi temel alındı:

- [`lEWFkRAD/hermes-excel-sidecar`](https://github.com/lEWFkRAD/hermes-excel-sidecar)
  — "task pane + Node köprüsü → Hermes gateway" mimarisi, action sözleşmesi ve
  formül rebase fikri. (Özel fonksiyonu yok.)
- [`tonbistudio/hermes-office`](https://github.com/tonbistudio/hermes-office)
  — resmi **Hermes API Server** sözleşmesini (`API_SERVER_ENABLED`,
  `API_SERVER_KEY`, OpenAI-uyumlu `/v1/chat/completions` @ `:8642`,
  `/v1/health`), idempotency-key ile sunucu-tarafı dedup ve genel amaçlı
  fonksiyonları (CLASSIFY/EXTRACT/SUMMARIZE/FORMULA_HELP →
  SINIFLA/CIKAR/OZETLE/FORMUL). Onların Caddy proxy'sinin işini (TLS + anahtar
  enjeksiyonu) bizde **Node köprüsü** üstlenir ve ek olarak ajan→Excel push
  kanalını sağlar.

Bu tasarım her ikisinin üzerine **hücre içi özel fonksiyonlar** (`=HERMES.*`) ve
**ajan → Excel komut kanalı** (SSE) ekler.

## Bileşenler

```
┌─ Excel (Office.js) ───────────────┐        ┌─ Köprü (Node, :8787) ────────────┐       ┌─ Hermes Agent ───┐
│  Özel fonksiyonlar (functions.js) │ HTTPS  │  GET  /api/health                │       │  gateway /v1     │
│    =HERMES.SOR / .EMSAL / .TAPU   │──────▶ │  POST /api/fn     (hücre fonk.)  │─────▶ │  + MCP yığını:   │
│    .HESAP / .RAPOR / .HAFIZA/.MCP │        │  POST /api/chat   (sohbet)       │  MCP  │   orchestrator   │
│  Task pane (taskpane.js) sohbet   │◀────── │  POST /api/push   (ajan→Excel)   │◀───── │   emsal_mcp      │
│    + gelen action'ları uygular    │  SSE   │  GET  /api/commands (SSE akışı)  │       │   parsel-konum   │
│  Ribbon (commands.js)             │        │  hermes-client.mjs (provider)    │       │   hesap-motoru   │
└────────────────────────────────────┘        └──────────────────────────────────┘       │   rapor-pipeline │
                                                                                          └──────────────────┘
```

Köprü yalnızca `127.0.0.1`'e bağlanır. Office eklentileri HTTPS zorunlu kıldığı
için köprü, `office-addin-dev-certs` ile üretilen güvenilir yerel sertifikayı
kullanır (`https://localhost:8787`).

## Veri akışı

### 1) Excel → Hermes (hücre içi)

`=HERMES.SOR("tapu bilgilerini getir")`

1. Özel fonksiyon `POST /api/fn` çağırır: `{ fn:"SOR", args:["tapu..."], cell, sheet }`.
2. Köprü, fonksiyona göre bir **sistem promptu + niyet** hazırlar (bkz.
   `config/mcp-map.json`) ve Hermes provider'ına iletir.
3. Hermes ajanı ilgili MCP'yi çağırır (ör. `parsel-konum`), sonucu döndürür.
4. Köprü düz metin / tek değer döndürür; hücreye yazılır. Sonuç TTL'li
   önbelleğe alınır (Excel'in tekrar-hesaplamalarında ajanı yormamak için).

### 2) Excel → Hermes (sohbet paneli)

Task pane, çalışma kitabı bağlamını (aktif sayfa, kullanılan aralıklar, seçim)
ve mesaj geçmişini `POST /api/chat`'e gönderir. Yanıt:
`{ message, actions[] }`. `actions[]` içindeki her komut **panel tarafından**
Office.js ile uygulanır — ajan araçları çalışma kitabına doğrudan dokunmaz
(savunma derinliği).

### 3) Hermes → Excel (ajan çalışma kitabını sürer)

Hermes, herhangi bir noktada (bir MCP aracı üzerinden) `POST /api/push`'a bir
action listesi gönderebilir. Köprü bunu açık SSE bağlantısı (`/api/commands`)
üzerinden bağlı panele iletir; panel action'ları uygular. Böylece Hermes,
sohbet dışında da tabloyu doldurabilir, biçimlendirebilir, sayfa oluşturabilir.

> Hermes tarafında eklenmesi gereken MCP aracı sözleşmesi: `docs/HERMES_MCP_TOOL.md`.

## Action sözleşmesi (Hermes → panel)

| type | alanlar | açıklama |
|---|---|---|
| `write_cells` | `start_cell`, `values[][]`, `allow_overwrite?`, `auto_format?` | Hücrelere yaz |
| `create_sheet` | `name`, `values[][]?` | Yeni sayfa + tablo |
| `format_cells` | `range`, `bold?`, `fill?`, `font_color?`, `number_format?`, `align?` | Biçim |
| `set_number_format` | `range`, `format` | Sayı formatı |
| `clear_range` | `range` | Temizle |
| `read_range` | `range`, `reason?` | Panelden değer oku (okuma döngüsü) |
| `message` | `text` | Panelde bilgi mesajı |

Formüller **tablo-yerel** (sanki tablonun sol-üstü A1'miş gibi) yazılır; köprü
bunları hedef `start_cell`'e göre yeniden konumlar (`actions.mjs`).

## Provider soyutlaması (`bridge/hermes-client.mjs`)

`HERMES_PROVIDER` ile seçilir:

- `gateway` (varsayılan) — OpenAI-uyumlu Hermes gateway'e `POST`
  (`HERMES_GATEWAY_URL`, `Authorization: Bearer HERMES_API_KEY`). Ajan MCP'leri
  kendi tarafında çağırır.
- `anthropic` — doğrudan Anthropic Messages API (Claude). Token, `hermes-auth-monitor`
  tarafından yönetilen `HERMES_ANTHROPIC_TOKEN`'dan okunur.
- `mock` — çevrimdışı test. Ajan olmadan tüm eklenti uçtan uca çalışır.

## Güvenlik

- Köprü yalnız `127.0.0.1`; opsiyonel `HERMES_BRIDGE_TOKEN` ile her `/api/*`
  çağrısı korunur.
- Ajan araçları çalışma kitabına doğrudan yazmaz; yalnız yapılandırılmış
  `actions[]` panel tarafından uygulanır.
- Fonksiyon önbelleği kişisel veriyi diske yazmaz (bellek içi, TTL'li).

## Dosya haritası

```
manifest.xml            Office manifesti (task pane + custom functions)
bridge/
  server.mjs            HTTPS statik + API sunucusu (0 bağımlılık)
  config.mjs            .env okuyucu + ayarlar
  hermes-client.mjs     provider soyutlaması (gateway/anthropic/mock)
  actions.mjs           action normalizasyonu + formül yeniden konumlama
  command-bus.mjs       SSE komut veri yolu (ajan → panel)
addin/
  functions.js          =HERMES.* özel fonksiyonları
  functions.json        özel fonksiyon meta verisi
  functions.html        özel fonksiyon runtime yükleyici
  taskpane.html/.css/.js sohbet paneli + action uygulayıcı
  commands.html/.js     ribbon komutları
config/mcp-map.json     HERMES.* → MCP sunucu/araç eşlemesi
install/*.ps1           sertifika + sideload betikleri
docs/HERMES_MCP_TOOL.md Hermes tarafına eklenecek MCP aracı sözleşmesi
```
