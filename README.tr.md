# Hermes for Excel

Excel'den **Hermes ajanına** ulaşmanızı sağlayan Office eklentisi. İki yönlü:

- **Hücreden Hermes'e:** `=HERMES.SOR("tapu bilgilerini getir")` yazın; Hermes
  kendi MCP'lerini (bilirkişi orchestrator, emsal, parsel-konum/tapu,
  hesap-motoru, rapor pipeline, genel-hafıza) kullanıp cevabı **hücreye** yazsın.
- **Hermes'ten Excel'e:** Yan sohbet panelinden "şu tabloyu doldur / değerleme
  sayfası oluştur" deyin; Hermes çalışma kitabını **doğrudan sürsün** (hücre yaz,
  biçimlendir, sayfa oluştur).

> Mimarinin ayrıntısı: [ARCHITECTURE.md](ARCHITECTURE.md).
> İlham: [`lEWFkRAD/hermes-excel-sidecar`](https://github.com/lEWFkRAD/hermes-excel-sidecar)
> (task pane + Node köprüsü). Buraya ek olarak **hücre içi özel fonksiyonlar** ve
> **ajan → Excel komut kanalı** eklenmiştir.

---

## Hızlı başlangıç (3 adım)

Gereksinim: **Excel masaüstü (Windows)**, **Node.js 18+**.

```powershell
# 1) Bağımlılık + yerel HTTPS sertifikası (Office HTTPS ister)
npm install
npm run cert                     #  = office-addin-dev-certs install

# 2) Köprüyü başlat  (varsayılan: mock — ajan olmadan test için)
copy .env.example .env
npm start                        #  https://localhost:8787

# 3) Eklentiyi Excel'e yükle  (yönetici GEREKMEZ)
npm run sideload
```

Sonra Excel'i tamamen kapatıp açın → Şeritte **Giriş (Home)** sekmesinde
**Hermes** grubu görünür → **Hermes'i Aç**. (Görünmezse: **Ekle → Eklentilerim →
Geliştirici Eklentileri**.)

Bir hücreye yazın:

```
=HERMES.SOR("İstanbul Kadıköy ada 123 parsel 4 için tapu bilgisi")
```

> `mock` modda gerçek cevap gelmez; bağlantıyı doğrulamak içindir. Gerçek ajanı
> bağlamak için aşağıya bakın.

---

## Hücre fonksiyonları

| Fonksiyon | Açıklama | Örnek |
|---|---|---|
| `=HERMES.SOR(soru; [bağlam])` | Serbest soru; ajan uygun MCP'yi kendi seçer | `=HERMES.SOR("bu parselin imar durumu?")` |
| `=HERMES.EMSAL(konu)` | Emsal karar getirir | `=HERMES.EMSAL("kamulaştırmasız el atma")` |
| `=HERMES.TAPU(sorgu)` | Tapu/parsel bilgisi | `=HERMES.TAPU("Ankara Çankaya ada 5 parsel 12")` |
| `=HERMES.PARSEL(il;ilçe;ada;parsel)` | Konum, yüzölçüm, nitelik | `=HERMES.PARSEL("İzmir";"Konak";"100";"7")` |
| `=HERMES.HESAP(ifade)` | Hesap motoru | `=HERMES.HESAP("850 m2 x 12000 TL/m2")` |
| `=HERMES.RAPOR(bölüm)` | Rapor bölümü üretir | `=HERMES.RAPOR("değerlendirme özeti")` |
| `=HERMES.HAFIZA(anahtar)` | Genel hafızadan hatırla | `=HERMES.HAFIZA("son dosya no")` |
| `=HERMES.MCP(sunucu;araç;argümanJSON)` | Genel MCP geçişi | `=HERMES.MCP("parsel-konum";"parsel_sorgu";"{""ada"":5}")` |
| `=HERMES.SINIFLA(değer;ölçüt)` | Metni sınıflandır (etiket) | `=HERMES.SINIFLA(A2;"risk: yüksek/orta/düşük")` |
| `=HERMES.CIKAR(değer;ne)` | Metinden alan çıkar | `=HERMES.CIKAR(A2;"parsel no")` |
| `=HERMES.OZETLE(aralık)` | Aralığı tek cümlede özetle | `=HERMES.OZETLE(A1:D20)` |
| `=HERMES.FORMUL(amaç)` | Açıklamadan Excel formülü | `=HERMES.FORMUL("yıllık büyüme %")` |

- Her fonksiyon önce hücrede **`⏳ Hermes…`** gösterir, cevap gelince değişir.
- Sonuçlar bellek içi **TTL'li önbelleğe** alınır (Excel'in tekrar-hesaplamaları
  ajanı yormaz). TTL: `.env` → `HERMES_FN_CACHE_TTL_MS`.
- Hata durumunda hücre `#HERMES! ...` döndürür (uydurma değer yazılmaz).

> Not: Argüman ayıracı yerel ayarınıza göre `;` veya `,` olabilir.

---

## Sohbet paneli (Hermes → Excel)

Şeritten **Hermes'i Aç** deyin. Panelden:

- Serbestçe konuşun; "Sayfa bağlamını gönder" işaretliyse Hermes aktif sayfayı,
  kullanılan aralıkları ve seçimi görür.
- Hermes çalışma kitabını değiştirmek istediğinde yanıtına **action**'lar ekler;
  panel bunları uygular: hücre yaz, sayfa oluştur, biçimlendir, sayı formatı,
  aralık temizle.
- Ayrıca Hermes **sohbet dışında da** çalışma kitabını sürebilir: bkz.
  [docs/HERMES_MCP_TOOL.md](docs/HERMES_MCP_TOOL.md) — Hermes'e küçük bir MCP
  aracı ekleyin, `POST /api/push` ile action gönderdiğinde panel anında uygular.

---

## Gerçek Hermes ajanını bağlama

`.env` içinde `HERMES_PROVIDER` seçin:

### `gateway` (önerilen) — Hermes API Server (tam ajan: araçlar + MCP + hafıza)

Hermes zaten **tam ajanı** OpenAI-uyumlu bir HTTP ucu olarak sunar (API Server —
`hermes proxy` değil; o yalnız model). Hermes tarafında `~/.hermes/.env`:

```env
API_SERVER_ENABLED=true
API_SERVER_KEY=<uzun rastgele gizli anahtar>
API_SERVER_CORS_ORIGINS=https://localhost:8787
```

`hermes gateway` ile başlatın; doğrulama: `curl http://localhost:8642/v1/health`.
Sonra bu projenin `.env`'i:

```env
HERMES_PROVIDER=gateway
HERMES_GATEWAY_URL=http://127.0.0.1:8642/v1/chat/completions
HERMES_API_KEY=...            # = API_SERVER_KEY
HERMES_MODEL=hermes-agent
```

Ajan, MCP araçlarını kendi tarafında çağırır; köprü sohbeti iletir, TLS'i
sonlandırır ve anahtarı **sunucu tarafında** ekler (anahtar hiçbir add-in
kodunda durmaz — ayrı bir Caddy proxy'ye gerek yoktur). `config/mcp-map.json`
her `HERMES.*` fonksiyonu için ajana verilecek sistem promptunu içerir — gerçek
MCP/araç adlarınızı oraya yazın.

> ⚠️ API Server ajanın **tüm araç setini (terminal komutları dahil)** açar.
> `API_SERVER_KEY`'i parola gibi saklayın, her şeyi `localhost`'a bağlayın,
> `API_SERVER_CORS_ORIGINS`'i dar tutun. Gerekmeyen toolset'leri Hermes tarafında
> kapatmayı düşünün.

### `anthropic` — doğrudan Claude

```env
HERMES_PROVIDER=anthropic
HERMES_ANTHROPIC_TOKEN=...    # hermes-auth-monitor tarafından yönetilebilir
HERMES_ANTHROPIC_MODEL=claude-opus-4-8
```

### `mock` — çevrimdışı test (varsayılan)

Ajan olmadan tüm eklenti uçtan uca çalışır.

---

## Yapılandırma (.env)

| Değişken | Varsayılan | Açıklama |
|---|---|---|
| `PORT` | `8787` | Köprü portu (yalnız 127.0.0.1) |
| `HERMES_BRIDGE_TOKEN` | boş | Ayarlıysa her `/api/*` çağrısı `X-Hermes-Token` ister |
| `HERMES_PROVIDER` | `mock` | `gateway` / `anthropic` / `mock` |
| `HERMES_FN_CACHE_TTL_MS` | `300000` | Fonksiyon önbellek süresi |
| `HERMES_FN_TIMEOUT_MS` | `90000` | Tek çağrı zaman aşımı |

> `HERMES_BRIDGE_TOKEN` ayarlarsanız aynı değeri `addin/functions.js` ve
> `addin/taskpane.js` içindeki `TOKEN` sabitlerine de yazın (SSE hariç).

---

## API (köprü)

| Uç nokta | İş |
|---|---|
| `GET /api/health` | Köprü + ajan durumu |
| `POST /api/fn` | Hücre fonksiyonu: `{fn,args}` → `{value}` |
| `POST /api/chat` | Sohbet: `{prompt,history,workbook,selection}` → `{message,actions}` |
| `POST /api/push` | Ajan → Excel: `{message,actions}` (SSE ile panele) |
| `GET /api/commands` | Panel için SSE komut akışı |

---

## Geliştirme

```powershell
npm run check     # tüm JS söz dizimi kontrolü
npm test          # formül rebase + action birim testleri
npm start         # köprüyü başlat
```

Dosya haritası için [ARCHITECTURE.md](ARCHITECTURE.md).

---

## Sorun giderme

- **Fonksiyon `#HERMES! köprü kapalı` diyor** → `npm start` çalışıyor mu?
  `https://localhost:8787/api/health` açılıyor mu?
- **Panel/fonksiyon yüklenmiyor** → Sertifika güvenilir mi? `npm run cert`
  tekrar çalıştırın, Excel'i yeniden başlatın.
- **`=HERMES.SOR` tanınmıyor** → Eklenti yüklü mü (`npm run sideload`)? Excel'i
  kapatıp açın; fonksiyonların kaydı birkaç saniye sürebilir.
- **Sideload görünmüyor** → Ekle → Eklentilerim → **Geliştirici Eklentileri**.
  Excel'i tam kapatıp açın. Kaldırıp tekrar: `npm run unsideload` sonra
  `npm run sideload`.
- **Ajan cevabı gelmiyor** ama sağlık "köprü açık · ajan ✗" → `.env` provider ve
  gateway URL/anahtarını kontrol edin; `mock` ile test edin.

## Güvenlik notları

- Köprü yalnız `127.0.0.1`'e bağlanır.
- Ajan araçları çalışma kitabına doğrudan yazmaz; yalnız yapılandırılmış
  `actions[]` panel tarafından uygulanır (savunma derinliği).
- Önbellek bellek içindir; kişisel veri diske yazılmaz.
- Üretimde `HERMES_BRIDGE_TOKEN` verin.

## Lisans

MIT
