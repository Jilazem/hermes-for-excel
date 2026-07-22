/* global CustomFunctions */
// =HERMES.* özel fonksiyonları. Her biri köprüye (/api/fn) çağrı yapar ve
// Hermes ajanının cevabını hücreye yazar. Streaming: önce "⏳" gösterir,
// cevap gelince değiştirir.

// Köprü adresi. Custom-functions runtime'ında self.location.origin güvenilir
// çözülmediği için (task pane'den farklı) SABİT tutuyoruz. Bu kurulumda add-in
// localhost:8799 proxy'sinden yükleniyor; proxy istekleri GX10 köprüsüne iletir.
const BRIDGE = "https://localhost:8799";
const TOKEN = ""; // HERMES_BRIDGE_TOKEN kullanıyorsanız buraya aynı değeri yazın (sırrı koda GÖMMEYİN — bu dosya paylaşılabilir)

async function callBridge(fn, args, invocation) {
  const streaming = invocation && typeof invocation.setResult === "function";
  if (streaming) invocation.setResult("⏳ Hermes…"); // anında geri bildirim
  let result;
  try {
    const res = await fetch(BRIDGE + "/api/fn", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        ...(TOKEN ? { "x-hermes-token": TOKEN } : {}),
      },
      body: JSON.stringify({ fn, args }),
    });
    if (!res.ok) result = `#HERMES! HTTP ${res.status}`;
    else {
      const data = await res.json();
      result = data.value != null ? String(data.value) : "#HERMES! boş";
    }
  } catch (e) {
    result = "#HERMES! köprü kapalı — bridge çalışıyor mu? (npm start)";
  }
  // Streaming fonksiyonlarda dönüş değeri yok sayılır; setResult zorunlu.
  if (streaming) {
    invocation.setResult(result);
    return;
  }
  return result;
}

// Matris (taşma dizisi) döndüren fonksiyonlar için köprü çağrısı.
async function callBridgeMatrix(fn, args, invocation) {
  const streaming = invocation && typeof invocation.setResult === "function";
  if (streaming) invocation.setResult([["⏳ Hermes…"]]);
  let result;
  try {
    const res = await fetch(BRIDGE + "/api/fn", {
      method: "POST",
      headers: { "content-type": "application/json", ...(TOKEN ? { "x-hermes-token": TOKEN } : {}) },
      body: JSON.stringify({ fn, args }),
    });
    if (!res.ok) result = [[`#HERMES! HTTP ${res.status}`]];
    else {
      const data = await res.json();
      result = Array.isArray(data.value) ? data.value : [[String(data.value == null ? "#HERMES! boş" : data.value)]];
    }
  } catch (e) {
    result = [["#HERMES! köprü kapalı — bridge çalışıyor mu?"]];
  }
  if (streaming) {
    invocation.setResult(result);
    return;
  }
  return result;
}

// ── HERMES.DOSYA(etiketler, klasor) → alan tablosu (kolon taşma dizisi) ──────
async function DOSYA(etiketler, klasor, invocation) {
  return callBridgeMatrix("DOSYA", [etiketler, klasor], invocation);
}
// ── HERMES.HISSEDAR(klasor) → hissedar tablosu (2B taşma dizisi) ─────────────
async function HISSEDAR(klasor, invocation) {
  return callBridgeMatrix("HISSEDAR", [klasor], invocation);
}

// ── HERMES.SOR(soru, [bağlam]) ────────────────────────────────────────────
async function SOR(soru, baglam, invocation) {
  return callBridge("SOR", [soru, baglam], invocation);
}
// ── HERMES.EMSAL(konu) ────────────────────────────────────────────────────
async function EMSAL(konu, invocation) {
  return callBridge("EMSAL", [konu], invocation);
}
// ── HERMES.TAPU(sorgu) ────────────────────────────────────────────────────
async function TAPU(sorgu, invocation) {
  return callBridge("TAPU", [sorgu], invocation);
}
// ── HERMES.PARSEL(il, ilce, ada, parsel) ──────────────────────────────────
async function PARSEL(il, ilce, ada, parsel, invocation) {
  return callBridge("PARSEL", [il, ilce, ada, parsel], invocation);
}
// ── HERMES.HESAP(ifade) ───────────────────────────────────────────────────
async function HESAP(ifade, invocation) {
  return callBridge("HESAP", [ifade], invocation);
}
// ── HERMES.RAPOR(bolum) ───────────────────────────────────────────────────
async function RAPOR(bolum, invocation) {
  return callBridge("RAPOR", [bolum], invocation);
}
// ── HERMES.HAFIZA(anahtar) ────────────────────────────────────────────────
async function HAFIZA(anahtar, invocation) {
  return callBridge("HAFIZA", [anahtar], invocation);
}
// ── HERMES.MCP(sunucu, arac, argumanlarJSON) — genel MCP geçişi ────────────
async function MCP(sunucu, arac, argumanlarJSON, invocation) {
  return callBridge("MCP", [sunucu, arac, argumanlarJSON], invocation);
}

// ── Genel amaçlı fonksiyonlar (hücre verisi üzerinde) ──────────────────────
// ── HERMES.SINIFLA(deger, olcut) ──────────────────────────────────────────
async function SINIFLA(deger, olcut, invocation) {
  return callBridge("SINIFLA", [deger, olcut], invocation);
}
// ── HERMES.CIKAR(deger, ne) ───────────────────────────────────────────────
async function CIKAR(deger, ne, invocation) {
  return callBridge("CIKAR", [deger, ne], invocation);
}
// ── HERMES.OZETLE(aralik) ─────────────────────────────────────────────────
async function OZETLE(aralik, invocation) {
  return callBridge("OZETLE", [aralik], invocation);
}
// ── HERMES.FORMUL(amac) ───────────────────────────────────────────────────
async function FORMUL(amac, invocation) {
  return callBridge("FORMUL", [amac], invocation);
}

// ── İngilizce takma adlar (English aliases — Türkçe olanlarla birebir aynı) ──
async function ASK(q, ctx, i) { return callBridge("SOR", [q, ctx], i); }
async function PRECEDENT(topic, i) { return callBridge("EMSAL", [topic], i); }
async function DEED(query, i) { return callBridge("TAPU", [query], i); }
async function PARCEL_EN(prov, dist, block, parcel, i) { return callBridge("PARSEL", [prov, dist, block, parcel], i); }
async function CALC(expr, i) { return callBridge("HESAP", [expr], i); }
async function REPORT(section, i) { return callBridge("RAPOR", [section], i); }
async function RECALL(k, i) { return callBridge("HAFIZA", [k], i); }
async function CLASSIFY(v, rule, i) { return callBridge("SINIFLA", [v, rule], i); }
async function EXTRACT(v, what, i) { return callBridge("CIKAR", [v, what], i); }
async function SUMMARIZE(range, i) { return callBridge("OZETLE", [range], i); }
async function FORMULA(goal, i) { return callBridge("FORMUL", [goal], i); }
async function FILES(labels, folder, i) { return callBridgeMatrix("DOSYA", [labels, folder], i); }
async function SHAREHOLDERS(folder, i) { return callBridgeMatrix("HISSEDAR", [folder], i); }

CustomFunctions.associate("ASK", ASK);
CustomFunctions.associate("PRECEDENT", PRECEDENT);
CustomFunctions.associate("DEED", DEED);
CustomFunctions.associate("PARCEL", PARCEL_EN);
CustomFunctions.associate("CALC", CALC);
CustomFunctions.associate("REPORT", REPORT);
CustomFunctions.associate("RECALL", RECALL);
CustomFunctions.associate("CLASSIFY", CLASSIFY);
CustomFunctions.associate("EXTRACT", EXTRACT);
CustomFunctions.associate("SUMMARIZE", SUMMARIZE);
CustomFunctions.associate("FORMULA", FORMULA);
CustomFunctions.associate("FILES", FILES);
CustomFunctions.associate("SHAREHOLDERS", SHAREHOLDERS);

CustomFunctions.associate("SOR", SOR);
CustomFunctions.associate("EMSAL", EMSAL);
CustomFunctions.associate("TAPU", TAPU);
CustomFunctions.associate("PARSEL", PARSEL);
CustomFunctions.associate("HESAP", HESAP);
CustomFunctions.associate("RAPOR", RAPOR);
CustomFunctions.associate("HAFIZA", HAFIZA);
CustomFunctions.associate("MCP", MCP);
CustomFunctions.associate("SINIFLA", SINIFLA);
CustomFunctions.associate("CIKAR", CIKAR);
CustomFunctions.associate("OZETLE", OZETLE);
CustomFunctions.associate("FORMUL", FORMUL);
CustomFunctions.associate("DOSYA", DOSYA);
CustomFunctions.associate("HISSEDAR", HISSEDAR);
