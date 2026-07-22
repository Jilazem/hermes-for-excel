/* HermesExcel — çok dilli arayüz (i18n).
 * Yeni dil eklemek: aşağıdaki "tr" bloğunu Hermes'e çevirtip (ör: "bu JSON'u
 * Almancaya çevir") yeni bir kod (de/fr/es…) altında ekleyin. Sonra i18n.js'i
 * kaydedip paneli yenileyin; dil menüsünde otomatik çıkar (LANG_NAMES'e ekleyin).
 */
/* eslint-disable */
const LANG_NAMES = { tr: "Türkçe", en: "English", zh: "中文" };

const I18N = {
  tr: {
    ready_title: "Hermes hazır.",
    ready_body: 'Sohbetten iste, tabloyu senin için doldurayım; ya da hücreye =HERMES.SOR("…") yaz.',
    ready_hint: 'Örn: =HERMES.TAPU("İstanbul Kadıköy ada 123 parsel 4")',
    status_connecting: "bağlanıyor…",
    status_bridge_off: "köprü kapalı",
    status_agent_off: "köprü açık · ajan ✗",
    pending: "Hermes düşünüyor…",
    watch: "⏺ İzle",
    stop: "⏹ Durdur",
    rec_head: "İzleniyor —",
    rec_ops: "işlem",
    rec_started: "⏺ İzleme başladı. Yaptığın her değişiklik aşağıda anlık görünecek. Bitince '⏹ Durdur'a bas — otomasyon önereyim.",
    rec_none: "İzleme durdu — hiç değişiklik algılanmadı.",
    rec_automate: "izlenen işlemi otomatikleştir.",
    sel_none: "seçim yok",
    sel_cells: "hücre",
    sel_use: "↳ kullan",
    settings_title: "Bağlantı Ayarları",
    lang_label: "Dil / Language",
    provider: "Sağlayıcı",
    gateway_url: "Gateway / API URL",
    api_key: "API Anahtarı",
    model: "Model",
    anth_token: "Anthropic Token",
    save: "Kaydet",
    close: "Kapat",
    saving: "Kaydediliyor…",
    saved: "✓ Kaydedildi.",
    input_ph: "Hermes'e yaz… (Enter gönder, Shift+Enter satır)",
    ctx_send: "Sayfa bağlamını gönder",
    send: "Gönder",
    unreachable: "⚠️ Köprüye ulaşılamadı: ",
    suggest_fail: "⚠️ Öneri alınamadı: ",
    empty: "(boş)",
    agent_updated: "⇢ Hermes çalışma kitabını güncelledi",
  },
  en: {
    ready_title: "Hermes is ready.",
    ready_body: 'Ask in chat and I\'ll fill the table for you, or type =HERMES.SOR("…") in a cell.',
    ready_hint: 'e.g. =HERMES.TAPU("Istanbul Kadikoy block 123 parcel 4")',
    status_connecting: "connecting…",
    status_bridge_off: "bridge offline",
    status_agent_off: "bridge up · agent ✗",
    pending: "Hermes is thinking…",
    watch: "⏺ Watch",
    stop: "⏹ Stop",
    rec_head: "Recording —",
    rec_ops: "ops",
    rec_started: "⏺ Recording started. Each change appears live below. Press '⏹ Stop' when done — I'll suggest an automation.",
    rec_none: "Recording stopped — no changes detected.",
    rec_automate: "automate the recorded steps.",
    sel_none: "no selection",
    sel_cells: "cells",
    sel_use: "↳ use",
    settings_title: "Connection Settings",
    lang_label: "Language / Dil",
    provider: "Provider",
    gateway_url: "Gateway / API URL",
    api_key: "API Key",
    model: "Model",
    anth_token: "Anthropic Token",
    save: "Save",
    close: "Close",
    saving: "Saving…",
    saved: "✓ Saved.",
    input_ph: "Message Hermes… (Enter to send, Shift+Enter for newline)",
    ctx_send: "Send sheet context",
    send: "Send",
    unreachable: "⚠️ Bridge unreachable: ",
    suggest_fail: "⚠️ Couldn't get suggestion: ",
    empty: "(empty)",
    agent_updated: "⇢ Hermes updated the workbook",
  },
  zh: {
    ready_title: "Hermes 已就绪。",
    ready_body: '在聊天中提问，我来帮你填表；或在单元格输入 =HERMES.SOR("…")。',
    ready_hint: '例如：=HERMES.TAPU("伊斯坦布尔 卡德柯伊 123街区 4号地块")',
    status_connecting: "连接中…",
    status_bridge_off: "桥接离线",
    status_agent_off: "桥接正常 · 代理 ✗",
    pending: "Hermes 思考中…",
    watch: "⏺ 记录",
    stop: "⏹ 停止",
    rec_head: "记录中 —",
    rec_ops: "步",
    rec_started: "⏺ 开始记录。每次更改会实时显示在下方。完成后点击“⏹ 停止”，我会给出自动化建议。",
    rec_none: "记录已停止 — 未检测到更改。",
    rec_automate: "将记录的步骤自动化。",
    sel_none: "未选择",
    sel_cells: "单元格",
    sel_use: "↳ 使用",
    settings_title: "连接设置",
    lang_label: "语言 / Language",
    provider: "提供商",
    gateway_url: "网关 / API URL",
    api_key: "API 密钥",
    model: "模型",
    anth_token: "Anthropic 令牌",
    save: "保存",
    close: "关闭",
    saving: "保存中…",
    saved: "✓ 已保存。",
    input_ph: "给 Hermes 发消息…（Enter 发送，Shift+Enter 换行）",
    ctx_send: "发送表格上下文",
    send: "发送",
    unreachable: "⚠️ 无法连接桥接：",
    suggest_fail: "⚠️ 无法获取建议：",
    empty: "(空)",
    agent_updated: "⇢ Hermes 已更新工作簿",
  },
};

let LANG = (typeof localStorage !== "undefined" && localStorage.getItem("hermesLang")) || "tr";
function t(k) {
  return (I18N[LANG] && I18N[LANG][k]) || I18N.tr[k] || k;
}
function setLang(l) {
  LANG = I18N[l] ? l : "tr";
  try { localStorage.setItem("hermesLang", LANG); } catch {}
  applyI18n();
}
function applyI18n() {
  // GÜVENLİK: yalnız yaprak öğelerin metnini değiştir; çocuk (input/select) içeren
  // öğeyi ezme — aksi halde form elemanları silinir.
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    if (el.children.length === 0) el.textContent = t(el.getAttribute("data-i18n"));
  });
  document.querySelectorAll("[data-i18n-ph]").forEach((el) => (el.placeholder = t(el.getAttribute("data-i18n-ph"))));
  document.querySelectorAll("[data-i18n-title]").forEach((el) => (el.title = t(el.getAttribute("data-i18n-title"))));
  // dil menüsünü doldur
  const sel = document.getElementById("setLang");
  if (sel && !sel.dataset.filled) {
    for (const code in LANG_NAMES) {
      const o = document.createElement("option");
      o.value = code; o.textContent = LANG_NAMES[code];
      sel.appendChild(o);
    }
    sel.dataset.filled = "1";
  }
  if (sel) sel.value = LANG;
  document.documentElement.lang = LANG;
}
