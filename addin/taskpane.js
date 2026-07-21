/* global Office, Excel */
// Hermes sohbet paneli: /api/chat ile konuşur, dönen action'ları uygular ve
// /api/commands (SSE) üzerinden ajanın gönderdiği komutları çalışma kitabına yansıtır.

// Köprü adresini, panelin yüklendiği origin'den türet (GX10 gibi uzak sunucuda otomatik).
const BRIDGE =
  window.location && /^https:/.test(window.location.origin || "")
    ? window.location.origin
    : "https://localhost:8787";
const TOKEN = ""; // HERMES_BRIDGE_TOKEN kullanıyorsanız aynı değeri yazın (sırrı koda GÖMMEYİN — bu dosya paylaşılabilir)
const history = [];

const $ = (id) => document.getElementById(id);

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  // Çok dilli arayüz
  applyI18n();
  updateRecBtn();
  $("setLang").addEventListener("change", (e) => setLang(e.target.value));

  $("send").addEventListener("click", send);
  $("input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      send();
    }
  });
  // Ayarlar
  $("settingsBtn").addEventListener("click", toggleSettings);
  $("settingsCancel").addEventListener("click", () => $("settings").classList.add("hidden"));
  $("settingsSave").addEventListener("click", saveSettings);
  $("setProvider").addEventListener("change", updateProviderFields);

  pollHealth();
  setInterval(pollHealth, 15000);
  connectCommandStream();

  // Seçili hücreleri canlı algıla (Claude add-in gibi)
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, updateSelectionBar);
  } catch (e) {
    /* olay desteklenmiyorsa yoksay */
  }
  updateSelectionBar();
  $("selUse").addEventListener("click", useSelection);

  // İşlem izleyici (makro-benzeri): hücre değişimlerini kaydet, otomasyon öner
  $("recBtn").addEventListener("click", toggleRecord);
  try {
    Excel.run(async (ctx) => {
      ctx.workbook.worksheets.onChanged.add(onWorkbookChanged);
      await ctx.sync();
    });
  } catch (e) {
    /* onChanged desteklenmiyorsa yoksay */
  }
});

// ── İşlem izleyici (otomasyon önerici) ───────────────────────────────────────
let recording = false;
let recordLog = [];

function toggleRecord() {
  recording = !recording;
  if (recording) {
    recordLog = [];
    $("recLines").innerHTML = "";
    $("recLog").classList.remove("hidden");
    updateRecBtn();
    addMsg("sys", t("rec_started"));
  } else {
    $("recLog").classList.add("hidden");
    updateRecBtn();
    suggestAutomation();
  }
}

function updateRecBtn() {
  const b = $("recBtn");
  b.textContent = recording ? `${t("stop")} (${recordLog.length})` : t("watch");
  b.style.color = recording ? "var(--off)" : "";
  const c = document.getElementById("recCount");
  if (c) c.textContent = recordLog.length;
}

// canlı log satırı ekle
function addRecLine(addr, formula, value) {
  const shown = typeof formula === "string" && formula.startsWith("=") ? formula : String(value ?? "");
  const div = document.createElement("div");
  div.className = "rec-line";
  div.textContent = `${addr} = ${shown}`.slice(0, 90);
  const box = $("recLines");
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
}

async function onWorkbookChanged(event) {
  if (!recording) return;
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(event.worksheetId);
      const rng = sheet.getRange(event.address);
      rng.load(["address", "formulas", "values"]);
      sheet.load("name");
      await ctx.sync();
      const addr = rng.address.replace(/^.*!/, "");
      const formula = rng.formulas?.[0]?.[0];
      const value = rng.values?.[0]?.[0];
      recordLog.push({ sheet: sheet.name, address: addr, formula, value, type: event.changeType });
      addRecLine(addr, formula, value);
    });
  } catch (e) {
    recordLog.push({ address: event.address, type: event.changeType });
    addRecLine(event.address, null, "(değişti)");
  }
  updateRecBtn();
}

async function suggestAutomation() {
  if (!recordLog.length) {
    addMsg("sys", t("rec_none"));
    return;
  }
  const steps = recordLog
    .map((r, i) => `${i + 1}. ${r.address}${r.formula && String(r.formula).startsWith("=") ? " = " + r.formula : ": " + String(r.value ?? "")}`)
    .join("\n");
  addMsg("user", `⏺ ${recordLog.length} · ${t("rec_automate")}`);
  setPending(true);
  try {
    const res = await fetch(BRIDGE + "/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json", ...(TOKEN ? { "x-hermes-token": TOKEN } : {}) },
      body: JSON.stringify({
        prompt:
          "Excel'de az önce şu adımları yaptım:\n" + steps +
          "\n\nBu tekrarlı işi otomatikleştiren EN İYİ çözümü öner: (a) tek bir Excel formülü, (b) VBA makrosu, veya (c) panelden uygulanabilir adımlar. " +
          "Formülü/makroyu kod olarak ver, kısa açıkla. Uygulanabilir bir tablo düzeni varsa 'actions' ile de önerebilirsin.",
        history: history.slice(-6),
      }),
    });
    const data = await res.json();
    history.push({ role: "assistant", content: data.message || "" });
    let note = "";
    if (Array.isArray(data.actions) && data.actions.length) note = await applyActions(data.actions);
    addMsg("bot", data.message || t("empty"), note || undefined);
  } catch (e) {
    addMsg("bot", t("suggest_fail") + (e.message || e));
  } finally {
    setPending(false);
  }
}

// ── Seçim takibi ─────────────────────────────────────────────────────────────
let lastSelection = null;
async function updateSelectionBar() {
  try {
    const info = await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load(["address", "rowCount", "columnCount", "values"]);
      await ctx.sync();
      return { address: sel.address, rows: sel.rowCount, cols: sel.columnCount, values: sel.values };
    });
    lastSelection = info;
    const n = info.rows * info.cols;
    const addr = (info.address || "").replace(/^.*!/, "");
    let preview = "";
    if (n <= 6) {
      const flat = info.values.flat().map((v) => String(v ?? "").slice(0, 14)).filter(Boolean);
      if (flat.length) preview = " — " + flat.join(", ").slice(0, 60);
    }
    $("selText").textContent = `${addr} · ${n} ${t("sel_cells")}${preview}`;
    $("selUse").classList.remove("hidden");
  } catch (e) {
    lastSelection = null;
    $("selText").textContent = t("sel_none");
    $("selUse").classList.add("hidden");
  }
}
function useSelection() {
  const addr = (lastSelection?.address || "").replace(/^.*!/, "") || "seçili aralık";
  const input = $("input");
  input.value = (input.value ? input.value + " " : "") + `(seçili aralık ${addr})`;
  input.focus();
}

// ── Ayarlar paneli ───────────────────────────────────────────────────────────
function updateProviderFields() {
  const p = $("setProvider").value;
  $("gatewayFields").classList.toggle("hidden", !(p === "gateway" || p === "openai"));
  $("anthropicFields").classList.toggle("hidden", p !== "anthropic");
}

async function toggleSettings() {
  const box = $("settings");
  if (!box.classList.contains("hidden")) {
    box.classList.add("hidden");
    return;
  }
  // aç + mevcut ayarları yükle
  try {
    const r = await fetch(BRIDGE + "/api/settings", { headers: TOKEN ? { "x-hermes-token": TOKEN } : {} });
    const s = await r.json();
    $("setProvider").value = s.provider || "gateway";
    $("setGatewayUrl").value = s.gatewayUrl || "";
    $("setModel").value = s.model || "";
    $("setApiKey").value = "";
    $("setApiKey").placeholder = s.apiKeySet ? "kayıtlı (" + (s.apiKeyMasked || "••••") + ") — değiştirmek için yaz" : "sk-…";
    $("setAnthModel").value = s.anthropicModel || "";
    $("setAnthToken").value = "";
    $("setAnthToken").placeholder = s.anthropicTokenSet ? "kayıtlı — değiştirmek için yaz" : "sk-ant-…";
    $("settingsMsg").textContent = "";
    updateProviderFields();
  } catch (e) {
    $("settingsMsg").textContent = "Ayarlar okunamadı: " + (e.message || e);
  }
  box.classList.remove("hidden");
}

async function saveSettings() {
  const patch = {
    provider: $("setProvider").value,
    gatewayUrl: $("setGatewayUrl").value.trim(),
    model: $("setModel").value.trim(),
    anthropicModel: $("setAnthModel").value.trim(),
  };
  // boş bırakılan sırlar değişmez; yeni değer yazıldıysa gönder
  const key = $("setApiKey").value.trim();
  if (key) patch.apiKey = key;
  const tok = $("setAnthToken").value.trim();
  if (tok) patch.anthropicToken = tok;

  $("settingsMsg").textContent = t("saving");
  try {
    const r = await fetch(BRIDGE + "/api/settings", {
      method: "POST",
      headers: { "content-type": "application/json", ...(TOKEN ? { "x-hermes-token": TOKEN } : {}) },
      body: JSON.stringify(patch),
    });
    if (!r.ok) throw new Error("HTTP " + r.status);
    $("settingsMsg").textContent = t("saved");
    pollHealth();
    setTimeout(() => $("settings").classList.add("hidden"), 800);
  } catch (e) {
    $("settingsMsg").textContent = "Kaydedilemedi: " + (e.message || e);
  }
}

// ── durum ──────────────────────────────────────────────────────────────────
async function pollHealth() {
  const el = $("status");
  try {
    const r = await fetch(BRIDGE + "/api/health");
    const d = await r.json();
    const ok = d.hermes?.ok;
    el.textContent = ok ? `● ${d.provider}` : t("status_agent_off");
    el.className = "status " + (ok ? "status--ok" : "status--off");
    el.title = d.hermes?.status || "";
  } catch {
    el.textContent = t("status_bridge_off");
    el.className = "status status--off";
  }
}

// ── mesaj arayüzü ────────────────────────────────────────────────────────────
function addMsg(role, text, note) {
  const div = document.createElement("div");
  div.className = "msg msg--" + role;
  div.textContent = text;
  if (note) {
    const n = document.createElement("div");
    n.className = "actions-note";
    n.textContent = note;
    div.appendChild(n);
  }
  $("log").appendChild(div);
  $("log").scrollTop = $("log").scrollHeight;
  return div;
}
function setPending(on) {
  $("pending").classList.toggle("hidden", !on);
  $("send").disabled = on;
}

// ── gönder ───────────────────────────────────────────────────────────────────
async function send() {
  const input = $("input");
  const prompt = input.value.trim();
  if (!prompt) return;
  input.value = "";
  addMsg("user", prompt);
  history.push({ role: "user", content: prompt });
  setPending(true);

  let workbook = null;
  let selection = null;
  if ($("ctxChk").checked) {
    try {
      ({ workbook, selection } = await gatherContext());
    } catch (e) {
      /* bağlam alınamazsa devam et */
    }
  }

  try {
    const res = await fetch(BRIDGE + "/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json", ...(TOKEN ? { "x-hermes-token": TOKEN } : {}) },
      body: JSON.stringify({ prompt, history: history.slice(-12), workbook, selection }),
    });
    const data = await res.json();
    history.push({ role: "assistant", content: data.message || "" });
    let note = "";
    if (Array.isArray(data.actions) && data.actions.length) {
      const summary = await applyActions(data.actions);
      note = summary;
    }
    addMsg("bot", data.message || t("empty"), note || undefined);
  } catch (e) {
    addMsg("bot", t("unreachable") + (e.message || e));
  } finally {
    setPending(false);
  }
}

// ── çalışma kitabı bağlamı ───────────────────────────────────────────────────
async function gatherContext() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    const active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    const sel = ctx.workbook.getSelectedRange();
    sel.load(["address", "values", "rowCount", "columnCount"]);
    await ctx.sync();

    const sheetInfos = [];
    for (const s of sheets.items) {
      const used = s.getUsedRangeOrNullObject(true);
      used.load(["address", "rowCount", "columnCount"]);
      sheetInfos.push({ sheet: s, used });
    }
    await ctx.sync();

    const wb = {
      activeSheet: active.name,
      sheets: sheetInfos.map(({ sheet, used }) => ({
        name: sheet.name,
        usedRange: used.isNullObject ? null : used.address,
        rowCount: used.isNullObject ? 0 : used.rowCount,
        columnCount: used.isNullObject ? 0 : used.columnCount,
      })),
    };
    const selection = {
      address: sel.address,
      values: sel.rowCount * sel.columnCount <= 200 ? sel.values : null,
      rowCount: sel.rowCount,
      columnCount: sel.columnCount,
    };
    return { workbook: wb, selection };
  });
}

// ── action uygulayıcı ────────────────────────────────────────────────────────
async function applyActions(actions) {
  const done = [];
  await Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    for (const a of actions) {
      try {
        if (a.type === "write_cells") {
          const sheet = a.sheet ? wb.worksheets.getItem(a.sheet) : wb.worksheets.getActiveWorksheet();
          const values = a.values || [];
          if (!values.length) continue;
          const rows = values.length, cols = Math.max(...values.map((r) => (r ? r.length : 0)));
          const start = sheet.getRange(a.start_cell || "A1");
          const range = start.getResizedRange(rows - 1, cols - 1);
          // dikdörtgen yap
          const rect = values.map((r) => {
            const rr = (r || []).slice(0, cols);
            while (rr.length < cols) rr.push("");
            return rr;
          });
          range.formulas = rect; // formül veya düz değer
          done.push(`✎ ${a.start_cell}: ${rows}×${cols} yazıldı`);
        } else if (a.type === "create_sheet") {
          let sheet;
          try { sheet = wb.worksheets.getItem(a.name); }
          catch { sheet = wb.worksheets.add(a.name); }
          if (Array.isArray(a.values) && a.values.length) {
            const rows = a.values.length, cols = Math.max(...a.values.map((r) => (r ? r.length : 0)));
            const rect = a.values.map((r) => { const rr = (r||[]).slice(0,cols); while (rr.length<cols) rr.push(""); return rr; });
            sheet.getRange("A1").getResizedRange(rows - 1, cols - 1).formulas = rect;
          }
          sheet.activate();
          done.push(`＋ sayfa oluşturuldu: ${a.name}`);
        } else if (a.type === "format_cells") {
          const sheet = a.sheet ? wb.worksheets.getItem(a.sheet) : wb.worksheets.getActiveWorksheet();
          const r = sheet.getRange(a.range);
          if (a.bold != null) r.format.font.bold = !!a.bold;
          if (a.italic != null) r.format.font.italic = !!a.italic;
          if (a.font_color) r.format.font.color = a.font_color;
          if (a.fill) r.format.fill.color = a.fill;
          if (a.align) r.format.horizontalAlignment = a.align;
          if (a.number_format) r.numberFormat = [[a.number_format]];
          done.push(`🎨 biçim: ${a.range}`);
        } else if (a.type === "set_number_format") {
          const sheet = wb.worksheets.getActiveWorksheet();
          sheet.getRange(a.range).numberFormat = [[a.format]];
          done.push(`# format: ${a.range}`);
        } else if (a.type === "clear_range") {
          const sheet = wb.worksheets.getActiveWorksheet();
          sheet.getRange(a.range).clear();
          done.push(`⌫ temizlendi: ${a.range}`);
        } else if (a.type === "read_range") {
          // okuma döngüsü v1'de panelde okunur ama tur atmadan geçilir
          done.push(`👁 okuma istendi: ${a.range}`);
        }
      } catch (e) {
        done.push(`✗ ${a.type}: ${e.message || e}`);
      }
    }
    await ctx.sync();
  }).catch((e) => done.push("✗ uygulama hatası: " + (e.message || e)));
  return done.join("  ·  ");
}

// ── ajan → Excel komut akışı (SSE) ───────────────────────────────────────────
function connectCommandStream() {
  try {
    // EventSource header gönderemez → token'ı sorgu parametresiyle ilet
    const url = BRIDGE + "/api/commands" + (TOKEN ? "?token=" + encodeURIComponent(TOKEN) : "");
    const es = new EventSource(url);
    es.addEventListener("commands", async (ev) => {
      try {
        const payload = JSON.parse(ev.data);
        const note = payload.actions?.length ? await applyActions(payload.actions) : "";
        addMsg("bot", t("agent_updated") + (payload.message ? ": " + payload.message : ""), note || undefined);
      } catch (e) {
        addMsg("sys", "SSE komutu işlenemedi: " + (e.message || e));
      }
    });
    es.onerror = () => { /* köprü kapanınca EventSource kendi yeniden dener */ };
  } catch (e) {
    /* EventSource yoksa yoksay */
  }
}
