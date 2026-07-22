// HermesExcel köprüsü: HTTPS statik sunucu + API.
// Yalnız 127.0.0.1'e bağlanır. Office eklentileri HTTPS zorunlu kılar; bu yüzden
// office-addin-dev-certs ile üretilen güvenilir yerel sertifika kullanılır.
import http from "node:http";
import https from "node:https";
import { readFileSync, existsSync } from "node:fs";
import { promises as fsp } from "node:fs";
import { join, extname, normalize, resolve } from "node:path";
import { homedir } from "node:os";
import { config, mcpMap } from "./config.mjs";
import { askHermes, healthProbe } from "./hermes-client.mjs";
import { getSettingsMasked, setSettings } from "./settings.mjs";
import { normalizeActions, extractJson, extractJsonArray } from "./actions.mjs";
import { addClient, pushCommands, clientCount } from "./command-bus.mjs";

const ADDIN_DIR = join(config.root, "addin");
const ASSETS_DIR = join(config.root, "assets");
const MAP = mcpMap();

// ── TLS sertifikalarını bul ───────────────────────────────────────────────
function findCerts() {
  if (config.tls.cert && config.tls.key && existsSync(config.tls.cert) && existsSync(config.tls.key)) {
    return { cert: readFileSync(config.tls.cert), key: readFileSync(config.tls.key) };
  }
  const dir = join(homedir(), ".office-addin-dev-certs");
  const crt = join(dir, "localhost.crt");
  const key = join(dir, "localhost.key");
  if (existsSync(crt) && existsSync(key)) {
    return { cert: readFileSync(crt), key: readFileSync(key) };
  }
  return null;
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
  ".xml": "text/xml; charset=utf-8",
};

// ── yardımcılar ───────────────────────────────────────────────────────────
function setCors(res) {
  res.setHeader("access-control-allow-origin", "*");
  res.setHeader("access-control-allow-headers", "content-type,authorization,x-hermes-token");
  res.setHeader("access-control-allow-methods", "GET,POST,OPTIONS");
}

function json(res, code, obj) {
  const body = JSON.stringify(obj);
  setCors(res);
  res.writeHead(code, { "content-type": "application/json; charset=utf-8" });
  res.end(body);
}

function readBody(req, limitBytes = 25 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (c) => {
      size += c.length;
      if (size > limitBytes) {
        reject(new Error("body too large"));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function authed(req) {
  if (!config.bridgeToken) return true;
  const h = req.headers["x-hermes-token"] || (req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  return h === config.bridgeToken;
}

// ── statik dosyalar ───────────────────────────────────────────────────────
async function serveStatic(req, res, pathname) {
  let rel = pathname === "/" ? "/taskpane.html" : pathname;
  // /assets/* -> assets klasörü, diğerleri -> addin klasörü
  let baseDir = ADDIN_DIR;
  if (rel.startsWith("/assets/")) {
    baseDir = ASSETS_DIR;
    rel = rel.slice("/assets".length);
  }
  const safe = normalize(rel).replace(/^(\.\.[/\\])+/, "");
  const file = resolve(baseDir, "." + safe);
  if (!file.startsWith(baseDir)) {
    res.writeHead(403);
    return res.end("forbidden");
  }
  try {
    const data = await fsp.readFile(file);
    setCors(res);
    res.writeHead(200, { "content-type": MIME[extname(file).toLowerCase()] || "application/octet-stream" });
    res.end(data);
  } catch {
    res.writeHead(404);
    res.end("not found");
  }
}

// ── /api/fn : hücre içi özel fonksiyon ────────────────────────────────────
const fnCache = new Map(); // key -> { value, at }

async function handleFn(req, res) {
  const body = JSON.parse((await readBody(req)) || "{}");
  const fn = String(body.fn || "SOR").toUpperCase();
  const args = Array.isArray(body.args) ? body.args : [];
  const query = args
    .map((a) => (a == null ? "" : typeof a === "object" ? JSON.stringify(a) : String(a)))
    .join(" | ")
    .trim();

  const key = `${fn}::${query}`;
  const cached = fnCache.get(key);
  if (cached && Date.now() - cached.at < config.fnCacheTtlMs) {
    return json(res, 200, { value: cached.value, cached: true });
  }

  const spec = MAP.functions?.[fn];

  // ── DOSYA(etiketler, klasör) → etiket sırasına göre değer KOLONU (taşma matrisi)
  if (fn === "DOSYA") {
    const labels = (Array.isArray(args[0]) ? args[0].flat() : [args[0]])
      .map((x) => String(x == null ? "" : x).trim())
      .filter(Boolean);
    const folder = translatePath(args[1]);
    if (!labels.length) return json(res, 200, { value: [["#HERMES! etiket aralığı boş"]] });
    const user =
      `Klasör: ${folder}\nÇıkarılacak alanlar (bu sırayla, ${labels.length} adet):\n` +
      labels.map((l, i) => `${i + 1}. ${l}`).join("\n") +
      `\n\nHer alan için değeri döndür.`;
    try {
      const raw = await askHermes(
        [{ role: "system", content: spec.system }, { role: "user", content: user }],
        { timeoutMs: 240000, maxTokens: 4096, idempotencyKey: key }
      );
      let arr = extractJsonArray(raw) || [];
      while (arr.length < labels.length) arr.push("Belirtilmemiş");
      arr = arr.slice(0, labels.length);
      const value = arr.map((v) => [v == null ? "" : String(v)]); // kolon matris
      fnCache.set(key, { value, at: Date.now() });
      return json(res, 200, { value });
    } catch (e) {
      return json(res, 200, { value: labels.map(() => [`#HERMES! ${errShort(e)}`]) });
    }
  }

  // ── HISSEDAR(klasör) → 5 sütunlu hissedar tablosu (taşma matrisi)
  if (fn === "HISSEDAR") {
    const folder = translatePath(args[0]);
    try {
      const raw = await askHermes(
        [{ role: "system", content: spec.system }, { role: "user", content: `Klasör: ${folder}` }],
        { timeoutMs: 240000, maxTokens: 4096, idempotencyKey: key }
      );
      const arr = extractJsonArray(raw) || [];
      const matrix = arr
        .filter(Array.isArray)
        .map((row) => {
          const r = row.slice(0, 5).map((v) => (v == null ? "" : String(v)));
          while (r.length < 5) r.push("");
          return r;
        });
      const value = matrix.length ? matrix : [["Belirtilmemiş", "", "", "", ""]];
      fnCache.set(key, { value, at: Date.now() });
      return json(res, 200, { value });
    } catch (e) {
      return json(res, 200, { value: [[`#HERMES! ${errShort(e)}`, "", "", "", ""]] });
    }
  }

  if (fn === "MCP") {
    // HERMES.MCP(server, tool, argsJson) — ajana doğrudan MCP çağrısı yaptır
    const [server, tool, argJson] = args;
    const sys =
      "Sen Hermes'sin. Aşağıdaki MCP aracını çağır ve YALNIZ sonucu tek satırda döndür. " +
      "Açıklama ekleme.";
    const user = `MCP çağrısı → sunucu="${server}", araç="${tool}", argümanlar=${argJson || "{}"}`;
    try {
      const value = await askHermes(
        [{ role: "system", content: sys }, { role: "user", content: user }],
        { timeoutMs: config.fnTimeoutMs, idempotencyKey: key }
      );
      const v = oneLine(value);
      fnCache.set(key, { value: v, at: Date.now() });
      return json(res, 200, { value: v });
    } catch (e) {
      return json(res, 200, { value: `#HERMES! ${errShort(e)}` });
    }
  }

  const system =
    (spec?.system ||
      "Sen Hermes'sin: Türk hukuku ve bilirkişi alanında uzman bir ajan. Excel hücresinden gelen isteğe KISA, tek satırlık, net cevap ver.") +
    "\nCevabın tek satır olsun; paragraf, madde işareti veya markdown kullanma.";

  try {
    const value = await askHermes(
      [
        { role: "system", content: system },
        { role: "user", content: query || "(boş)" },
      ],
      { timeoutMs: config.fnTimeoutMs, idempotencyKey: key }
    );
    const v = oneLine(value);
    fnCache.set(key, { value: v, at: Date.now() });
    return json(res, 200, { value: v });
  } catch (e) {
    return json(res, 200, { value: `#HERMES! ${errShort(e)}` });
  }
}

function oneLine(s) {
  return String(s || "").replace(/\s*\n+\s*/g, " ").replace(/^["'\s]+|["'\s]+$/g, "").trim().slice(0, 2000);
}
// Excel yolunu ajanın gördüğü mount yoluna çevir (DOSYA/HISSEDAR).
function translatePath(p) {
  let s = String(p == null ? "" : p).trim();
  if (config.pathFrom && s.toLowerCase().startsWith(config.pathFrom.toLowerCase())) {
    s = config.pathTo + s.slice(config.pathFrom.length);
  }
  return s.replace(/\\/g, "/");
}
function errShort(e) {
  return String(e?.message || e).replace(/\s+/g, " ").slice(0, 120);
}

// ── /api/chat : sohbet paneli ─────────────────────────────────────────────
const CHAT_SYSTEM = `Sen Hermes'sin: Türk hukuku ve bilirkişi alanında uzman bir ajan; bir Excel çalışma kitabına bağlısın.
Kullanıcı seninle yan panelden konuşuyor. Gerektiğinde MCP araçlarını (bilirkişi orchestrator, emsal, parsel-konum/tapu, hesap-motoru, rapor-pipeline, genel-hafıza) kullan.
Çalışma kitabını DEĞİŞTİRMEK istediğinde bunu araç çağrısıyla değil, yanıtındaki "actions" dizisiyle yap; panel bu action'ları uygular.
YALNIZ şu JSON şemasıyla yanıt ver (başka metin yazma):
{"message":"kullanıcıya kısa Türkçe cevap","actions":[ ... ]}
actions türleri: write_cells{start_cell,values[][]}, create_sheet{name,values[][]}, format_cells{range,bold?,fill?,font_color?,number_format?}, set_number_format{range,format}, clear_range{range}, read_range{range,reason}.
Hücrelere uzun paragraf yazma; kısa etiket/sayı/formül yaz. Formülleri tablo-yerel yaz (sanki tablo A1'den başlıyormuş gibi).`;

function workbookContextBlock(wb, selection) {
  const lines = [];
  if (wb?.sheets?.length) {
    lines.push("Çalışma kitabı sayfaları:");
    for (const s of wb.sheets) lines.push(`- ${s.name}: kullanılan aralık ${s.usedRange || "-"} (${s.rowCount||0}x${s.columnCount||0})`);
  }
  if (wb?.activeSheet) lines.push(`Aktif sayfa: ${wb.activeSheet}`);
  if (selection?.address) {
    lines.push(`Seçim: ${selection.address} (${selection.rowCount || 0}x${selection.columnCount || 0})`);
    if (selection.values) {
      lines.push(`Seçim değerleri (satır dizileri):\n${JSON.stringify(selection.values).slice(0, 8000)}`);
      if (selection.truncated) {
        lines.push(`NOT: Yukarıda tablonun yalnız ilk ${selection.values.length} satırı var; toplam ${selection.rowCount} satır. Tümünü görmen gerekirse kullanıcıdan daha küçük bir aralık seçmesini veya dosya analizini (=HERMES.DOSYA) kullanmasını iste.`);
      }
    } else {
      lines.push(`NOT: Seçim çok büyük (${selection.rowCount}x${selection.columnCount}), değerler gönderilmedi. Kullanıcıdan daha küçük bir aralık seçmesini iste.`);
    }
  }
  return lines.join("\n");
}

async function handleChat(req, res) {
  const body = JSON.parse((await readBody(req)) || "{}");
  const history = Array.isArray(body.history) ? body.history.slice(-12) : [];
  const ctx = workbookContextBlock(body.workbook, body.selection);

  const messages = [
    { role: "system", content: CHAT_SYSTEM },
    ...(ctx ? [{ role: "system", content: "Bağlam:\n" + ctx }] : []),
    ...history.map((m) => ({ role: m.role === "assistant" ? "assistant" : "user", content: String(m.content || "") })),
    { role: "user", content: String(body.prompt || "") },
  ];

  try {
    const raw = await askHermes(messages, { timeoutMs: config.chatTimeoutMs, maxTokens: 2048 });
    const parsed = extractJson(raw);
    if (parsed && typeof parsed === "object") {
      return json(res, 200, {
        message: String(parsed.message || "").trim() || "(boş yanıt)",
        actions: normalizeActions(parsed.actions || []),
        source: config.provider,
      });
    }
    // JSON değilse düz metni mesaj olarak göster
    return json(res, 200, { message: oneLineKeepBreaks(raw), actions: [], source: config.provider });
  } catch (e) {
    return json(res, 200, { message: `⚠️ Hermes'e ulaşılamadı: ${errShort(e)}`, actions: [], source: "error" });
  }
}
function oneLineKeepBreaks(s) {
  return String(s || "").trim().slice(0, 4000);
}

// ── /api/push : ajan → Excel (SSE üzerinden panele iletir) ────────────────
async function handlePush(req, res) {
  const body = JSON.parse((await readBody(req)) || "{}");
  const actions = normalizeActions(body.actions || []);
  const message = body.message ? String(body.message) : "";
  if (!actions.length && !message) return json(res, 400, { ok: false, error: "actions boş" });
  const delivered = pushCommands({ actions, message, ts: Date.now() });
  return json(res, 200, { ok: true, delivered, panels: clientCount() });
}

// ── router ────────────────────────────────────────────────────────────────
async function router(req, res) {
  const url = new URL(req.url, `https://${config.host}:${config.port}`);
  const path = url.pathname;

  if (req.method === "OPTIONS") {
    setCors(res);
    res.writeHead(204);
    return res.end();
  }

  // API
  if (path.startsWith("/api/")) {
    if (path === "/api/health") {
      const hermes = await healthProbe();
      return json(res, 200, {
        ok: true, service: "hermes-excel", port: config.port, provider: config.provider,
        panels: clientCount(), hermes, time: new Date().toISOString(),
      });
    }
    if (path === "/api/commands") {
      // EventSource özel header gönderemez → token'ı sorgu parametresiyle de kabul et
      const qToken = url.searchParams.get("token");
      const okAuth = authed(req) || (config.bridgeToken && qToken === config.bridgeToken);
      if (!okAuth) return json(res, 401, { ok: false, error: "yetkisiz" });
      setCors(res);
      addClient(res);
      return; // SSE açık kalır
    }
    if (!authed(req)) return json(res, 401, { ok: false, error: "yetkisiz" });
    try {
      if (path === "/api/settings" && req.method === "GET") return json(res, 200, getSettingsMasked());
      if (path === "/api/settings" && req.method === "POST") {
        const body = JSON.parse((await readBody(req)) || "{}");
        return json(res, 200, setSettings(body));
      }
      if (path === "/api/fn" && req.method === "POST") return await handleFn(req, res);
      if (path === "/api/chat" && req.method === "POST") return await handleChat(req, res);
      if (path === "/api/push" && req.method === "POST") return await handlePush(req, res);
    } catch (e) {
      return json(res, 200, { error: errShort(e), value: `#HERMES! ${errShort(e)}`, message: `⚠️ ${errShort(e)}`, actions: [] });
    }
    return json(res, 404, { ok: false, error: "bilinmeyen uç nokta" });
  }

  // statik
  return serveStatic(req, res, path);
}

// ── başlat ────────────────────────────────────────────────────────────────
const certs = findCerts();
const banner = (proto) => {
  console.log(`\n  HermesExcel köprüsü çalışıyor`);
  console.log(`  dinleme: ${proto}://${config.host}:${config.port}`);
  console.log(`  erişim : ${proto}://${config.publicHost}:${config.publicPort}`);
  console.log(`  provider = ${config.provider}   token = ${config.bridgeToken ? "açık" : "kapalı"}`);
  console.log(`  manifest'i Excel'e sideload edin (install/sideload.ps1)\n`);
};

if (certs) {
  https.createServer(certs, router).listen(config.port, config.host, () => banner("https"));
} else {
  console.warn("\n  ⚠️  TLS sertifikası bulunamadı. Office eklentileri HTTPS ister.");
  console.warn("  Çalıştırın:  npm run cert    (office-addin-dev-certs install)\n");
  console.warn("  Şimdilik geliştirme amaçlı HTTP ile başlatılıyor (Excel bağlanmayabilir).\n");
  http.createServer(router).listen(config.port, config.host, () => banner("http"));
}
