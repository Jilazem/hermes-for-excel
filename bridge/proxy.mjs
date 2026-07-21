// Win11 localhost köprüsü (hibrit) — yalnız uzak dağıtımda gereklidir.
//  - Add-in STATİK dosyalarını (taskpane, functions, assets...) YEREL sunar.
//  - /api/* isteklerini uzak sunucudaki köprüye iletir (SSE dahil).
// Neden: Office add-in WebView2'si LAN IP'sine erişemez ama localhost'a erişir.
// Statik dosyaları yerel sunmak, uzak dosya senkronundan bağımsız kılar; sadece
// gerçek API çağrıları LAN üzerinden uzak köprüye gider.
//
//  PROXY_PORT      (8799)                        localhost dinleme portu
//  PROXY_UPSTREAM  (https://<sunucu-adı>:8799)    uzak köprü (/api/*) — ZORUNLU, örn: https://192.168.1.50:8799
//  WWW_ROOT        (./addin veya C:\HermesExcel\www)  yerel add-in dosyaları
import https from "node:https";
import { readFileSync, existsSync, appendFileSync, readFile } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, extname, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const LISTEN_PORT = Number(process.env.PROXY_PORT || 8799);
const UPSTREAM = process.env.PROXY_UPSTREAM;
if (!UPSTREAM) {
  console.error("Hata: PROXY_UPSTREAM ayarlanmadı. Örnek: PROXY_UPSTREAM=https://192.168.1.50:8799 node bridge/proxy.mjs");
  console.error("(Tek makinede çalışıyorsanız bu proxy'ye gerek yok — doğrudan bridge/server.mjs'i kullanın.)");
  process.exit(1);
}
const WWW = process.env.WWW_ROOT || join(dirname(fileURLToPath(import.meta.url)), "..", "addin");
const up = new URL(UPSTREAM);

const LOG_FILE = process.env.PROXY_LOG || join(tmpdir(), "hermes-proxy.log");
const logLine = (s) => { try { appendFileSync(LOG_FILE, `${new Date().toISOString()} ${s}\n`); } catch {} };

const certDir = join(homedir(), ".office-addin-dev-certs");
const crt = join(certDir, "localhost.crt");
const key = join(certDir, "localhost.key");
if (!existsSync(crt) || !existsSync(key)) { console.error("localhost dev sertifikası yok: npm run cert"); process.exit(1); }
const tls = { cert: readFileSync(crt), key: readFileSync(key) };

const MIME = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8",
  ".png": "image/png", ".ico": "image/x-icon", ".svg": "image/svg+xml", ".map": "application/json",
};

function forwardToGx10(req, res) {
  const fwd = https.request(
    { host: up.hostname, port: up.port || 443, path: req.url, method: req.method,
      headers: { ...req.headers, host: up.host }, rejectUnauthorized: false },
    (upRes) => { logLine(`API ${req.method} ${req.url} -> ${upRes.statusCode}`); res.writeHead(upRes.statusCode || 502, upRes.headers); upRes.pipe(res); }
  );
  fwd.on("error", (e) => { logLine(`API ${req.method} ${req.url} -> HATA ${e.message}`); if (!res.headersSent) res.writeHead(502, { "content-type": "application/json" }); res.end(JSON.stringify({ error: "proxy: GX10'a ulaşılamadı — " + e.message })); });
  req.pipe(fwd);
}

const handler = (req, res) => {
  const path = decodeURIComponent(req.url.split("?")[0]);
  // Statik add-in dosyaları YEREL; /api/* ve bilinmeyenler GX10'a
  if (req.method === "GET" && !path.startsWith("/api/")) {
    const rel = (path === "/" ? "/taskpane.html" : path).replace(/\\/g, "/").replace(/\.\.+/g, "");
    const file = join(WWW, rel);
    return readFile(file, (err, data) => {
      if (err) { forwardToGx10(req, res); return; } // yerelde yoksa GX10'a sor
      logLine(`STATIC ${path} -> 200`);
      res.writeHead(200, { "content-type": MIME[extname(file).toLowerCase()] || "application/octet-stream", "access-control-allow-origin": "*", "cache-control": "no-store" });
      res.end(data);
    });
  }
  forwardToGx10(req, res);
};

console.log(`\n  HermesExcel proxy  (statik: yerel ${WWW} | /api/*: ${UPSTREAM})`);
for (const host of ["127.0.0.1", "::1"]) {
  const s = https.createServer(tls, handler);
  s.on("error", (e) => console.log(`  ${host}:${LISTEN_PORT} dinlenemedi: ${e.message}`));
  s.listen(LISTEN_PORT, host, () => console.log(`  dinliyor: https://${host === "::1" ? "[::1]" : host}:${LISTEN_PORT}  (localhost)`));
}
