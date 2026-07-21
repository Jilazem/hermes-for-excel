#!/usr/bin/env node
// GX10 uzak dağıtım hazırlığı (Linux/mac/Win — sadece Node + openssl gerekir).
// - GX10'un adresine (DNS/IP) uygun SAN'lı TLS sertifikası üretir.
// - manifest.xml'i o adrese göre yeniden yazıp dist/manifest.xml üretir.
//
// Kullanım:
//   node install/configure-remote.mjs --host gx10.local [--port 8787]
//   node install/configure-remote.mjs --host 192.168.1.50
//   node install/configure-remote.mjs --host gx10.tailXXXX.ts.net --tailscale
import { execFileSync } from "node:child_process";
import { mkdirSync, writeFileSync, readFileSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const opt = (name, def) => {
  const i = args.indexOf(name);
  if (i < 0) return def;
  const v = args[i + 1];
  return v && !v.startsWith("--") ? v : true;
};

const host = opt("--host", process.env.HERMES_PUBLIC_HOST);
const port = Number(opt("--port", process.env.HERMES_PUBLIC_PORT || 8787));
const useTailscale = !!opt("--tailscale", false);
if (!host) {
  console.error("Hata: --host gerekli. Ör: --host gx10.local  |  --host 192.168.1.50");
  process.exit(1);
}
const isIp = /^\d{1,3}(\.\d{1,3}){3}$/.test(host);
const certs = join(ROOT, "certs");
const dist = join(ROOT, "dist");
mkdirSync(certs, { recursive: true });
mkdirSync(dist, { recursive: true });
const p = (d, f) => join(d, f);

// 1) Host'a göre manifest üret ------------------------------------------------
const manifest = readFileSync(join(ROOT, "manifest.xml"), "utf8");
const remapped = manifest.split("https://localhost:8787").join(`https://${host}:${port}`);
writeFileSync(p(dist, "manifest.xml"), remapped);
console.log(`✓ manifest → dist/manifest.xml   (URL tabanı: https://${host}:${port})`);

// 2) TLS sertifikaları --------------------------------------------------------
if (useTailscale) {
  execFileSync("tailscale", ["cert", "--cert-file", p(certs, "server.crt"), "--key-file", p(certs, "server.key"), host], { stdio: "inherit" });
  console.log("✓ Tailscale gerçek sertifikası üretildi — Win11'de CA güvenmeye GEREK YOK.");
} else {
  const san = isIp ? `IP:${host}` : `DNS:${host}`;
  // Kendi CA — VARSA yeniden üretme (Win11'deki güveni bozmamak için)
  if (existsSync(p(certs, "ca.crt")) && existsSync(p(certs, "ca.key"))) {
    console.log("↺ mevcut CA yeniden kullanılıyor (certs/ca.crt) — Win11 güveni korunur");
  } else {
    execFileSync("openssl", ["req", "-x509", "-newkey", "rsa:2048", "-nodes",
      "-keyout", p(certs, "ca.key"), "-out", p(certs, "ca.crt"), "-days", "3650",
      "-subj", "/CN=HermesExcel Dev CA/O=HermesExcel"], { stdio: "inherit" });
  }
  // Server anahtarı + CSR
  execFileSync("openssl", ["req", "-newkey", "rsa:2048", "-nodes",
    "-keyout", p(certs, "server.key"), "-out", p(certs, "server.csr"),
    "-subj", `/CN=${host}`], { stdio: "inherit" });
  // SAN uzantısı ile imzala
  writeFileSync(p(certs, "san.ext"), `subjectAltName=${san}\nextendedKeyUsage=serverAuth\nkeyUsage=digitalSignature,keyEncipherment\n`);
  // NOT: leaf sertifika geçerliliği ≤398 gün olmalı (Chromium/WebView2 ERR_CERT_VALIDITY_TOO_LONG)
  execFileSync("openssl", ["x509", "-req", "-in", p(certs, "server.csr"),
    "-CA", p(certs, "ca.crt"), "-CAkey", p(certs, "ca.key"), "-CAcreateserial",
    "-out", p(certs, "server.crt"), "-days", "397", "-extfile", p(certs, "san.ext")], { stdio: "inherit" });
  console.log(`✓ Özel sertifika üretildi: certs/server.crt (SAN=${san}) + certs/ca.crt`);
}

console.log(`
────────────────────────────────────────────────────────────────────
GX10 .env (bu dosyayı oluşturun/düzenleyin):
  HERMES_HOST=0.0.0.0
  HERMES_PUBLIC_HOST=${host}
  HERMES_PUBLIC_PORT=${port}
  HERMES_TLS_CERT=./certs/server.crt
  HERMES_TLS_KEY=./certs/server.key
  HERMES_PROVIDER=gateway
  HERMES_GATEWAY_URL=http://127.0.0.1:8642/v1/chat/completions
  HERMES_API_KEY=<Hermes API_SERVER_KEY>
  HERMES_MODEL=hermes-agent

GX10'da köprüyü başlatın:
  node bridge/server.mjs        (kalıcı için: pm2 / systemd — bkz docs/REMOTE-GX10.md)

Win11'de:
${useTailscale
  ? "  1) (CA güven adımı YOK — Tailscale gerçek sertifika)\n"
  : "  1) certs/ca.crt → Win11'e kopyalayın, sonra: install\\trust-ca.ps1 -CaPath ca.crt\n"}  2) dist/manifest.xml → Win11'e kopyalayın, sonra: install\\sideload.ps1 -Manifest <yol>\\manifest.xml
  3) GX10 köprüsü çalışırken Excel'i açın → Hermes grubu → hücrede =HERMES.SOR("...")
────────────────────────────────────────────────────────────────────`);
