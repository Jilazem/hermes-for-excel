// Ayar yükleyici — .env dosyasını bağımlılıksız okur, ortam değişkenleriyle
// birleştirir ve normalize edilmiş bir config nesnesi döndürür.
import { readFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join, resolve } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));
export const ROOT = resolve(__dirname, "..");

function loadDotenv(path) {
  if (!existsSync(path)) return;
  const text = readFileSync(path, "utf8");
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq === -1) continue;
    const key = line.slice(0, eq).trim();
    let val = line.slice(eq + 1).trim();
    // satır sonu yorumlarını at (tırnak içinde değilse)
    if (!/^["']/.test(val)) val = val.replace(/\s+#.*$/, "").trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1);
    }
    if (!(key in process.env)) process.env[key] = val;
  }
}

loadDotenv(join(ROOT, ".env"));

const env = process.env;
const num = (v, d) => (v === undefined || v === "" ? d : Number(v));

export const config = {
  port: num(env.PORT, 8787),
  host: env.HERMES_HOST || "127.0.0.1", // GX10'da LAN'da dinlemek için: 0.0.0.0
  publicHost: env.HERMES_PUBLIC_HOST || "localhost", // Excel'in ulaşacağı ad/IP (ör. gx10.local)
  publicPort: num(env.HERMES_PUBLIC_PORT, num(env.PORT, 8787)),
  bridgeToken: env.HERMES_BRIDGE_TOKEN || "",

  provider: (env.HERMES_PROVIDER || "mock").toLowerCase(),

  gateway: {
    url: env.HERMES_GATEWAY_URL || "http://127.0.0.1:8642/v1/chat/completions",
    apiKey: env.HERMES_API_KEY || "",
    model: env.HERMES_MODEL || "hermes-agent",
  },

  anthropic: {
    url: env.HERMES_ANTHROPIC_URL || "https://api.anthropic.com/v1/messages",
    token: env.HERMES_ANTHROPIC_TOKEN || "",
    model: env.HERMES_ANTHROPIC_MODEL || "claude-opus-4-8",
    version: "2023-06-01",
  },

  fnCacheTtlMs: num(env.HERMES_FN_CACHE_TTL_MS, 300000),
  fnTimeoutMs: num(env.HERMES_FN_TIMEOUT_MS, 120000),
  // Sohbet paneli ajanı: araç/rapor işleri uzun sürebilir → hücre
  // fonksiyonlarından daha uzun bir zaman aşımı ver (varsayılan 240 sn).
  chatTimeoutMs: num(env.HERMES_CHAT_TIMEOUT_MS, 240000),

  // Yol çevirisi: Excel'deki Windows yolunu (ör. G:\Drive'ım\) ajanın gördüğü
  // mount yoluna (ör. /home/kullanici/WinShare/) çevirir. DOSYA/HISSEDAR için.
  pathFrom: env.HERMES_PATH_FROM || "",
  pathTo: env.HERMES_PATH_TO || "",

  tls: {
    cert: env.HERMES_TLS_CERT || "",
    key: env.HERMES_TLS_KEY || "",
  },

  root: ROOT,
};

export function mcpMap() {
  try {
    const raw = readFileSync(join(ROOT, "config", "mcp-map.json"), "utf8");
    return JSON.parse(raw);
  } catch {
    return { functions: {} };
  }
}
