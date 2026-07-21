// Çalışma-zamanı ayar deposu. Add-in'in "Ayarlar" panelinden gelen sağlayıcı
// yapılandırmasını (gateway/anthropic/mock, URL, API anahtarı, model) saklar.
// data/settings.json'a yazılır; .env varsayılanlarını geçersiz kılar.
// Böylece başka kullanıcılar kendi gateway'lerini panelden ayarlayabilir.
import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { config, ROOT } from "./config.mjs";

const FILE = join(ROOT, "data", "settings.json");
const KEYS = ["provider", "gatewayUrl", "apiKey", "model", "anthropicUrl", "anthropicToken", "anthropicModel"];
let store = null;

function load() {
  if (store) return store;
  try {
    store = existsSync(FILE) ? JSON.parse(readFileSync(FILE, "utf8")) : {};
  } catch {
    store = {};
  }
  return store;
}

// .env varsayılanları + kayıtlı ayarlar (kayıtlı olan kazanır)
export function getSettings() {
  const s = load();
  return {
    provider: s.provider || config.provider,
    gatewayUrl: s.gatewayUrl || config.gateway.url,
    apiKey: s.apiKey ?? config.gateway.apiKey,
    model: s.model || config.gateway.model,
    anthropicUrl: s.anthropicUrl || config.anthropic.url,
    anthropicToken: s.anthropicToken ?? config.anthropic.token,
    anthropicModel: s.anthropicModel || config.anthropic.model,
  };
}

// Panelde göstermek için — sırlar maskeli
export function getSettingsMasked() {
  const s = getSettings();
  const mask = (v) => (v ? "••••••" + String(v).slice(-4) : "");
  return {
    provider: s.provider,
    gatewayUrl: s.gatewayUrl,
    apiKeySet: !!s.apiKey,
    apiKeyMasked: mask(s.apiKey),
    model: s.model,
    anthropicUrl: s.anthropicUrl,
    anthropicTokenSet: !!s.anthropicToken,
    anthropicModel: s.anthropicModel,
  };
}

export function setSettings(patch) {
  const s = load();
  for (const k of KEYS) {
    if (patch[k] !== undefined && patch[k] !== null) {
      // boş string sırrı silmez (kullanıcı temizlemek isterse "__CLEAR__" gönderir)
      if (patch[k] === "__CLEAR__") delete s[k];
      else if (patch[k] !== "") s[k] = patch[k];
      else if (k !== "apiKey" && k !== "anthropicToken") s[k] = patch[k];
    }
  }
  store = s;
  try {
    mkdirSync(join(ROOT, "data"), { recursive: true });
    writeFileSync(FILE, JSON.stringify(s, null, 2));
  } catch {
    /* yazılamazsa bellek içi kalır */
  }
  return getSettingsMasked();
}
