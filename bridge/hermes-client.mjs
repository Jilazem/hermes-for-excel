// Hermes provider soyutlaması.
// Bir "messages" dizisi alır, ajanın metin yanıtını döndürür.
// Ayarlar çalışma-zamanı deposundan gelir (Ayarlar paneli → .env varsayılanları).
// provider: gateway/openai (OpenAI-uyumlu) | anthropic (Claude Messages) | mock
import { config } from "./config.mjs";
import { getSettings } from "./settings.mjs";

const enc = (s) => (typeof s === "string" ? s : JSON.stringify(s));

async function withTimeout(promise, ms) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  try {
    return await promise(ctrl.signal);
  } finally {
    clearTimeout(t);
  }
}

// ── gateway / openai: OpenAI-uyumlu chat/completions ──────────────────────
// (Hermes API Server, OpenAI, Ollama, LM Studio, LocalAI... hepsi bu şema)
async function askGateway(messages, opts, s) {
  const body = {
    model: opts.model || s.model,
    messages,
    stream: false,
    temperature: opts.temperature ?? 0.2,
  };
  const res = await withTimeout(
    (signal) =>
      fetch(s.gatewayUrl, {
        method: "POST",
        signal,
        headers: {
          "content-type": "application/json",
          ...(s.apiKey ? { authorization: `Bearer ${s.apiKey}` } : {}),
          // HTTP başlıkları Latin1 (≤255) olmalı; Türkçe karakterli anahtarı base64'le
          ...(opts.idempotencyKey
            ? { "idempotency-key": Buffer.from(String(opts.idempotencyKey), "utf8").toString("base64").slice(0, 200) }
            : {}),
        },
        body: JSON.stringify(body),
      }),
    opts.timeoutMs || config.fnTimeoutMs
  );
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new Error(`gateway ${res.status}: ${detail.slice(0, 300)}`);
  }
  const data = await res.json();
  const text = data?.choices?.[0]?.message?.content;
  if (typeof text !== "string") throw new Error("gateway: beklenmeyen yanıt biçimi");
  return text.trim();
}

// ── anthropic: Claude Messages API ────────────────────────────────────────
async function askAnthropic(messages, opts, s) {
  const system = messages.filter((m) => m.role === "system").map((m) => enc(m.content)).join("\n\n");
  const conv = messages
    .filter((m) => m.role !== "system")
    .map((m) => ({ role: m.role === "assistant" ? "assistant" : "user", content: enc(m.content) }));
  const body = {
    model: opts.model || s.anthropicModel,
    max_tokens: opts.maxTokens || 1024,
    ...(system ? { system } : {}),
    messages: conv,
  };
  const res = await withTimeout(
    (signal) =>
      fetch(s.anthropicUrl, {
        method: "POST",
        signal,
        headers: {
          "content-type": "application/json",
          "x-api-key": s.anthropicToken,
          "anthropic-version": config.anthropic.version,
        },
        body: JSON.stringify(body),
      }),
    opts.timeoutMs || config.fnTimeoutMs
  );
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new Error(`Anthropic ${res.status}: ${detail.slice(0, 300)}`);
  }
  const data = await res.json();
  const text = (data?.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
  if (!text) throw new Error("Anthropic: boş yanıt");
  return text.trim();
}

// ── mock: çevrimdışı test ─────────────────────────────────────────────────
function askMock(messages) {
  const last = [...messages].reverse().find((m) => m.role === "user");
  const q = enc(last?.content || "").slice(0, 120);
  const sys = messages.find((m) => m.role === "system");
  const wantsJson = /JSON|actions/i.test(enc(sys?.content || ""));
  if (wantsJson) {
    return Promise.resolve(
      JSON.stringify({ message: `(mock) İsteğinizi aldım: "${q}". Ayarlar panelinden bir sağlayıcı bağlayın.`, actions: [] })
    );
  }
  return Promise.resolve(`(mock) yanıt — "${q}". Gerçek sağlayıcı için Ayarlar panelini kullanın.`);
}

export async function askHermes(messages, opts = {}) {
  const s = getSettings();
  switch (s.provider) {
    case "gateway":
    case "openai":
      return askGateway(messages, opts, s);
    case "anthropic":
      return askAnthropic(messages, opts, s);
    case "mock":
    default:
      return askMock(messages, opts);
  }
}

export async function healthProbe() {
  const s = getSettings();
  if (s.provider === "mock") return { ok: true, status: "mock" };

  if (s.provider === "gateway" || s.provider === "openai") {
    // Model çağrısı yapmadan yokla: /v1/models (OpenAI/Ollama/LM Studio) veya
    // /v1/health (Hermes API Server). Herhangi biri OK ise sağlıklı say.
    const base = s.gatewayUrl.replace(/\/chat\/completions.*$/, "");
    const candidates = [base + "/models", base + "/health"];
    let lastErr = "";
    for (const url of candidates) {
      try {
        const res = await withTimeout(
          (signal) => fetch(url, { signal, headers: s.apiKey ? { authorization: `Bearer ${s.apiKey}` } : {} }),
          5000
        );
        if (res.ok) return { ok: true, status: `ok (${s.provider})` };
        lastErr = `${res.status}`;
      } catch (e) {
        lastErr = String(e.message || e);
      }
    }
    return { ok: false, status: `${s.provider}: ${lastErr}` };
  }

  try {
    const txt = await askHermes([{ role: "user", content: "ping" }], { timeoutMs: 8000, maxTokens: 8 });
    return { ok: true, status: "ok (anthropic)", sample: txt.slice(0, 40) };
  } catch (e) {
    return { ok: false, status: String(e.message || e) };
  }
}
