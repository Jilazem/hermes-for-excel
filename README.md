# Hermes for Excel

[![Built with Claude](https://img.shields.io/badge/Built%20with-Claude-8A2BE2?logo=anthropic&logoColor=white)](https://claude.com/claude-code)
[![Anthropic](https://img.shields.io/badge/Anthropic-Claude-191919?logo=anthropic&logoColor=white)](https://www.anthropic.com)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

> 🤖 Designed and built with **[Claude](https://claude.com/claude-code)** by **[Anthropic](https://www.anthropic.com)** (@AnthropicAI).

Bring an **AI agent (or any LLM)** into Microsoft Excel — as **in-cell functions**
and a **chat side panel**. Ask questions from a cell with
`=HERMES.SOR("…")`, chat with the agent about your sheet, let it **fill and
format the workbook**, watch your steps and **suggest an automation**, and detect
your **live selection** — all wired to a provider you configure (a self-hosted
Hermes agent, OpenAI, Ollama, LM Studio, or Anthropic Claude).

> 🌍 UI in **Turkish / English / 中文** (add your own language easily).
> 🇹🇷 Türkçe belgeler: [README.tr.md](README.tr.md) · [docs/CALISAN-MIMARI.md](docs/CALISAN-MIMARI.md)

```
Excel (localhost proxy)  ──HTTPS──▶  Bridge (Node)  ──▶  Your LLM / Agent
  =HERMES.* functions                 /api/fn,/chat                (Hermes gateway,
  chat pane · selection · watch       + workbook actions            OpenAI, Ollama,
  settings · i18n                                                   Anthropic…)
```

---

## Features

- **In-cell functions** — `=HERMES.SOR`, `.EMSAL`, `.TAPU`, `.PARSEL`, `.HESAP`,
  `.RAPOR`, `.HAFIZA`, `.MCP`, plus generic `.SINIFLA` (classify), `.CIKAR`
  (extract), `.OZETLE` (summarize), `.FORMUL` (formula help). Streaming, cached,
  drag-fillable.
- **Chat side panel** — talk about the active sheet; the agent replies and can
  **write cells, create sheets, format ranges** via a structured actions
  contract (applied by the pane, never by raw agent tools).
- **Live selection detection** (like the Claude add-in) — the panel shows your
  current selection and sends it to the agent.
- **Watch → Automate** — press ⏺, do a few steps in Excel; each change is logged
  live; press ⏹ and the agent suggests a **formula / VBA macro / applyable
  steps** to automate it.
- **Settings panel (⚙)** — pick provider + gateway URL + API key + model in the
  UI; stored on the bridge, no code edits. Works with OpenAI-compatible
  endpoints (self-hosted agent, OpenAI, **Ollama**, LM Studio) and Anthropic.
- **Multilingual UI** — TR / EN / ZH, switchable in Settings; trivially
  extensible.
- **Agent → Excel push** — the agent can drive the workbook out-of-band via a
  documented MCP tool (`docs/HERMES_MCP_TOOL.md`).

## Requirements

- **Excel desktop** (Windows; Microsoft 365) with add-in sideloading allowed.
- **Node.js 18+** (the bridge/proxy is zero-dependency).
- A reachable **LLM/agent endpoint** (OpenAI-compatible or Anthropic). For a
  fully local test, `mock` mode needs nothing.

---

## Quick start (everything on one Windows machine)

```powershell
# 1) Dev HTTPS certificate (Office requires HTTPS, even for localhost)
npm install
npm run cert                         # office-addin-dev-certs install

# 2) Configure the bridge
copy .env.example .env               # then edit .env (provider, URL, key)
npm start                            # bridge on https://localhost:8799

# 3) Sideload the add-in into Excel (no admin needed)
npm run sideload                     # office-addin-dev-settings sideload
```

Open Excel → **Home → Hermes → “Hermes'i Aç”** for the pane, or type
`=HERMES.SOR("hello")` in a cell. Configure your provider from the pane's **⚙**.

> On some Microsoft 365 builds a plain registry sideload doesn't reliably load
> the add-in on a normal launch. Use the launcher (below), which sideloads and
> opens Excel with the add-in every time.

### One-click launcher (recommended on Windows)

`install/baslat.bat` sideloads the add-in and opens Excel with it loaded:

```
npx office-addin-dev-settings sideload "<path>\manifest.xml" desktop --app Excel
```

Create a desktop shortcut to a hidden wrapper (`install/proxy-launcher.vbs`
pattern) and double-click it — Excel opens with Hermes ready, no debug dialog.

---

## Remote deployment (bridge on a server, Excel on your PC)

**Why a proxy?** Office add-in WebView2 **cannot reach a LAN IP** (e.g.
`192.168.1.50`) — only `localhost` or a public URL. So the add-in is served from
a **localhost proxy** on the Excel machine, which forwards `/api/*` to the remote
bridge over the LAN.

```
Windows PC (Excel)                          Server (Linux/ARM/etc.)
  Excel add-in ← localhost:8799 proxy ──LAN──▶ bridge :8799 ──▶ agent/LLM :xxxx
                 (serves add-in files,
                  forwards /api/* to server)
```

- **Server:** run `bridge/server.mjs` with a SAN certificate for the server's
  address (`install/configure-remote.mjs --host <ip-or-name>`; **≤398-day** leaf,
  Chromium rejects longer). Bind `HERMES_HOST=0.0.0.0`.
- **PC:** run `bridge/proxy.mjs` (serves `www/` locally, forwards `/api/*` to the
  server). Trust the server CA in **`LocalMachine\Root`** (`certutil -addstore -f
  Root ca.crt`) — the WebView2 sandbox needs the machine store, not just the user
  store.

Full guide + the exact commands: [docs/CALISAN-MIMARI.md](docs/CALISAN-MIMARI.md)
and [docs/REMOTE-GX10.md](docs/REMOTE-GX10.md).

> **Tailscale** gives the cleanest remote path: a real (Let's Encrypt) cert and a
> stable `*.ts.net` name — no CA install on the client
> (`configure-remote.mjs --host <name>.ts.net --tailscale`).

---

## Configuration

`.env` (server-side defaults) — every value is also overridable from the **⚙
Settings** panel at runtime:

| Variable | Meaning |
|---|---|
| `PORT` (8799) | Bridge port (loopback only unless `HERMES_HOST` set) |
| `HERMES_HOST` | `0.0.0.0` to listen on the LAN (remote setup) |
| `HERMES_BRIDGE_TOKEN` | If set, every `/api/*` call needs `X-Hermes-Token` |
| `HERMES_PROVIDER` | `gateway`/`openai` · `anthropic` · `mock` |
| `HERMES_GATEWAY_URL` | OpenAI-compatible `…/v1/chat/completions` |
| `HERMES_API_KEY` | Bearer key for the gateway |
| `HERMES_MODEL` | e.g. `hermes-agent`, `gpt-4o`, `llama3` |

Providers via the Settings panel: **self-hosted agent** (Hermes API Server),
**OpenAI** (`https://api.openai.com/v1/chat/completions`), **Ollama**
(`http://localhost:11434/v1/chat/completions`), **LM Studio**, **Anthropic**.

---

## Custom function reference

| Function | Purpose |
|---|---|
| `=HERMES.SOR(q; [ctx])` | Free-form question; the agent picks tools and answers |
| `=HERMES.SINIFLA(v; rule)` | Classify a value → label |
| `=HERMES.CIKAR(v; what)` | Extract a field from text |
| `=HERMES.OZETLE(range)` | One-line summary of a range |
| `=HERMES.FORMUL(goal)` | Suggest an Excel formula |
| `=HERMES.MCP(server; tool; argsJSON)` | Generic MCP passthrough |
| `.EMSAL .TAPU .PARSEL .HESAP .RAPOR .HAFIZA` | Domain tools (map in `config/mcp-map.json`) |

Each shows `⏳` while working; errors return `#HERMES! …` (never fabricated
data). Argument separator is `;` or `,` depending on your locale.

---

## Troubleshooting (hard-won)

| Symptom | Cause → Fix |
|---|---|
| `#NAME?` / `#AD?` for `=HERMES.*` | `CustomFunctions` ExtensionPoint must be under **`<AllFormFactors>`** in the manifest, not `<DesktopFormFactor>`. (Manifest validators don't catch this; the runtime log does.) |
| Panel stuck “Loading”, no requests hit the proxy | `localhost` resolves to **IPv6 `::1`** first — the proxy must listen on **both** `127.0.0.1` and `::1`. |
| “Add-in could not download a required resource” | (a) cert validity **> 398 days** → Chromium rejects; regenerate ≤397 days. (b) CA only in `CurrentUser\Root` → add to **`LocalMachine\Root`**. (c) add-in served from a **LAN IP** → use the localhost proxy. |
| `#HERMES! Cannot convert argument to a ByteString … 255` | Non-Latin1 (Turkish/Chinese) chars in an HTTP header. Fixed: the idempotency key is base64-encoded. |
| Custom function uses the wrong URL | In the custom-functions runtime `self.location.origin` is unreliable → the bridge URL is pinned in `functions.js`. |
| Old code keeps running after an edit | Office caches custom functions in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`. Close Excel, delete that folder, re-sideload. |
| A “Debug event-based handler” dialog keeps popping | Debugging is enabled. Use `office-addin-dev-settings sideload` (not `office-addin-debugging start`), and remove `WEF\Developer\<id>` debug subkeys. |
| Doesn't load on a normal Excel launch | Use the launcher (sideload + open) each session. For a permanent install, host the add-in on a **public HTTPS** URL. |

Diagnostic logs: proxy → `%TEMP%\hermes-proxy.log`; Office add-in →
`%TEMP%\OfficeAddins.log.txt`.

---

## Add your language

Open `addin/i18n.js`. Ask Hermes to translate the `tr` block
(*“translate this JSON to German”*), add it under a new code (`de`, `fr`, `es`…),
add the code to `LANG_NAMES`, save, and reload the pane — it appears in the
Settings language menu automatically.

## Project layout

```
manifest.xml           Office manifest (task pane + custom functions)
addin/                 taskpane (chat, selection, watch, settings), functions, i18n
bridge/                server.mjs, proxy.mjs, hermes-client, settings, actions, command-bus
config/mcp-map.json    HERMES.* → provider/MCP mapping
install/               cert, sideload, launcher, remote cert+manifest generator
docs/                  architecture, remote setup, MCP tool contract (TR)
```

## Credits & license

**Designed and written with [Claude](https://claude.com/claude-code) — Anthropic's
AI assistant** ([Anthropic](https://www.anthropic.com), [@AnthropicAI](https://x.com/AnthropicAI)).
The architecture, bridge, custom functions, task pane, i18n and this documentation
were built in a pair-programming session with Claude (Claude Code).

Design informed by [`lEWFkRAD/hermes-excel-sidecar`](https://github.com/lEWFkRAD/hermes-excel-sidecar)
(task-pane + Node-bridge pattern, action contract, formula rebasing) and
[`tonbistudio/hermes-office`](https://github.com/tonbistudio/hermes-office)
(OpenAI-compatible API Server contract, idempotency, generic functions).

MIT — see [LICENSE](LICENSE).
