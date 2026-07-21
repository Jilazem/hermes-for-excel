# Hermes → Excel: MCP aracı sözleşmesi

Hermes ajanının **kendiliğinden** (sohbet dışında da) çalışma kitabını
sürebilmesi için Hermes tarafına küçük bir MCP aracı ekleyin. Bu araç, köprünün
`POST /api/push` uç noktasına bir action listesi gönderir; köprü bunu açık SSE
bağlantısı üzerinden Excel paneline iletir ve panel action'ları uygular.

## Uç nokta

```
POST https://localhost:8787/api/push
Content-Type: application/json
X-Hermes-Token: <HERMES_BRIDGE_TOKEN varsa>

{
  "message": "Kamulaştırma tablosu dolduruldu.",
  "actions": [
    { "type": "create_sheet", "name": "Değerleme",
      "values": [["Parsel","m2","Birim","Bedel"],
                 ["123/4", 850, 12000, "=B2*C2"]] },
    { "type": "format_cells", "range": "A1:D1", "bold": true, "fill": "#7c3aed", "font_color": "#ffffff" },
    { "type": "set_number_format", "range": "D2:D2", "format": "#,##0 ₺" }
  ]
}
```

Yanıt: `{ "ok": true, "delivered": <panel_sayısı> }`. `delivered` 0 ise açık
panel yoktur (kullanıcı Excel'de Hermes panelini açmalı).

## Action türleri

| type | alanlar |
|---|---|
| `write_cells` | `start_cell`, `values[][]`, `sheet?`, `allow_overwrite?` |
| `create_sheet` | `name`, `values[][]?` |
| `format_cells` | `range`, `sheet?`, `bold?`, `italic?`, `fill?`, `font_color?`, `align?`, `number_format?` |
| `set_number_format` | `range`, `format` |
| `clear_range` | `range` |

Formüller **tablo-yerel** yazılır (tablo sol-üstü A1 kabul edilir); köprü
`start_cell`'e göre yeniden konumlar.

## Örnek MCP aracı (Python, FastMCP)

Hermes'in mevcut MCP sunucularından birine ekleyebileceğiniz asgari araç:

```python
import httpx
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("excel-bridge")
BRIDGE = "https://localhost:8787/api/push"
TOKEN = ""  # HERMES_BRIDGE_TOKEN

@mcp.tool()
async def excel_yaz(actions: list, message: str = "") -> str:
    """Excel çalışma kitabına action listesi gönderir (write_cells, create_sheet, format_cells...)."""
    headers = {"content-type": "application/json"}
    if TOKEN:
        headers["x-hermes-token"] = TOKEN
    async with httpx.AsyncClient(verify=False) as c:  # yerel dev sertifikası
        r = await c.post(BRIDGE, json={"actions": actions, "message": message}, headers=headers, timeout=30)
        return r.text
```

Bunu `bilirkisi-orchestrator` veya ayrı bir `excel-bridge` MCP sunucusu olarak
kaydedin. Artık Hermes, herhangi bir görev sırasında `excel_yaz(...)` çağırıp
Excel'i doğrudan güncelleyebilir.

> Not: Yerel dev sertifikası için `verify=False` (veya CA'yı güvenilir yapın).
> Üretimde `HERMES_BRIDGE_TOKEN` verin ki yalnız yetkili çağrılar geçsin.
