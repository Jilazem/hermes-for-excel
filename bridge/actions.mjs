// Action normalizasyonu + formül yeniden konumlama (formula rebasing).
// Ajan, tablo formüllerini "tablo-yerel" yazar (sol-üst = A1). Köprü bunları
// hedef start_cell'e göre kaydırır. $-sabitleri ve sayfa-nitelikli referanslar
// kaydırılmaz.

const COL_RE = /^[A-Za-z]{1,3}$/;

export function colToNum(col) {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function numToCol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// "H23" -> {col:8,row:23}
export function parseCell(addr) {
  const m = String(addr).replace(/\$/g, "").match(/^([A-Za-z]{1,3})(\d+)$/);
  if (!m) return null;
  return { col: colToNum(m[1]), row: Number(m[2]) };
}

// Tek bir A1 referansını (ör. B2, $B$2, C$3) dcol/row deltasına göre kaydır.
function shiftRef(ref, dCol, dRow) {
  const m = ref.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
  if (!m) return ref;
  const [, colAbs, col, rowAbs, row] = m;
  const newCol = colAbs === "$" ? col : numToCol(colToNum(col) + dCol);
  const newRow = rowAbs === "$" ? row : String(Number(row) + dRow);
  return `${colAbs}${newCol}${rowAbs}${newRow}`;
}

// Bir formül metnindeki (=... ) tablo-yerel referansları kaydır.
// Sayfa-nitelikli referanslar ('Sheet'!A1, Sheet1!B2) ve string literaller korunur.
export function rebaseFormula(formula, startCell) {
  if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
  const anchor = parseCell(startCell);
  if (!anchor) return formula;
  const dCol = anchor.col - 1; // A1 tabanlı
  const dRow = anchor.row - 1;
  if (dCol === 0 && dRow === 0) return formula;

  const s = formula;
  let out = "";
  let i = 0;
  const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
  // 'Ad'!A1 veya Ad!A1 (tek hücre ya da A1:B2 aralığı) -> hiç kaydırılmaz
  const QUALIFIED = new RegExp(`^(?:'[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!${CELL}(?::${CELL})?`);
  const PLAIN = new RegExp(`^${CELL}`);
  const WORD = /[A-Za-z0-9_]/;

  while (i < s.length) {
    const ch = s[i];
    const prev = s[i - 1] || "";
    const rest = s.slice(i);

    // string literal
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      const j = end === -1 ? s.length : end + 1;
      out += s.slice(i, j);
      i = j;
      continue;
    }
    // sayfa-nitelikli referans (tırnaklı veya tırnaksız) — kaydırma
    const qm = rest.match(QUALIFIED);
    if (qm && !WORD.test(prev) && prev !== "!") {
      out += qm[0];
      i += qm[0].length;
      continue;
    }
    // düz A1 referansı — kaydır (öncesinde harf/rakam/_/'/! yoksa ve sonrası "!" değilse)
    const pm = rest.match(PLAIN);
    if (pm && !WORD.test(prev) && prev !== "!" && prev !== "'") {
      const after = s[i + pm[0].length];
      if (after !== "!") {
        out += shiftRef(pm[0], dCol, dRow);
        i += pm[0].length;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return out;
}

// values matrisindeki formülleri (= ile başlayan hücreler) yeniden konumla.
function rebaseValues(values, startCell) {
  if (!Array.isArray(values)) return values;
  return values.map((row) =>
    Array.isArray(row)
      ? row.map((cell) => (typeof cell === "string" && cell.startsWith("=") ? rebaseFormula(cell, startCell) : cell))
      : row
  );
}

const KNOWN = new Set([
  "write_cells", "create_sheet", "format_cells", "set_number_format",
  "clear_range", "read_range", "message", "conditional_format",
]);

// Ajan yanıtındaki action listesini normalize et.
export function normalizeActions(actions) {
  if (!Array.isArray(actions)) return [];
  const out = [];
  for (const a of actions) {
    if (!a || typeof a !== "object") continue;
    const type = a.type;
    if (!KNOWN.has(type)) continue;
    const n = { ...a, type };
    if (type === "write_cells") {
      n.start_cell = a.start_cell || a.cell || "A1";
      n.values = rebaseValues(a.values || [], n.start_cell);
      n.allow_overwrite = a.allow_overwrite !== false;
    } else if (type === "create_sheet") {
      n.name = String(a.name || "Sayfa");
      n.values = rebaseValues(a.values || [], "A1");
    }
    out.push(n);
  }
  return out;
}

// Ajan bazen JSON'u ```json blokları içinde veya gevşek biçimde döndürür.
export function extractJson(text) {
  if (typeof text !== "string") return null;
  let t = text.trim();
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) t = fence[1].trim();
  const start = t.indexOf("{");
  const end = t.lastIndexOf("}");
  if (start === -1 || end === -1 || end < start) return null;
  const slice = t.slice(start, end + 1);
  try {
    return JSON.parse(slice);
  } catch {
    // basit onarım: sondaki virgülleri kaldır
    try {
      return JSON.parse(slice.replace(/,\s*([}\]])/g, "$1"));
    } catch {
      return null;
    }
  }
}
