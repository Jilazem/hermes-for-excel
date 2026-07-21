import { test } from "node:test";
import assert from "node:assert/strict";
import { rebaseFormula, normalizeActions, extractJson, numToCol, colToNum } from "./actions.mjs";

test("col <-> num", () => {
  assert.equal(colToNum("A"), 1);
  assert.equal(colToNum("Z"), 26);
  assert.equal(colToNum("AA"), 27);
  assert.equal(numToCol(1), "A");
  assert.equal(numToCol(27), "AA");
});

test("rebase basit formül", () => {
  // A1-yerel =B2*C2, H23'e yazılınca =I24*J24 olmalı
  assert.equal(rebaseFormula("=B2*C2", "H23"), "=I24*J24");
});

test("rebase $ sabitleri kaydırmaz", () => {
  assert.equal(rebaseFormula("=$B$2", "H23"), "=$B$2");
  assert.equal(rebaseFormula("=B$2", "H23"), "=I$2");
  assert.equal(rebaseFormula("=$B2", "H23"), "=$B24");
});

test("rebase sayfa-nitelikli referansı kaydırmaz", () => {
  assert.equal(rebaseFormula("=Sheet1!B5", "H23"), "=Sheet1!B5");
  assert.equal(rebaseFormula("='Bank Rec'!A1", "H23"), "='Bank Rec'!A1");
});

test("rebase string literal korunur", () => {
  assert.equal(rebaseFormula('="B2 hücresi"', "H23"), '="B2 hücresi"');
});

test("A1'de kaydırma yok", () => {
  assert.equal(rebaseFormula("=B2*C2", "A1"), "=B2*C2");
});

test("normalizeActions bilinmeyeni atar, write_cells rebaseler", () => {
  const out = normalizeActions([
    { type: "write_cells", start_cell: "H23", values: [["=B2*C2", 5]] },
    { type: "execute_office_js", code: "hack()" },
  ]);
  assert.equal(out.length, 1);
  assert.equal(out[0].values[0][0], "=I24*J24");
  assert.equal(out[0].values[0][1], 5);
});

test("extractJson fence + gevşek", () => {
  assert.deepEqual(extractJson('```json\n{"a":1}\n```'), { a: 1 });
  assert.deepEqual(extractJson('önek {"a":1,} sonek'), { a: 1 });
  assert.equal(extractJson("json yok"), null);
});
