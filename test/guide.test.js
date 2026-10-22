import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

import { GUIDE_SHEETS, GUIDE_NAMES, guideFormulas } from "../src/lib/guide-content.js";

const functionsJson = JSON.parse(fs.readFileSync("src/functions/functions.json", "utf8"));
const PARAMS = Object.fromEntries(functionsJson.functions.map((f) => [f.id, f.parameters]));
const EXCEL_FUNCTIONS = new Set(["INDEX", "FORMULATEXT"]);

/** Split the argument text of a call at top-level commas (outside quotes, parens, braces). */
function splitArgs(text) {
  const args = [];
  let depth = 0;
  let inQuote = false;
  let cur = "";
  for (const ch of text) {
    if (ch === '"') inQuote = !inQuote;
    if (!inQuote) {
      if (ch === "(" || ch === "{") depth++;
      if (ch === ")" || ch === "}") depth--;
      if (ch === "," && depth === 0) {
        args.push(cur.trim());
        cur = "";
        continue;
      }
    }
    cur += ch;
  }
  if (cur.trim() !== "" || args.length) args.push(cur.trim());
  return args;
}

/** Every TB.<ID>( ... ) call in a formula, with its argument list. */
function tbCalls(formula) {
  const calls = [];
  const re = /TB\.([A-Z]+)\(/g;
  let m;
  while ((m = re.exec(formula))) {
    let depth = 1;
    let i = re.lastIndex;
    let inQuote = false;
    for (; i < formula.length && depth > 0; i++) {
      const ch = formula[i];
      if (ch === '"') inQuote = !inQuote;
      else if (!inQuote && ch === "(") depth++;
      else if (!inQuote && ch === ")") depth--;
    }
    calls.push({ id: m[1], args: splitArgs(formula.slice(re.lastIndex, i - 1)) });
  }
  return calls;
}

test("guide: every example calls a real TB function with no more arguments than it takes", () => {
  const formulas = guideFormulas();
  assert.ok(formulas.length >= 10, "the guide has a real set of examples");
  for (const { sheet, formula } of formulas) {
    const calls = tbCalls(formula);
    assert.ok(calls.length >= 1, `${sheet}: ${formula} calls a TB function`);
    for (const { id, args } of calls) {
      assert.ok(PARAMS[id], `${sheet}: TB.${id} exists in functions.json`);
      assert.ok(args.length <= PARAMS[id].length, `${sheet}: TB.${id} takes ${PARAMS[id].length} arguments, got ${args.length}`);
      const required = PARAMS[id].filter((p) => !p.optional).length;
      assert.ok(args.filter((a, i) => i < required && a !== "").length >= required, `${sheet}: TB.${id} gets its required arguments`);
    }
  }
});

test("guide: examples refer only to names the Setup sheet defines", () => {
  const defined = new Set();
  for (const sheet of GUIDE_SHEETS) {
    for (const b of sheet.blocks) if (b.input && b.input.name) defined.add(b.input.name);
  }
  assert.deepEqual([...defined].sort(), [...GUIDE_NAMES].sort(), "GUIDE_NAMES matches the inputs");

  for (const { sheet, formula } of guideFormulas()) {
    const withoutStrings = formula.replace(/"[^"]*"/g, '""');
    // Whole identifiers not followed by "(" -- i.e. names, not function calls.
    for (const ident of withoutStrings.match(/(?<![A-Za-z0-9_.])[A-Za-z_][A-Za-z0-9_.]*(?![A-Za-z0-9_.(])/g) ?? []) {
      if (/^[A-Z]{1,3}\d+$/.test(ident) || /^(TRUE|FALSE)$/.test(ident)) continue; // a cell address or boolean
      assert.ok(defined.has(ident), `${sheet}: "${ident}" in ${formula} is a defined name`);
    }
    for (const fn of withoutStrings.match(/\b([A-Z][A-Z0-9.]*)\(/g) ?? []) {
      const name = fn.slice(0, -1);
      assert.ok(name.startsWith("TB.") || EXCEL_FUNCTIONS.has(name), `${sheet}: ${name} is allowed`);
    }
  }
});

test("guide: sheet names are valid and unique; contents links point at real sheets", () => {
  const names = GUIDE_SHEETS.map((s) => s.name);
  assert.equal(new Set(names).size, names.length);
  for (const n of names) {
    assert.ok(n.length <= 31, `${n} fits Excel's 31-character limit`);
    assert.doesNotMatch(n, /[\\/?*[\]:]/, `${n} has no characters Excel forbids`);
  }
  for (const sheet of GUIDE_SHEETS) {
    for (const b of sheet.blocks) {
      if (b.table && b.table.links !== undefined) {
        for (const row of b.table.rows) assert.ok(names.includes(row[b.table.links]), `link to "${row[b.table.links]}"`);
      }
    }
  }
});

test("guide: live results sit right of the text column and do not share a column", () => {
  for (const sheet of GUIDE_SHEETS) {
    const cols = [];
    for (const b of sheet.blocks) {
      if (!b.example) continue;
      const m = /^([A-Z]+)(\d+)$/.exec(b.example.at);
      assert.ok(m, `${sheet.name}: "${b.example.at}" is a cell address`);
      assert.ok(m[1] >= "D" || m[1].length > 1, `${sheet.name}: ${b.example.at} is right of the text column`);
      assert.ok(Number(m[2]) >= 4, `${sheet.name}: ${b.example.at} leaves room for its label and formula text above`);
      cols.push({ col: m[1], row: Number(m[2]) });
    }
    // Two examples in one column must be far enough apart that neither is a spilling table.
    const byCol = new Map();
    for (const c of cols) byCol.set(c.col, [...(byCol.get(c.col) ?? []), c.row]);
    for (const [col, rows] of byCol) {
      if (rows.length > 1) {
        const scalars = sheet.blocks.filter((b) => b.example && b.example.at.startsWith(col)).every((b) => /TB\.VALUE\(/.test(b.example.formula));
        assert.ok(scalars, `${sheet.name}: only single-value examples may stack in column ${col}`);
      }
    }
  }
});

test("guide: docs/user-guide.md is generated from the current guide content", async () => {
  const { renderUserGuide, USER_GUIDE } = await import("../tools/make-docs.js");
  assert.equal(fs.readFileSync(USER_GUIDE, "utf8"), renderUserGuide(), "run: node tools/make-docs.js");
});
