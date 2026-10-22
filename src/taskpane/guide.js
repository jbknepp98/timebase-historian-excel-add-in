/* global Excel */
// Writes the getting-started guide (src/lib/guide-content.js) into the open
// workbook. The TB.* examples are entered as real formulas through Office.js,
// so Excel stores them exactly as if a person had typed them -- which is what
// makes them spill correctly when the shipped guide is opened later.

import { GUIDE_SHEETS, GUIDE_NAMES } from "../lib/guide-content.js";

const COLORS = {
  ink: "#1A1F26",
  muted: "#5B6570",
  accent: "#0B7FB3",
  headerFill: "#1F6F8B",
  inputFill: "#FFF4CC",
  codeFill: "#F5F7F9",
  border: "#D8DCE1",
};

/** Column letters for 0-based index (0 = A). */
function col(i) {
  let s = "";
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function parseAddress(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { col: c - 1, row: Number(m[2]) };
}

/** Sheets in this workbook that the guide would overwrite. */
export async function guideConflicts() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const existing = new Set(sheets.items.map((s) => s.name));
    return GUIDE_SHEETS.map((s) => s.name).filter((n) => existing.has(n));
  });
}

/**
 * Build the guide. With `replace`, existing guide sheets and Guide_* names are
 * removed first; otherwise a conflict throws. An empty default "Sheet1" is
 * removed so the guide opens on its Welcome sheet.
 */
export async function buildGuide({ replace = false } = {}) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const sheets = wb.worksheets;
    sheets.load("items/name");
    await context.sync();

    const existing = sheets.items.map((s) => s.name);
    const clashing = GUIDE_SHEETS.map((s) => s.name).filter((n) => existing.includes(n));
    if (clashing.length && !replace) {
      throw new Error(`These sheets already exist: ${clashing.join(", ")}.`);
    }

    // A workbook must keep one sheet, so park a temporary one while replacing.
    const parking = clashing.length ? sheets.add("_guide_building") : null;
    for (const name of clashing) sheets.getItem(name).delete();
    for (const n of GUIDE_NAMES) wb.names.getItemOrNullObject(n).delete();
    await context.sync();

    for (const spec of GUIDE_SHEETS) {
      const ws = sheets.add(spec.name);
      writeSheet(wb, ws, spec);
      await context.sync();
    }

    if (parking) parking.delete();
    const sheet1 = sheets.getItemOrNullObject("Sheet1");
    sheet1.load("isNullObject");
    await context.sync();
    if (!sheet1.isNullObject) {
      const used = sheet1.getUsedRangeOrNullObject(true);
      used.load("isNullObject");
      await context.sync();
      if (used.isNullObject) sheet1.delete();
    }
    sheets.getItem(GUIDE_SHEETS[0].name).activate();
    await context.sync();
  });
}

function writeSheet(wb, ws, spec) {
  ws.showGridlines = false;
  // Widths in points. C and D are wide only where this sheet uses them (table
  // columns, input values and their notes), so live results in F onward stay
  // close to the text they explain.
  const tableWidth = Math.max(0, ...spec.blocks.filter((b) => b.table).map((b) => b.table.header.length));
  const hasInputs = spec.blocks.some((b) => b.input);
  const hasExamples = spec.blocks.some((b) => b.example);
  ws.getRange("A:A").format.columnWidth = 12;
  ws.getRange("B:B").format.columnWidth = hasExamples ? 300 : 260;
  ws.getRange("C:C").format.columnWidth = tableWidth >= 2 ? 300 : hasInputs ? 130 : 12;
  ws.getRange("D:D").format.columnWidth = tableWidth >= 3 ? 320 : hasInputs ? 200 : 12;
  ws.getRange("E:E").format.columnWidth = 16;
  ws.getRange("A:Z").format.verticalAlignment = "Top";

  const examples = spec.blocks.filter((b) => b.example);
  for (const { example } of examples) {
    // Result columns: timestamps first, so they get room for mm/dd/yy hh:mm:ss.000.
    const { col: c } = parseAddress(example.at);
    ws.getRange(`${col(c)}:${col(c)}`).format.columnWidth = 140;
  }

  let row = 2;
  let first = true;
  const cell = (r, c = 1) => ws.getRange(`${col(c)}${r}`);

  for (const b of spec.blocks) {
    if (b.title) {
      const r = cell(row);
      r.values = [[b.title]];
      r.format.font.size = 20;
      r.format.font.bold = true;
      r.format.font.color = COLORS.ink;
      row += 1;
    } else if (b.subtitle) {
      const r = cell(row);
      r.values = [[b.subtitle]];
      r.format.font.size = 12;
      r.format.font.color = COLORS.muted;
      r.format.wrapText = true;
      row += 2;
    } else if (b.heading) {
      if (!first) row += 1;
      const r = cell(row);
      r.values = [[b.heading]];
      r.format.font.size = 14;
      r.format.font.bold = true;
      r.format.font.color = COLORS.accent;
      row += 1;
    } else if (b.text) {
      const r = cell(row);
      r.values = [[b.text]];
      r.format.wrapText = true;
      row += 2;
    } else if (b.bullets) {
      for (const text of b.bullets) {
        const r = cell(row);
        r.values = [[/^\d+\./.test(text) ? text : `•  ${text}`]];
        r.format.wrapText = true;
        row += 1;
      }
      row += 1;
    } else if (b.table) {
      const { header, rows, links } = b.table;
      const width = header.length;
      const all = ws.getRange(`B${row}:${col(width)}${row + rows.length}`);
      all.values = [header, ...rows];
      all.format.wrapText = true;
      for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
        const border = all.format.borders.getItem(edge);
        border.style = "Continuous";
        border.color = COLORS.border;
      }
      const head = ws.getRange(`B${row}:${col(width)}${row}`);
      head.format.fill.color = COLORS.headerFill;
      head.format.font.color = "#FFFFFF";
      head.format.font.bold = true;
      if (links !== undefined) {
        rows.forEach((r, i) => {
          const target = r[links];
          cell(row + 1 + i, 1 + links).hyperlink = {
            documentReference: `'${target.replace(/'/g, "''")}'!A1`,
            textToDisplay: target,
            screenTip: `Go to ${target}`,
          };
        });
      }
      row += rows.length + 2;
    } else if (b.input) {
      const { label, name, values, note } = b.input;
      const labelCell = cell(row);
      labelCell.values = [[label]];
      labelCell.format.font.bold = true;
      const range = ws.getRange(`C${row}:C${row + values.length - 1}`);
      range.formulas = values.map((v) => [v]);
      range.format.fill.color = COLORS.inputFill;
      for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal"]) {
        const border = range.format.borders.getItem(edge);
        border.style = "Continuous";
        border.color = COLORS.border;
      }
      if (note) {
        const n = cell(row, 3);
        n.values = [[note]];
        n.format.font.color = COLORS.muted;
        n.format.font.italic = true;
        n.format.wrapText = true;
      }
      if (name) wb.names.add(name, range, label);
      row += values.length + 1;
    } else if (b.example) {
      const { label, formula, at } = b.example;
      const { col: c, row: r } = parseAddress(at);
      const lab = ws.getRange(`${col(c)}${r - 2}`);
      lab.values = [[label]];
      lab.format.font.bold = true;
      lab.format.font.color = COLORS.accent;
      const shown = ws.getRange(`${col(c)}${r - 1}`);
      shown.formulas = [[`=FORMULATEXT(${at})`]];
      shown.format.font.name = "Menlo";
      shown.format.font.size = 10;
      shown.format.font.color = COLORS.muted;
      ws.getRange(at).formulas = [[formula]];
    }
    first = false;
  }

  ws.getRange(`B1:D${row}`).format.autofitRows();
}
