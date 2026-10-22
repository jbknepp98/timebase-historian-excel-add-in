// Pre-flight checks that need no Windows machine and no Excel install.
//
// Catches the two mistakes that are otherwise invisible until you actually
// load the add-in: a custom function registered in code but not described in
// functions.json (Excel silently never sees it), and a manifest that points at
// a file that is not there.
//
//   node tools/check.js

import { readFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const rel = (...p) => path.join(ROOT, ...p);

const problems = [];
const notes = [];

// --- functions.js <-> functions.json ---------------------------------------

const functionsJs = readFileSync(rel("src/functions/functions.js"), "utf8");
const functionsJsonRaw = readFileSync(rel("src/functions/functions.json"), "utf8");

let functionsJson;
try {
  functionsJson = JSON.parse(functionsJsonRaw);
} catch (err) {
  problems.push(`functions.json is not valid JSON: ${err.message}`);
  functionsJson = { functions: [] };
}

const associatedIds = [...functionsJs.matchAll(/CustomFunctions\.associate\(\s*"([^"]+)"/g)].map((m) => m[1]);
const metadataIds = functionsJson.functions.map((f) => f.id);

notes.push(`associated in functions.js: ${associatedIds.join(", ") || "(none)"}`);
notes.push(`described in functions.json: ${metadataIds.join(", ") || "(none)"}`);

for (const id of associatedIds) {
  if (!metadataIds.includes(id)) {
    problems.push(`CustomFunctions.associate("${id}", ...) has no matching entry in functions.json`);
  }
}
for (const id of metadataIds) {
  if (!associatedIds.includes(id)) {
    problems.push(`functions.json describes "${id}" but functions.js never associates it`);
  }
}

// Every JSDoc @customfunction tag should also have a matching id, so a
// forgotten CustomFunctions.associate() call is caught even though the
// function itself exists.
const jsdocNames = [...functionsJs.matchAll(/@customfunction\s+(\S+)/g)].map((m) => m[1]);
for (const name of jsdocNames) {
  if (!associatedIds.includes(name)) {
    problems.push(`@customfunction ${name} is documented but never passed to CustomFunctions.associate()`);
  }
}

// Required parameters must precede optional ones; Excel's calling convention
// otherwise cannot tell which trailing arguments were omitted.
for (const fn of functionsJson.functions) {
  let seenOptional = false;
  for (const p of fn.parameters ?? []) {
    if (p.optional) seenOptional = true;
    else if (seenOptional) {
      problems.push(`${fn.id}: required parameter "${p.name}" follows an optional one`);
    }
  }
}

// A parameter that takes a list of tags must be declared as a matrix. Declared
// scalar, Excel "lifts" the call over an array argument -- one call per tag --
// and the per-tag tables cannot nest, so the cell shows #CALC! and the
// function is never even invoked.
if (functionsJson) {
  for (const f of functionsJson.functions ?? []) {
    for (const p of f.parameters ?? []) {
      if (p.name === "tags" && p.dimensionality !== "matrix") {
        problems.push(`functions.json: ${f.id}'s "tags" parameter must have "dimensionality": "matrix", or an array of tags gives #CALC!`);
      }
    }
  }
}

// --- manifest.xml ------------------------------------------------------------

const manifestPath = rel("manifest.xml");
if (!existsSync(manifestPath)) {
  problems.push("manifest.xml is missing");
} else {
  const original = readFileSync(manifestPath, "utf8");

  // Blank out comments before scanning for tags, preserving line numbers and
  // string length (spaces, not deletion), so a human sentence inside a
  // comment -- "immediately after <Host>" -- can never be mistaken for a real
  // tag, and every later index/line count still lines up with the real file.
  const manifest = original.replace(/<!--[\s\S]*?-->/g, (m) => m.replace(/[^\n]/g, " "));

  // XML forbids "--" inside a comment. Excel's manifest parser rejects such a
  // file outright, and the add-in then never appears, with no error shown.
  for (const c of original.matchAll(/<!--([\s\S]*?)-->/g)) {
    if (c[1].includes("--")) {
      const at = original.slice(0, c.index).split("\n").length;
      problems.push(`manifest.xml: the comment starting at line ${at} contains "--", which makes the file invalid XML`);
    }
  }

  // Cheap well-formedness check: every opening tag has a matching close, using
  // a stack. Good enough to catch a stray unclosed element without pulling in
  // an XML parser dependency.
  const stack = [];
  const tagRe = /<(\/?)([A-Za-z0-9_:.-]+)([^>]*?)(\/?)>/g;
  let m;
  let line = 1;
  let lastIndex = 0;
  while ((m = tagRe.exec(manifest))) {
    line += manifest.slice(lastIndex, m.index).split("\n").length - 1;
    lastIndex = m.index;
    const [, closing, name, , selfClosing] = m;
    if (manifest[m.index + 1] === "?") continue; // <?xml ...?>
    if (selfClosing) continue;
    if (closing) {
      const top = stack.pop();
      if (!top || top.name !== name) {
        problems.push(`manifest.xml: <${name}> closed at line ~${line} does not match ${top ? `<${top.name}> opened at line ~${top.line}` : "anything open"}`);
      }
    } else {
      stack.push({ name, line });
    }
  }
  for (const left of stack) problems.push(`manifest.xml: <${left.name}> opened at line ~${left.line} is never closed`);

  // resid is a generic reference into <Resources>: it can point at a bt:Url, a
  // bt:Image, or a bt:String (Short- or LongStrings alike) by id. Collecting
  // only bt:Url would misreport every icon and every UI label as undeclared.
  const declaredIds = [...manifest.matchAll(/<bt:(?:Url|Image|String)\s+id="([^"]+)"/g)].map((x) => x[1]);
  const referencedIds = [...manifest.matchAll(/resid="([^"]+)"/g)].map((x) => x[1]);
  for (const id of new Set(referencedIds)) {
    if (!declaredIds.includes(id)) problems.push(`manifest.xml: resid="${id}" is never declared in <Resources>`);
  }

  // Custom functions run in their own runtime page, not a shared runtime: a
  // sideloaded shared-runtime add-in never executes its functions on Excel for
  // Mac 16.105 (formulas stay #BUSY!). Reintroducing <Runtime> would bring
  // that back silently, so it is refused here.
  if (/<Runtime\b/.test(manifest)) {
    problems.push("manifest.xml: declares a <Runtime> (shared runtime); custom functions must run in their own page -- see docs/decisions.md");
  }
  const cfPageResid = (manifest.match(/xsi:type="CustomFunctions">[\s\S]*?<Page>\s*<SourceLocation resid="([^"]+)"/) ?? [])[1];
  if (!cfPageResid) {
    problems.push("manifest.xml: the CustomFunctions extension point has no <Page>");
  } else if (cfPageResid === "Taskpane.Url") {
    problems.push("manifest.xml: the CustomFunctions <Page> is the task pane; it must be the functions page");
  }

  notes.push(`manifest resource ids: ${declaredIds.join(", ")}`);
  notes.push(`custom functions page resid: ${cfPageResid ?? "(none)"}`);
}

// --- icons referenced by the manifest actually exist -------------------------

if (existsSync(manifestPath)) {
  const manifest = readFileSync(manifestPath, "utf8");
  const iconFiles = [...manifest.matchAll(/<bt:Image\s+id="[^"]*"\s+DefaultValue="([^"]+)"/g)].map((x) => x[1]);
  for (const url of iconFiles) {
    const filename = url.split("/").pop();
    if (!existsSync(rel("assets", filename))) {
      problems.push(`manifest.xml references icon "${url}" but assets/${filename} does not exist`);
    }
  }
}

// --- report -------------------------------------------------------------------

for (const n of notes) console.log("note:  " + n);
if (problems.length === 0) {
  console.log("\nOK: no problems found.");
} else {
  console.log("");
  for (const p of problems) console.log("FAIL:  " + p);
  process.exitCode = 1;
}
