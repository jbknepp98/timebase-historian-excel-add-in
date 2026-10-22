// Install the built add-in into Excel for Mac.
//
//   node tools/sideload-mac.js            install dist/manifest.xml
//   node tools/sideload-mac.js --remove   uninstall
//
// Excel for Mac loads add-in manifests from its sandbox "wef" folder. Two files
// there with the same add-in <Id> leave Excel unable to resolve it ("This
// add-in is no longer available"), so any other copy of this add-in is removed
// first. Excel must be restarted to pick up a change; then open a workbook
// linked to the add-in (dist/Timebase Historian.xlsx, or the guide).

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const WEF = path.join(os.homedir(), "Library", "Containers", "com.microsoft.Excel", "Data", "Documents", "wef");
const BUILT = path.join(ROOT, "dist", "manifest.xml");

if (process.platform !== "darwin") {
  console.error("This installs into Excel for Mac. On Windows see docs/installation.md (Upload My Add-in, or a shared-folder catalog).");
  process.exit(1);
}
if (!fs.existsSync(BUILT)) {
  console.error("dist/manifest.xml not found. Build first: node tools/package.js https://localhost:3443");
  process.exit(1);
}

const idOf = (xml) => (xml.match(/<Id>([^<]+)<\/Id>/) || [])[1];
const manifest = fs.readFileSync(BUILT, "utf8");
const id = idOf(manifest);
const version = (manifest.match(/<Version>([^<]+)<\/Version>/) || [])[1];

fs.mkdirSync(WEF, { recursive: true });
let removed = 0;
for (const f of fs.readdirSync(WEF)) {
  const file = path.join(WEF, f);
  if (!fs.statSync(file).isFile()) continue;
  if (idOf(fs.readFileSync(file, "utf8")) === id) {
    fs.rmSync(file);
    removed++;
  }
}

if (process.argv.includes("--remove")) {
  console.log(removed ? `Removed the add-in (${removed} manifest file${removed === 1 ? "" : "s"}). Restart Excel.` : "The add-in was not installed.");
  process.exit(0);
}

const target = path.join(WEF, `${id}.manifest.xml`);
fs.writeFileSync(target, manifest);
console.log(`Installed Timebase Historian ${version} into Excel for Mac:`);
console.log(`  ${target}`);
if (removed) console.log(`  (replaced ${removed} earlier copy${removed === 1 ? "" : "ies"})`);
console.log("");
console.log("Next: quit Excel (Cmd-Q), make sure `node tools/serve.js` is running, then open");
console.log("  dist/Timebase Historian Guide.xlsx   (the getting-started guide), or");
console.log("  dist/Timebase Historian.xlsx         (a blank workbook linked to the add-in)");
