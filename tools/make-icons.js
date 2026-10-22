// Resizes assets/icon-source.png (the Timebase star) into the three PNG icons
// the manifest references: 16, 32 and 80 px.
//
//   node tools/make-icons.js
//
// Uses macOS's built-in `sips`, so there is no image dependency to install.
// On another platform, resize icon-source.png to the same three sizes with any
// image tool and save them as assets/icon-<size>.png.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ASSETS = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "assets");
const SOURCE = path.join(ASSETS, "icon-source.png");
const SIZES = [16, 32, 80];

if (!fs.existsSync(SOURCE)) {
  console.error(`Missing ${SOURCE}. Put the square source image there first.`);
  process.exit(1);
}
if (process.platform !== "darwin") {
  console.error(`This script uses macOS's sips. Resize ${SOURCE} to ${SIZES.join(", ")} px by hand on this platform.`);
  process.exit(1);
}

for (const size of SIZES) {
  const out = path.join(ASSETS, `icon-${size}.png`);
  execFileSync("sips", ["-z", String(size), String(size), SOURCE, "--out", out], { stdio: "ignore" });
  console.log(`wrote ${path.relative(process.cwd(), out)}`);
}
