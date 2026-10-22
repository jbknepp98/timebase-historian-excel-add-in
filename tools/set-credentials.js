// Store the Pulse client credentials the dev server uses on Excel's behalf.
//
//   node tools/set-credentials.js
//
// Prompts for the client id and secret (the secret is not echoed) and writes
// them to ~/.config/timebase-excel/credentials.json, readable by you only.
// tools/serve.js then fetches Pulse tokens itself and attaches them to every
// Historian request, so no workbook, task pane or Excel setting ever holds the
// secret. Restart tools/serve.js afterwards.
//
//   node tools/set-credentials.js --clear    remove the file

import fs from "node:fs";
import readline from "node:readline";
import { CREDENTIALS_PATH, loadCredentials, saveCredentials } from "./credentials.js";

if (process.argv.includes("--clear")) {
  if (fs.existsSync(CREDENTIALS_PATH)) fs.rmSync(CREDENTIALS_PATH);
  console.log(`Removed ${CREDENTIALS_PATH}`);
  process.exit(0);
}

if (!process.stdin.isTTY) {
  console.error("Run this in a terminal: it reads the secret without echoing it.");
  process.exit(1);
}

function ask(question, { hidden = false, fallback = "" } = {}) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true });
    if (hidden) {
      // Print the prompt, then swallow everything readline would echo.
      rl._writeToOutput = (s) => {
        if (s.startsWith(question)) rl.output.write(question);
      };
    }
    rl.question(question, (answer) => {
      rl.close();
      if (hidden) process.stdout.write("\n");
      resolve(answer.trim() === "" ? fallback : answer.trim());
    });
  });
}

let existing = null;
try {
  existing = loadCredentials();
} catch {
  /* a broken file is simply replaced */
}

const clientId = await ask(`Pulse client id${existing ? ` [${existing.clientId}]` : " [Excel]"}: `, {
  fallback: existing ? existing.clientId : "Excel",
});
const clientSecret = await ask(
  `Pulse client secret${existing && existing.clientSecret ? " [keep current]" : ""}: `,
  { hidden: true, fallback: existing ? existing.clientSecret : "" },
);
if (!clientSecret) {
  console.error("No secret entered; nothing written.");
  process.exit(1);
}

const file = saveCredentials({ clientId, clientSecret, scope: existing ? existing.scope : "" });
console.log(`Saved client "${clientId}" to ${file} (readable by you only).`);
console.log("Restart tools/serve.js to use it.");
