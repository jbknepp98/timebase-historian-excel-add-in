// Where the dev server's Pulse client credentials live, and how they are read
// and written.
//
// The file sits outside the repository on purpose -- it can never be committed
// or packaged -- and is written readable by its owner only (0600). It holds:
//
//   { "clientId": "Excel", "clientSecret": "...", "scope": "" }
//
// Override the location with TIMEBASE_EXCEL_CREDENTIALS.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";

export const CREDENTIALS_PATH =
  process.env.TIMEBASE_EXCEL_CREDENTIALS || path.join(os.homedir(), ".config", "timebase-excel", "credentials.json");

/** The credentials, or null when no file exists. Throws on a malformed file. */
export function loadCredentials(file = CREDENTIALS_PATH) {
  if (!fs.existsSync(file)) return null;

  const mode = fs.statSync(file).mode & 0o777;
  if (mode & 0o077) {
    throw new Error(`${file} is readable by other users (mode ${mode.toString(8)}). Run: chmod 600 "${file}"`);
  }

  let json;
  try {
    json = JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (err) {
    throw new Error(`${file} is not valid JSON: ${err.message}`);
  }
  if (!json || typeof json.clientId !== "string" || !json.clientId.trim()) {
    throw new Error(`${file} has no clientId.`);
  }
  return {
    clientId: json.clientId.trim(),
    clientSecret: typeof json.clientSecret === "string" ? json.clientSecret : "",
    scope: typeof json.scope === "string" ? json.scope.trim() : "",
  };
}

/** Write the credentials file with owner-only permissions. */
export function saveCredentials({ clientId, clientSecret, scope = "" }, file = CREDENTIALS_PATH) {
  fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
  const tmp = `${file}.tmp-${process.pid}`;
  fs.writeFileSync(tmp, JSON.stringify({ clientId, clientSecret, scope }, null, 2) + "\n", { mode: 0o600 });
  fs.renameSync(tmp, file);
  fs.chmodSync(file, 0o600);
  return file;
}
