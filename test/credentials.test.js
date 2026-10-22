import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import { loadCredentials, saveCredentials } from "../tools/credentials.js";

function tempFile() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "tb-cred-"));
  return path.join(dir, "nested", "credentials.json");
}

test("credentials: saved owner-only and read back", () => {
  const file = tempFile();
  saveCredentials({ clientId: "Excel", clientSecret: "s3cret" }, file);
  assert.equal(fs.statSync(file).mode & 0o777, 0o600);
  assert.deepEqual(loadCredentials(file), { clientId: "Excel", clientSecret: "s3cret", scope: "" });
});

test("credentials: a missing file means none, not an error", () => {
  assert.equal(loadCredentials(tempFile()), null);
});

test("credentials: a file other users can read is refused", () => {
  const file = tempFile();
  saveCredentials({ clientId: "Excel", clientSecret: "s3cret" }, file);
  fs.chmodSync(file, 0o644);
  assert.throws(() => loadCredentials(file), /readable by other users/);
});

test("credentials: malformed or id-less files are refused", () => {
  const file = tempFile();
  saveCredentials({ clientId: "Excel", clientSecret: "x" }, file);
  fs.writeFileSync(file, "{not json", { mode: 0o600 });
  assert.throws(() => loadCredentials(file), /not valid JSON/);
  fs.writeFileSync(file, JSON.stringify({ clientSecret: "x" }), { mode: 0o600 });
  assert.throws(() => loadCredentials(file), /no clientId/);
});
