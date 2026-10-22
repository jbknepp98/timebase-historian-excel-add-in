import { test } from "node:test";
import assert from "node:assert/strict";

import {
  defaultSettings,
  parseSettings,
  serializeSettings,
  validateSettings,
  createSettingsStore,
  SETTINGS_VERSION,
} from "../src/lib/settings.js";
import { ConfigError } from "../src/lib/errors.js";

test("defaultSettings: authEnabled defaults true, matching a stock Historian", () => {
  assert.equal(defaultSettings().authEnabled, true);
});

test("parseSettings: missing or null input falls back to defaults", () => {
  assert.deepEqual(parseSettings(null), defaultSettings());
  assert.deepEqual(parseSettings(undefined), defaultSettings());
});

test("parseSettings: merges a partial or older-shaped object over defaults", () => {
  const s = parseSettings({ historianUrl: "https://h:4512" });
  assert.equal(s.historianUrl, "https://h:4512");
  assert.equal(s.maxRows, defaultSettings().maxRows, "untouched fields keep their default");
  assert.equal(s.version, SETTINGS_VERSION, "version is stamped to current, not carried from disk");
});

test("parseSettings: corrupt JSON string falls back to defaults rather than throwing", () => {
  assert.deepEqual(parseSettings("{not json"), defaultSettings());
});

test("serializeSettings then parseSettings round-trips", () => {
  const s = { ...defaultSettings(), historianUrl: "https://h:4512", clientId: "Excel" };
  assert.deepEqual(parseSettings(serializeSettings(s)), s);
});

test("validateSettings: requires a Historian address", () => {
  assert.throws(() => validateSettings({ historianUrl: "" }), ConfigError);
});

test("validateSettings: normalises the Historian address", () => {
  const v = validateSettings({ historianUrl: "historian", authEnabled: false });
  assert.equal(v.historianUrl, "https://historian:4512");
});

test("validateSettings: auth on requires a client id", () => {
  assert.throws(
    () => validateSettings({ historianUrl: "historian", authEnabled: true, clientId: "" }),
    ConfigError,
  );
});

test("validateSettings: derives the Pulse url from the Historian host when not set", () => {
  const v = validateSettings({ historianUrl: "historian", authEnabled: true, clientId: "Excel" });
  assert.equal(v.pulseUrl, "https://historian:4542");
});

test("validateSettings: an explicit Pulse url is normalised and kept", () => {
  const v = validateSettings({
    historianUrl: "historian",
    authEnabled: true,
    clientId: "Excel",
    pulseUrl: "otherpulse",
  });
  assert.equal(v.pulseUrl, "https://otherpulse:4542");
});

test("validateSettings: auth off clears client id, secret and pulse url", () => {
  const v = validateSettings({
    historianUrl: "historian",
    authEnabled: false,
    clientId: "leftover",
    clientSecret: "leftover",
    pulseUrl: "leftover",
  });
  assert.equal(v.clientId, "");
  assert.equal(v.clientSecret, "");
  assert.equal(v.pulseUrl, "");
});

test("validateSettings: rejects a non-positive max row limit", () => {
  assert.throws(
    () => validateSettings({ historianUrl: "historian", authEnabled: false, maxRows: 0 }),
    ConfigError,
  );
  assert.throws(
    () => validateSettings({ historianUrl: "historian", authEnabled: false, maxRows: "many" }),
    ConfigError,
  );
});

test("validateSettings: trims whitespace around client id, secret and scope", () => {
  const v = validateSettings({
    historianUrl: "historian",
    authEnabled: true,
    clientId: "  Excel  ",
    scope: "  openid  ",
  });
  assert.equal(v.clientId, "Excel");
  assert.equal(v.scope, "openid");
});

// A minimal fake of OfficeRuntime.storage: async getItem/setItem.
function fakeStorage() {
  const map = new Map();
  return {
    map,
    getItem: async (key) => (map.has(key) ? map.get(key) : null),
    setItem: async (key, value) => {
      map.set(key, value);
    },
  };
}

test("createSettingsStore: without storage, behaves as in-memory storage", async () => {
  const store = createSettingsStore(null);
  assert.deepEqual(store.get(), defaultSettings());
  const saved = await store.save({ historianUrl: "historian", authEnabled: false });
  assert.equal(store.get().historianUrl, "https://historian:4512");
  assert.equal(saved.historianUrl, "https://historian:4512");
});

test("createSettingsStore: a second store (the other runtime) sees a save after load()", async () => {
  const storage = fakeStorage();
  const taskpane = createSettingsStore(storage);
  const functions = createSettingsStore(storage);
  await taskpane.save({ historianUrl: "historian", authEnabled: false });
  assert.equal(functions.get().historianUrl, "", "not visible until loaded");
  await functions.load();
  assert.equal(functions.get().historianUrl, "https://historian:4512");
});

test("createSettingsStore: save() rejects invalid input before touching storage", async () => {
  const storage = fakeStorage();
  const store = createSettingsStore(storage);
  await assert.rejects(() => store.save({ historianUrl: "" }), ConfigError);
  assert.equal(storage.map.size, 0);
});

test("createSettingsStore: a failed setItem surfaces as a ConfigError", async () => {
  const storage = fakeStorage();
  storage.setItem = async () => {
    throw new Error("disk full");
  };
  const store = createSettingsStore(storage);
  await assert.rejects(() => store.save({ historianUrl: "historian", authEnabled: false }), /disk full/);
});
