// Connection settings, persisted and shared between the task pane and custom
// functions.
//
// Stored in OfficeRuntime.storage (per add-in, per machine), which both the
// task pane and the separate custom-functions runtime can read -- see
// createSettingsStore below and docs/security.md. This module only decides
// *what* is stored.

import { ConfigError } from "./errors.js";
import { normalizeBaseUrl } from "./historian.js";
import { normalizePulseUrl, pulseUrlFromHistorian } from "./auth.js";

export const SETTINGS_KEY = "timebase.historian.connection";
export const SETTINGS_VERSION = 1;

export function defaultSettings() {
  return {
    version: SETTINGS_VERSION,
    historianUrl: "",
    authEnabled: true,
    pulseUrl: "",
    clientId: "",
    clientSecret: "",
    scope: "",
    defaultDataset: "",
    maxRows: 100000,
    // Points with no value (quality notes): "flag" keeps the row, carries the
    // value and shows it red; "drop" writes novalue "drop" into new formulas.
    noValueRows: "flag",
    // Companion columns after each tag in new formulas (columns.js kinds).
    extraColumns: [],
  };
}

/** Merge stored JSON with defaults, tolerating an older or missing shape. */
export function parseSettings(raw) {
  const base = defaultSettings();
  if (!raw) return base;
  let obj;
  if (typeof raw === "string") {
    try {
      obj = JSON.parse(raw);
    } catch {
      return base;
    }
  } else if (typeof raw === "object") {
    obj = raw;
  } else {
    return base;
  }
  return { ...base, ...obj, version: SETTINGS_VERSION };
}

export function serializeSettings(settings) {
  return JSON.stringify({ ...defaultSettings(), ...settings, version: SETTINGS_VERSION });
}

/**
 * Validate and normalise settings coming out of the task pane form, before
 * they are saved. Throws ConfigError with a message meant for direct display.
 */
export function validateSettings(input) {
  const s = { ...defaultSettings(), ...input };

  const historianUrl = normalizeBaseUrl(s.historianUrl);
  if (!historianUrl) {
    throw new ConfigError("Enter the Historian server address, for example https://historian:4512.");
  }

  let pulseUrl = "";
  if (s.authEnabled) {
    pulseUrl = s.pulseUrl ? normalizePulseUrl(s.pulseUrl) : pulseUrlFromHistorian(historianUrl);
    if (!s.clientId || !s.clientId.trim()) {
      throw new ConfigError("Authentication is on: enter the Pulse client id.");
    }
  }

  const maxRows = Number(s.maxRows);
  if (!isFinite(maxRows) || maxRows <= 0) {
    throw new ConfigError("Max rows must be a positive number.");
  }

  return {
    ...s,
    historianUrl,
    pulseUrl,
    clientId: s.authEnabled ? s.clientId.trim() : "",
    clientSecret: s.authEnabled ? s.clientSecret ?? "" : "",
    scope: s.authEnabled ? (s.scope ?? "").trim() : "",
    maxRows: Math.trunc(maxRows),
    noValueRows: s.noValueRows === "drop" ? "drop" : "flag",
    extraColumns: Array.isArray(s.extraColumns)
      ? ["quality", "status", "samples", "good%"].filter((k) => s.extraColumns.includes(k))
      : [],
  };
}

/**
 * Settings persisted in OfficeRuntime.storage.
 *
 * Why not Office.context.document.settings: custom functions run in their own
 * runtime (see docs/decisions.md), which has no document object, so a value
 * saved into the workbook from the task pane would be invisible to TB.*
 * formulas. OfficeRuntime.storage is the store Microsoft provides for exactly
 * this: one key/value store per add-in on this machine, readable from both the
 * task pane and the custom-functions runtime. A side effect is that the Pulse
 * client secret no longer travels inside the workbook file.
 *
 * Storage is asynchronous, so the store keeps a cached copy: `get()` is
 * synchronous and returns the last value loaded or saved, and `load()` refreshes
 * it. Callers that may be first in a fresh runtime (every custom function) await
 * `load()` before `get()`.
 *
 * `storage` defaults to the global OfficeRuntime.storage, resolved lazily so
 * this module can be imported before Office has finished initialising. Without
 * one (unit tests, a plain browser preview) the store is in-memory only.
 */
export function createSettingsStore(storage = undefined) {
  const resolve = () => {
    if (storage !== undefined) return storage;
    return typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage ? OfficeRuntime.storage : null;
  };

  let cached = defaultSettings();

  return {
    get() {
      return cached;
    },

    async load() {
      const backend = resolve();
      if (!backend) return cached;
      try {
        cached = parseSettings(await backend.getItem(SETTINGS_KEY));
      } catch (err) {
        throw new ConfigError(`Could not read the saved connection settings: ${err && err.message ? err.message : err}`);
      }
      return cached;
    },

    async save(settings) {
      const validated = validateSettings(settings);
      const backend = resolve();
      if (backend) {
        try {
          await backend.setItem(SETTINGS_KEY, serializeSettings(validated));
        } catch (err) {
          throw new ConfigError(`Could not save settings: ${err && err.message ? err.message : err}`);
        }
      }
      cached = validated;
      return validated;
    },
  };
}
