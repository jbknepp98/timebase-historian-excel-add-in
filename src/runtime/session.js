// One configured client per runtime.
//
// The task pane and the custom functions run in separate runtimes (see
// docs/decisions.md), so each has its own copy of this module. They share
// settings through OfficeRuntime.storage: the task pane saves, and every
// custom function calls loadSettings() first, so the next recalculation picks
// up a change. The client and its token cache are rebuilt only when the
// loaded settings actually differ.

import { createClient } from "../lib/historian.js";
import { createTokenProvider } from "../lib/auth.js";
import { createSettingsStore } from "../lib/settings.js";
import { ConfigError } from "../lib/errors.js";

const store = createSettingsStore();

let tokenProvider = null;
let client = null;
let clientKeyBuiltFrom = null;

function clientKey(settings) {
  return `${settings.historianUrl}|${settings.authEnabled}|${settings.pulseUrl}|${settings.clientId}|${settings.clientSecret}|${settings.scope}`;
}

/**
 * The current client, rebuilt only when the underlying settings actually
 * changed. Rebuilding on every call would throw away the auth token cache on
 * every single cell recalculation.
 */
export function getClient() {
  const settings = store.get();
  const key = clientKey(settings);

  if (!client || key !== clientKeyBuiltFrom) {
    if (!tokenProvider || key !== clientKeyBuiltFrom) {
      tokenProvider = createTokenProvider({ getConfig: () => store.get() });
    }
    client = createClient({
      baseUrl: settings.historianUrl,
      getToken: settings.authEnabled ? () => tokenProvider.getToken() : null,
    });
    clientKeyBuiltFrom = key;
  }
  return client;
}

/** Refresh settings from storage; call before getClient() in a fresh runtime. */
export async function loadSettings() {
  await store.load();
  return store.get();
}

export function getSettings() {
  return store.get();
}

export async function saveSettings(settings) {
  const saved = await store.save(settings);
  // Force the next getClient() to rebuild, and drop any cached token for the
  // credentials that just changed.
  clientKeyBuiltFrom = null;
  if (tokenProvider) tokenProvider.invalidate();
  return saved;
}

export function requireConfigured() {
  const settings = getSettings();
  if (!settings.historianUrl) {
    throw new ConfigError(
      "The Historian address is not set. Open the Timebase task pane and enter your Historian's address.",
    );
  }
  return settings;
}
