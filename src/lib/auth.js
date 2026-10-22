// Pulse OAuth2 client-credentials tokens.
//
// Pulse issues tokens at POST {pulse}/auth/token as form-encoded
// client_credentials. The token's `aud` claim carries the client's own id first,
// then its configured audiences -- so for Excel to reach the Historian, the
// Pulse client used here must list the Historian's client id among its
// audiences. That single fact causes most first-run 401s, so it is named in the
// error text rather than left to the reader.
//
// Tokens live an hour by default and client_credentials grants carry no refresh
// token, so the only caching that matters is "reuse until nearly expired, then
// ask again".

import { AuthError, ConfigError, snippet, looksLikeHtml } from "./errors.js";

export const DEFAULT_PULSE_PORT = 4542;

/** Expand a Pulse address the same way the Historian address is expanded. */
export function normalizePulseUrl(input) {
  if (input === null || input === undefined) return "";
  const raw = String(input).trim();
  if (raw === "") return "";
  const withScheme = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
  let parsed;
  try {
    parsed = new URL(withScheme);
  } catch {
    throw new ConfigError(`"${input}" is not a usable Pulse address. Use something like https://pulse:4542.`);
  }
  const port = parsed.port || String(DEFAULT_PULSE_PORT);
  return `${parsed.protocol}//${parsed.hostname}:${port}`;
}

/**
 * Derive the Pulse address from the Historian address when it is not set
 * explicitly: same host, Pulse's default port. Correct in the common
 * single-server install and harmless to override.
 */
export function pulseUrlFromHistorian(historianBaseUrl) {
  if (!historianBaseUrl) return "";
  try {
    const u = new URL(historianBaseUrl);
    return `${u.protocol}//${u.hostname}:${DEFAULT_PULSE_PORT}`;
  } catch {
    return "";
  }
}

/**
 * A token provider with caching.
 *
 * `getConfig()` is called on each acquisition rather than captured once, so
 * editing the credentials in the task pane takes effect on the next call
 * without rebuilding the client.
 */
export function createTokenProvider({ getConfig, fetchImpl = null, nowFn = () => Date.now(), skewMs = 60000 }) {
  const doFetch = fetchImpl ?? (typeof fetch === "function" ? fetch.bind(globalThis) : null);
  if (!doFetch) throw new ConfigError("No fetch implementation is available in this runtime.");

  let cached = null; // { token, expiresAtMs, key }
  let inFlight = null;

  /** Cache key: any credential change must invalidate the cached token. */
  const keyOf = (c) => `${c.pulseUrl}|${c.clientId}|${c.clientSecret}|${c.scope ?? ""}`;

  async function acquire(config) {
    const body = new URLSearchParams();
    body.append("grant_type", "client_credentials");
    body.append("client_id", config.clientId);
    body.append("client_secret", config.clientSecret ?? "");
    if (config.scope) body.append("scope", config.scope);

    const url = `${config.pulseUrl}/auth/token`;

    let res;
    try {
      res = await doFetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          Accept: "application/json",
        },
        body: body.toString(),
        credentials: "omit",
      });
    } catch (err) {
      throw new AuthError(
        `Could not reach Pulse at ${config.pulseUrl}. Check the address and that its certificate is trusted by this machine.`,
        { url, cause: err },
      );
    }

    const contentType = res.headers && res.headers.get ? res.headers.get("content-type") : null;
    const text = await res.text();

    if (looksLikeHtml(contentType, text)) {
      throw new AuthError(
        `Pulse returned an HTML page rather than a token. Check that ${config.pulseUrl} is the Pulse address and port (4542 by default).`,
        { status: res.status, url, body: text },
      );
    }

    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      /* fall through to the generic message below */
    }

    if (!res.ok || !json || !json.access_token) {
      const detail = json
        ? snippet([json.error, json.error_description].filter(Boolean).join(": "))
        : snippet(text);
      throw new AuthError(
        `Pulse did not issue a token (HTTP ${res.status}).${detail ? ` It said: ${detail}` : ""} ` +
          `On a fresh install the usual cause is the client secret: Pulse generates a random secret per client and never prints it again, so rotate the secret in Pulse and paste that value here.`,
        { status: res.status, url, body: text },
      );
    }

    const lifetimeMs = Number(json.expires_in ?? 3600) * 1000;
    return {
      token: json.access_token,
      // Expire early by `skewMs` so a token is never presented in the second it
      // lapses; Historian-side clock validation has no tolerance to spare.
      expiresAtMs: nowFn() + Math.max(0, lifetimeMs - skewMs),
    };
  }

  return {
    async getToken() {
      const config = getConfig();
      if (!config || !config.authEnabled) return null;

      if (!config.pulseUrl) {
        throw new ConfigError(
          "Authentication is on but no Pulse address is set. Enter it in the task pane, or turn authentication off if this Historian does not require it.",
        );
      }
      if (!config.clientId) {
        throw new ConfigError(
          "Authentication is on but no Pulse client id is set. Enter the client id and secret in the task pane.",
        );
      }

      const key = keyOf(config);
      if (cached && cached.key === key && cached.expiresAtMs > nowFn()) return cached.token;

      // Collapse concurrent callers onto one request. Excel recalculates many
      // custom functions at once, and without this every cell would mint its
      // own token on first calculation.
      if (inFlight && inFlight.key === key) return inFlight.promise;

      const promise = acquire(config)
        .then((result) => {
          cached = { ...result, key };
          inFlight = null;
          return result.token;
        })
        .catch((err) => {
          inFlight = null;
          throw err;
        });

      inFlight = { key, promise };
      return promise;
    },

    /** Force the next call to re-acquire; used after editing credentials. */
    invalidate() {
      cached = null;
      inFlight = null;
    },
  };
}
