// Timebase Historian REST client.
//
// Built from the Historian OpenAPI document (v1). Read paths only -- this
// add-in never writes tags or data, and never deletes a dataset.
//
// The API's JSON is heavily abbreviated (`n` name, `d` description, `t` is both
// "tag" and "timestamp" and "data type" depending on where it appears), so the
// abbreviations are decoded once here and nothing downstream has to know them.

import { describeHttpFailure, looksLikeHtml, HistorianError, ConfigError } from "./errors.js";

export const DEFAULT_HTTP_PORT = 4511;
export const DEFAULT_HTTPS_PORT = 4512;

/**
 * Normalise a server address.
 *
 * An empty address means "same origin", which is the right configuration when
 * the add-in is served by the Historian itself: relative URLs then need no CORS
 * at all. Anything else is expanded to scheme://host:port with no path, so a
 * pasted `/api/help` or trailing slash cannot double up.
 */
export function normalizeBaseUrl(input) {
  if (input === null || input === undefined) return "";
  const raw = String(input).trim();
  if (raw === "") return "";

  const hasScheme = /^https?:\/\//i.test(raw);
  const withScheme = hasScheme ? raw : `https://${raw}`;

  let parsed;
  try {
    parsed = new URL(withScheme);
  } catch {
    throw new ConfigError(`"${input}" is not a usable Historian address. Use something like https://historian:4512.`);
  }
  if (!parsed.hostname) {
    throw new ConfigError(`"${input}" has no host name. Use something like https://historian:4512.`);
  }

  const isHttps = parsed.protocol === "https:";
  const port = parsed.port || String(isHttps ? DEFAULT_HTTPS_PORT : DEFAULT_HTTP_PORT);
  return `${parsed.protocol}//${parsed.hostname}:${port}`;
}

/** Decode a Tag object from the API's abbreviated form. */
export function decodeTag(t) {
  if (!t || typeof t !== "object") return null;
  const last = t.l ?? null;
  return {
    name: t.n ?? "",
    description: t.d ?? "",
    format: t.f ?? "",
    // `u` is a unit dictionary whose key signs and count carry meaning; see format.js.
    units: t.u ?? null,
    fields: t.fl ?? null,
    dataType: t.t ?? "",
    last: last
      ? {
          time: last.t ?? null,
          value: last.v ?? null,
          quality: last.q ?? null,
          display: last.vs ?? null,
        }
      : null,
  };
}

/**
 * The tags route's body, as a flat list.
 *
 * Historian 2.0 groups tags by category -- `{"User": [...], "System": [...]}`
 * -- where the v1 contract this add-in was first built from showed a bare
 * array. Both are accepted. Groups keep their response order, so user tags
 * (listed first) come before the Historian's own system tags.
 */
export function tagListFromResponse(json) {
  if (Array.isArray(json)) return json;
  if (!json || typeof json !== "object") return [];
  return Object.values(json).flatMap((group) => (Array.isArray(group) ? group : []));
}

/** Decode a Dataset object from the API's abbreviated form. */
export function decodeDataset(d) {
  if (!d || typeof d !== "object") return null;
  return {
    name: d.n ?? "",
    purgeAgeDays: d.pa ?? null,
    purgeSizeGb: d.ps ?? null,
    lateDataToleranceMs: d.ldt ?? null,
    lateDataAgeDays: d.lda ?? null,
    sizeGb: d.s ?? null,
    tagCount: d.c ?? null,
    dataAgeDays: d.a ?? null,
  };
}

/**
 * Build the query string for a data read.
 *
 * `tagname` repeats once per tag, which is how the API takes a batch. Time
 * bounds arrive already classified by time.js as start/relativeStart etc., so
 * this function does not need to know Excel from ISO from relative.
 */
export function buildDataQuery({ tags, startParam, endParam }) {
  const q = new URLSearchParams();
  for (const t of tags) q.append("tagname", t);
  if (startParam) q.append(startParam.param, startParam.value);
  if (endParam) q.append(endParam.param, endParam.value);
  return q.toString();
}

export function createClient({ baseUrl = "", getToken = null, fetchImpl = null, timeoutMs = 120000 } = {}) {
  const base = normalizeBaseUrl(baseUrl);
  const doFetch = fetchImpl ?? (typeof fetch === "function" ? fetch.bind(globalThis) : null);
  if (!doFetch) throw new ConfigError("No fetch implementation is available in this runtime.");

  async function request(path, query) {
    const url = `${base}${path}${query ? `?${query}` : ""}`;

    const headers = { Accept: "application/json" };
    if (getToken) {
      const token = await getToken();
      if (token) headers.Authorization = `Bearer ${token}`;
    }

    // AbortController rather than a bare promise race, so an abandoned request
    // is actually cancelled instead of left running behind Excel's back.
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    const timer =
      controller && timeoutMs
        ? setTimeout(() => controller.abort(), timeoutMs)
        : null;

    let res;
    try {
      res = await doFetch(url, {
        method: "GET",
        headers,
        signal: controller ? controller.signal : undefined,
        // Never send cookies: this is a bearer-token API, and an ambient cookie
        // would make a CORS request non-simple for no benefit.
        credentials: "omit",
        redirect: "manual",
      });
    } catch (err) {
      if (err && (err.name === "AbortError" || err.name === "TimeoutError")) {
        throw new HistorianError(
          `The Historian did not respond within ${Math.round(timeoutMs / 1000)} seconds. The API has no server-side aggregation, so a wide time range on a fast tag can take a long time -- narrow the range or pass an interval.`,
          { url },
        );
      }
      throw new HistorianError(
        `Could not reach the Historian at ${base || "this add-in's own origin"}. ${err && err.message ? err.message : ""} Check the address, that the service is running, and that its certificate is trusted.`,
        { url, cause: err },
      );
    } finally {
      if (timer) clearTimeout(timer);
    }

    // redirect: "manual" surfaces a 3xx instead of quietly following it to a
    // sign-in page and handing us HTML where JSON was expected.
    if (res.type === "opaqueredirect" || (res.status >= 300 && res.status < 400)) {
      throw new HistorianError(
        `The Historian redirected the request instead of answering it. This usually means authentication is required and the token was missing or not accepted.`,
        { status: res.status, url },
      );
    }

    const contentType = res.headers && res.headers.get ? res.headers.get("content-type") : null;
    const text = await res.text();

    // Checked before the ok/status branch on purpose: a misrouted request can
    // come back as a 200 carrying an HTML shell (a reverse proxy's default
    // page, a captive portal, a SPA fallback) just as easily as a non-2xx one,
    // and that failure mode is far more confusing without this check, because
    // it would otherwise be reported as "not valid JSON" instead of naming the
    // actual cause.
    if (looksLikeHtml(contentType, text)) {
      throw describeHttpFailure({ status: res.status, url, contentType, body: text });
    }

    if (!res.ok) {
      throw describeHttpFailure({ status: res.status, url, contentType, body: text });
    }

    if (text.trim() === "") return null;

    try {
      return JSON.parse(text);
    } catch (err) {
      throw new HistorianError(
        `The Historian returned a body that is not valid JSON. First bytes: ${text.slice(0, 120)}`,
        { status: res.status, url, body: text, cause: err },
      );
    }
  }

  return {
    baseUrl: base,

    async listDatasets() {
      const json = await request("/api/datasets");
      return (Array.isArray(json) ? json : []).map(decodeDataset).filter(Boolean);
    },

    async getDataset(dataset) {
      const json = await request(`/api/datasets/${encodeURIComponent(dataset)}`);
      return decodeDataset(json);
    },

    async datasetExists(dataset) {
      const json = await request(`/api/datasets/${encodeURIComponent(dataset)}/exists`);
      return json === true;
    },

    /**
     * Tag search.
     *
     * `query` takes the Historian's filter language -- free text, `name:`,
     * `description:`, `datatype:`, `type:`, `format:`, AND/OR/NOT, wildcards.
     * `contains` is the older simple substring filter. They are separate
     * parameters on the API, so both are exposed rather than guessed between.
     */
    async listTags(dataset, { query = null, contains = null } = {}) {
      const q = new URLSearchParams();
      if (query) q.append("query", query);
      if (contains) q.append("contains", contains);
      const json = await request(
        `/api/datasets/${encodeURIComponent(dataset)}/tags`,
        q.toString(),
      );
      return tagListFromResponse(json).map(decodeTag).filter(Boolean);
    },

    /**
     * Read data for one or more tags.
     *
     * Returns the raw TVQsResponse rather than a decoded shape, because
     * align.js consumes `s`, `e` and `tl` directly and decoding twice would
     * cost a full copy of what can be hundreds of thousands of points.
     */
    async readData(dataset, tags, startParam, endParam) {
      if (!Array.isArray(tags) || tags.length === 0) {
        throw new ConfigError("At least one tag name is required.");
      }
      const query = buildDataQuery({ tags, startParam, endParam });
      const json = await request(`/api/datasets/${encodeURIComponent(dataset)}/data`, query);
      // A single-tag read still comes back in the tl envelope on this route.
      // Normalise a null body to an empty response so callers need no guard.
      return json ?? { s: null, e: null, tl: [] };
    },
  };
}
