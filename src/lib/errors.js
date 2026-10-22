// Error shaping.
//
// Excel is a hostile place to surface an error: a custom function can only
// return a single cell value, and #VALUE! with a one-line tooltip is all the
// user sees. So every message here has to name the cause and the fix in one
// sentence, without a stack trace.

export class HistorianError extends Error {
  constructor(message, { status = null, url = null, body = null, cause = null } = {}) {
    super(message);
    this.name = "HistorianError";
    this.status = status;
    this.url = url;
    this.body = body;
    if (cause) this.cause = cause;
  }
}

export class ConfigError extends Error {
  constructor(message) {
    super(message);
    this.name = "ConfigError";
  }
}

export class AuthError extends HistorianError {
  constructor(message, details) {
    super(message, details);
    this.name = "AuthError";
  }
}

/** First few hundred characters of a response body, for diagnostics. */
export function snippet(text, limit = 300) {
  if (typeof text !== "string") return "";
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length > limit ? `${flat.slice(0, limit)}...` : flat;
}

/** Does this body look like an HTML page rather than an API response? */
export function looksLikeHtml(contentType, body) {
  if (contentType && /text\/html/i.test(contentType)) return true;
  return typeof body === "string" && body.trimStart().startsWith("<");
}

/**
 * Turn a non-2xx response into an error whose message tells the user what to do.
 *
 * The Historian's OpenAPI document declares 400/401/404 but no error body
 * schema, so these messages are built from the status and whatever the body
 * happens to contain rather than from a documented envelope.
 */
export function describeHttpFailure({ status, url, contentType, body, clientId }) {
  const tail = snippet(body);

  if (looksLikeHtml(contentType, body)) {
    return new HistorianError(
      `The Historian returned an HTML page instead of data. This usually means the URL is wrong, or the request was redirected to a sign-in page. Check the server address: it should be the host and port only, such as https://historian:4512, with no /api suffix.`,
      { status, url, body },
    );
  }

  switch (status) {
    case 400:
      return new HistorianError(
        `The Historian rejected the request as invalid (400).${tail ? ` It said: ${tail}` : ""} Check the dataset name and the time range.`,
        { status, url, body },
      );
    case 401:
      return new AuthError(
        `The Historian rejected the access token (401). The most common cause is an audience mismatch: the Pulse client${clientId ? ` "${clientId}"` : ""} must list the Historian's own client id in its audiences. Also check that the workstation and server clocks agree.`,
        { status, url, body },
      );
    case 403:
      return new AuthError(
        `The token authenticated but is not permitted to read this (403). Check the Pulse client's scopes and permissions.`,
        { status, url, body },
      );
    case 404:
      return new HistorianError(
        `Not found (404). Either the dataset does not exist on this Historian, or the server address points at something that is not a Historian. Check the dataset name against the list in the task pane.`,
        { status, url, body },
      );
    case 500:
    case 502:
    case 503:
      return new HistorianError(
        `The Historian returned ${status}.${tail ? ` It said: ${tail}` : ""} Check the Historian service and its log.`,
        { status, url, body },
      );
    default:
      return new HistorianError(
        `The Historian returned HTTP ${status}.${tail ? ` ${tail}` : ""}`,
        { status, url, body },
      );
  }
}

/**
 * Collapse any thrown value into the one-line string Excel will show.
 *
 * A custom function that throws a raw object gets a bare #VALUE! with no
 * tooltip at all, which is the worst possible outcome for the user.
 */
export function toCellMessage(err) {
  if (err instanceof Error && err.message) return err.message;
  if (typeof err === "string" && err) return err;
  return "The Timebase add-in hit an unexpected error. Open the task pane for details.";
}
