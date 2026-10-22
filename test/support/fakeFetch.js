// A minimal fetch stand-in for tests: no network, no timers left running.
//
// Register handlers by exact URL (including query string) or by a predicate,
// checked in registration order. Each call is logged so tests can assert on
// what was actually requested (headers included, for auth tests).

export function makeFakeFetch() {
  const routes = [];
  const calls = [];

  function respond(body, { status = 200, contentType = "application/json" } = {}) {
    const text = typeof body === "string" ? body : JSON.stringify(body);
    return {
      ok: status >= 200 && status < 300,
      status,
      type: status >= 300 && status < 400 ? "opaqueredirect" : "basic",
      headers: { get: (name) => (name.toLowerCase() === "content-type" ? contentType : null) },
      async text() {
        return text;
      },
    };
  }

  function on(matcher, handler) {
    routes.push({ matcher, handler });
  }

  async function fetchImpl(url, init) {
    calls.push({ url, init });
    for (const { matcher, handler } of routes) {
      const matches = typeof matcher === "function" ? matcher(url, init) : url === matcher;
      if (matches) return handler(url, init);
    }
    throw new Error(`fakeFetch: no route registered for ${url}`);
  }

  return { fetchImpl, on, respond, calls };
}
