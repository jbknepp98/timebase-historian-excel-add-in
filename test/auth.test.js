import { test } from "node:test";
import assert from "node:assert/strict";

import { normalizePulseUrl, pulseUrlFromHistorian, createTokenProvider } from "../src/lib/auth.js";
import { AuthError, ConfigError } from "../src/lib/errors.js";
import { makeFakeFetch } from "./support/fakeFetch.js";

test("normalizePulseUrl: defaults to port 4542", () => {
  assert.equal(normalizePulseUrl("pulse"), "https://pulse:4542");
  assert.equal(normalizePulseUrl("pulse:9999"), "https://pulse:9999");
});

test("pulseUrlFromHistorian: same host, Pulse's default port", () => {
  assert.equal(pulseUrlFromHistorian("https://historian:4512"), "https://historian:4542");
});

test("pulseUrlFromHistorian: blank input is blank output, not an error", () => {
  assert.equal(pulseUrlFromHistorian(""), "");
  assert.equal(pulseUrlFromHistorian(null), "");
});

function baseConfig(overrides = {}) {
  return {
    authEnabled: true,
    pulseUrl: "https://pulse:4542",
    clientId: "Excel",
    clientSecret: "s3cret",
    scope: "",
    ...overrides,
  };
}

test("getToken: returns null immediately when authEnabled is false, no request made", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => { throw new Error("must not be called"); });
  const provider = createTokenProvider({ getConfig: () => baseConfig({ authEnabled: false }), fetchImpl: fake.fetchImpl });
  assert.equal(await provider.getToken(), null);
});

test("getToken: posts form-encoded client_credentials and returns the token", async () => {
  const fake = makeFakeFetch();
  fake.on(
    (url) => url === "https://pulse:4542/auth/token",
    (_url, init) => {
      assert.equal(init.headers["Content-Type"], "application/x-www-form-urlencoded");
      const body = new URLSearchParams(init.body);
      assert.equal(body.get("grant_type"), "client_credentials");
      assert.equal(body.get("client_id"), "Excel");
      assert.equal(body.get("client_secret"), "s3cret");
      return fake.respond({ access_token: "tok-1", token_type: "Bearer", expires_in: 3600 });
    },
  );
  const provider = createTokenProvider({ getConfig: () => baseConfig(), fetchImpl: fake.fetchImpl });
  assert.equal(await provider.getToken(), "tok-1");
});

test("getToken: caches the token until it nears expiry", async () => {
  let calls = 0;
  const fake = makeFakeFetch();
  fake.on(() => true, () => {
    calls++;
    return fake.respond({ access_token: `tok-${calls}`, expires_in: 3600 });
  });

  let now = 1000;
  const provider = createTokenProvider({ getConfig: () => baseConfig(), fetchImpl: fake.fetchImpl, nowFn: () => now });

  assert.equal(await provider.getToken(), "tok-1");
  now += 1000; // well within the hour
  assert.equal(await provider.getToken(), "tok-1", "reused, no second request");
  assert.equal(calls, 1);

  now += 3600 * 1000; // past expiry
  assert.equal(await provider.getToken(), "tok-2", "re-acquired after expiry");
  assert.equal(calls, 2);
});

test("getToken: a credential change invalidates the cache even before expiry", async () => {
  let calls = 0;
  const fake = makeFakeFetch();
  fake.on(() => true, () => {
    calls++;
    return fake.respond({ access_token: `tok-${calls}`, expires_in: 3600 });
  });

  let config = baseConfig({ clientId: "First" });
  const provider = createTokenProvider({ getConfig: () => config, fetchImpl: fake.fetchImpl });

  assert.equal(await provider.getToken(), "tok-1");
  config = baseConfig({ clientId: "Second" });
  assert.equal(await provider.getToken(), "tok-2", "different client id means a different token");
  assert.equal(calls, 2);
});

test("getToken: concurrent callers collapse onto a single request", async () => {
  let calls = 0;
  const fake = makeFakeFetch();
  let releaseFirst;
  const gate = new Promise((resolve) => { releaseFirst = resolve; });
  fake.on(() => true, async () => {
    calls++;
    await gate;
    return fake.respond({ access_token: "tok-1", expires_in: 3600 });
  });

  const provider = createTokenProvider({ getConfig: () => baseConfig(), fetchImpl: fake.fetchImpl });
  const p1 = provider.getToken();
  const p2 = provider.getToken();
  releaseFirst();
  const [t1, t2] = await Promise.all([p1, p2]);
  assert.equal(t1, "tok-1");
  assert.equal(t2, "tok-1");
  assert.equal(calls, 1, "Excel recalculating many cells at once must not mint many tokens");
});

test("getToken: invalidate() forces a fresh acquisition", async () => {
  let calls = 0;
  const fake = makeFakeFetch();
  fake.on(() => true, () => {
    calls++;
    return fake.respond({ access_token: `tok-${calls}`, expires_in: 3600 });
  });
  const provider = createTokenProvider({ getConfig: () => baseConfig(), fetchImpl: fake.fetchImpl });
  assert.equal(await provider.getToken(), "tok-1");
  provider.invalidate();
  assert.equal(await provider.getToken(), "tok-2");
});

test("getToken: a wrong secret produces an AuthError naming the seeded-secret trap", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond({ error: "invalid_client", error_description: "InvalidClientSecret" }, { status: 400 }));
  const provider = createTokenProvider({ getConfig: () => baseConfig(), fetchImpl: fake.fetchImpl });
  await assert.rejects(() => provider.getToken(), (err) => {
    assert.ok(err instanceof AuthError);
    assert.match(err.message, /secret/i);
    assert.match(err.message, /InvalidClientSecret/);
    return true;
  });
});

test("getToken: an HTML response from the token endpoint is diagnosed, not parsed", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond("<html>nope</html>", { status: 200, contentType: "text/html" }));
  const provider = createTokenProvider({ getConfig: () => baseConfig(), fetchImpl: fake.fetchImpl });
  await assert.rejects(() => provider.getToken(), (err) => {
    assert.ok(err instanceof AuthError);
    assert.match(err.message, /HTML|Pulse address/);
    return true;
  });
});

test("getToken: missing pulse url or client id fail fast with a config error, no request made", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => { throw new Error("must not be called"); });

  const noPulse = createTokenProvider({ getConfig: () => baseConfig({ pulseUrl: "" }), fetchImpl: fake.fetchImpl });
  await assert.rejects(() => noPulse.getToken(), ConfigError);

  const noClientId = createTokenProvider({ getConfig: () => baseConfig({ clientId: "" }), fetchImpl: fake.fetchImpl });
  await assert.rejects(() => noClientId.getToken(), ConfigError);
});

test("getToken: scope is included only when configured", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, (_url, init) => {
    const body = new URLSearchParams(init.body);
    assert.equal(body.get("scope"), "openid");
    return fake.respond({ access_token: "tok", expires_in: 3600 });
  });
  const provider = createTokenProvider({ getConfig: () => baseConfig({ scope: "openid" }), fetchImpl: fake.fetchImpl });
  await provider.getToken();
});
