import { test } from "node:test";
import assert from "node:assert/strict";

import {
  normalizeBaseUrl,
  decodeTag,
  decodeDataset,
  buildDataQuery,
  createClient,
} from "../src/lib/historian.js";
import { ConfigError, HistorianError, AuthError } from "../src/lib/errors.js";
import { makeFakeFetch } from "./support/fakeFetch.js";

test("normalizeBaseUrl: adds https and the right default port", () => {
  assert.equal(normalizeBaseUrl("historian"), "https://historian:4512");
  assert.equal(normalizeBaseUrl("historian:9000"), "https://historian:9000");
  assert.equal(normalizeBaseUrl("http://historian"), "http://historian:4511");
  assert.equal(normalizeBaseUrl("https://historian"), "https://historian:4512");
});

test("normalizeBaseUrl: strips a path, so a pasted /api/help cannot double up", () => {
  assert.equal(normalizeBaseUrl("https://historian:4512/api/help"), "https://historian:4512");
});

test("normalizeBaseUrl: empty means same-origin, not an error", () => {
  assert.equal(normalizeBaseUrl(""), "");
  assert.equal(normalizeBaseUrl(null), "");
  assert.equal(normalizeBaseUrl(undefined), "");
});

test("normalizeBaseUrl: rejects unusable input with a message naming the fix", () => {
  assert.throws(() => normalizeBaseUrl("http://"), ConfigError);
});

test("decodeTag: expands every abbreviated field", () => {
  const t = decodeTag({
    n: "131-TT-001.PV",
    d: "Filler 1 Temperature",
    f: "0.00",
    u: { 1: "°F" },
    fl: { Site: "Southlake" },
    t: "System.Double",
    l: { t: "2026-08-10T16:42:22Z", v: 57.96, q: 192, vs: "57.96 °F" },
  });
  assert.equal(t.name, "131-TT-001.PV");
  assert.equal(t.description, "Filler 1 Temperature");
  assert.equal(t.dataType, "System.Double");
  assert.deepEqual(t.units, { 1: "°F" });
  assert.equal(t.last.display, "57.96 °F");
});

test("decodeTag: a tag with no last-value snapshot decodes with last: null", () => {
  const t = decodeTag({ n: "A" });
  assert.equal(t.last, null);
});

test("decodeTag: junk input decodes to null rather than throwing", () => {
  assert.equal(decodeTag(null), null);
  assert.equal(decodeTag("not an object"), null);
});

test("decodeDataset: expands every abbreviated field", () => {
  const d = decodeDataset({ n: "OPCUA Demo", pa: 0, ps: 0, ldt: 100, lda: 30, s: 123.45, c: 97831, a: 640.53 });
  assert.equal(d.name, "OPCUA Demo");
  assert.equal(d.lateDataToleranceMs, 100);
  assert.equal(d.tagCount, 97831);
});

test("buildDataQuery: repeats tagname once per tag, per the documented batch form", () => {
  const q = buildDataQuery({
    tags: ["131-TT-001.PV", "131-TT-002.PV"],
    startParam: { param: "relativeStart", value: "-1h" },
    endParam: null,
  });
  const params = new URLSearchParams(q);
  assert.deepEqual(params.getAll("tagname"), ["131-TT-001.PV", "131-TT-002.PV"]);
  assert.equal(params.get("relativeStart"), "-1h");
  assert.equal(params.has("end"), false);
});

test("client.listDatasets: decodes the array response", async () => {
  const { fetchImpl, on } = makeFakeFetch();
  on(
    (url) => url === "https://historian:4512/api/datasets",
    (_url, init) => {
      assert.equal(init.method, "GET");
      return { ok: true, status: 200, type: "basic", headers: { get: () => "application/json" }, text: async () => JSON.stringify([{ n: "A", pa: 0, ps: 0 }]) };
    },
  );
  const client = createClient({ baseUrl: "historian", fetchImpl });
  const datasets = await client.listDatasets();
  assert.equal(datasets.length, 1);
  assert.equal(datasets[0].name, "A");
});

test("client: attaches a bearer token when a token provider is given", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond([]));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl, getToken: async () => "abc123" });
  await client.listDatasets();
  assert.equal(fake.calls[0].init.headers.Authorization, "Bearer abc123");
});

test("client: sends no Authorization header when getToken resolves null (auth disabled)", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond([]));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl, getToken: async () => null });
  await client.listDatasets();
  assert.equal("Authorization" in fake.calls[0].init.headers, false);
});

test("client: a redirect is reported as an error, never silently followed", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond("", { status: 302 }));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  await assert.rejects(() => client.listDatasets(), HistorianError);
});

test("client: 401 becomes an AuthError naming the audience as the likely cause", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond("", { status: 401, contentType: "text/plain" }));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  await assert.rejects(() => client.listDatasets(), (err) => {
    assert.ok(err instanceof AuthError);
    assert.match(err.message, /audience/i);
    return true;
  });
});

test("client: an HTML body on a 200 is treated as a routing failure, not data", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond("<!doctype html><html>nope</html>", { status: 200, contentType: "text/html" }));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  await assert.rejects(() => client.listDatasets(), (err) => {
    assert.ok(err instanceof HistorianError);
    assert.match(err.message, /HTML/);
    return true;
  });
});

test("client: a body that claims 200 but is not JSON fails clearly", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond("not json at all", { status: 200, contentType: "application/json" }));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  await assert.rejects(() => client.listDatasets(), /not valid JSON/);
});

test("client.listTags: passes query and contains through as separate parameters", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, (url) => {
    const u = new URL(url);
    assert.equal(u.pathname, "/api/datasets/MyDataset/tags");
    assert.equal(u.searchParams.get("query"), "name:Pump_*");
    assert.equal(u.searchParams.get("contains"), "temp");
    return fake.respond([]);
  });
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  await client.listTags("MyDataset", { query: "name:Pump_*", contains: "temp" });
});

test("client.readData: refuses an empty tag list before making a request", async () => {
  const client = createClient({ baseUrl: "historian", fetchImpl: async () => { throw new Error("should not be called"); } });
  await assert.rejects(() => client.readData("A", [], null, null), ConfigError);
});

test("client.readData: a null body (no data) normalises to an empty envelope", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, () => fake.respond("", { status: 200 }));
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  const result = await client.readData("A", ["tag1"], null, null);
  assert.deepEqual(result, { s: null, e: null, tl: [] });
});

test("client: dataset and tag names are percent-encoded in the path", async () => {
  const fake = makeFakeFetch();
  fake.on(() => true, (url) => {
    assert.ok(url.includes("/api/datasets/My%20Dataset/tags"), url);
    return fake.respond([]);
  });
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl });
  await client.listTags("My Dataset");
});

test("client: a timeout raises an error that suggests narrowing the range", async () => {
  const fake = { fetchImpl: async (_url, init) => new Promise((_resolve, reject) => {
    init.signal.addEventListener("abort", () => {
      const err = new Error("aborted");
      err.name = "AbortError";
      reject(err);
    });
  }) };
  const client = createClient({ baseUrl: "historian", fetchImpl: fake.fetchImpl, timeoutMs: 10 });
  await assert.rejects(() => client.listDatasets(), (err) => {
    assert.ok(err instanceof HistorianError);
    assert.match(err.message, /interval|narrow/i);
    return true;
  });
});

test("tagListFromResponse: accepts Historian 2.0's grouped shape and the bare array", async () => {
  const { tagListFromResponse } = await import("../src/lib/historian.js");
  const a = { n: "FL001.State" };
  const b = { n: "131-TT-001.PV" };
  const sys = { n: "Dataset.State" };
  assert.deepEqual(tagListFromResponse({ User: [a, b], System: [sys] }), [a, b, sys]);
  assert.deepEqual(tagListFromResponse([a, b]), [a, b]);
  assert.deepEqual(tagListFromResponse(null), []);
  assert.deepEqual(tagListFromResponse({ User: null }), []);
});
