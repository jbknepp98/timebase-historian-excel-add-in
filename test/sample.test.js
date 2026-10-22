import { test } from "node:test";
import assert from "node:assert/strict";

import {
  parseSampleInterval,
  sampleBoundaries,
  lastValueSamples,
  seriesNeedingSeed,
  buildSampleGrid,
} from "../src/lib/sample.js";
import { normalizeSeries } from "../src/lib/align.js";
import { predicateFor } from "../src/lib/quality.js";
import { parseApiTimestamp, _internal } from "../src/lib/time.js";

const T = (iso) => parseApiTimestamp(iso);
const TAG = "[default]INVERTERS/INV_11A01/DATA/CURRENT_STATUS";

// Raw on-change values from the SI's system, local time GMT-05:00 (spec section 5).
const RAW = [
  ["2026-10-20T07:08:47-05:00", 6.0],
  ["2026-10-20T17:34:30-05:00", 7.0],
  ["2026-10-20T17:34:31-05:00", 4.0],
  ["2026-10-20T17:34:33-05:00", 5.0],
  ["2026-10-20T17:38:19-05:00", 4.0],
  ["2026-10-20T17:51:26-05:00", 12.0],
  ["2026-10-21T07:03:35-05:00", 4.0],
  ["2026-10-21T07:08:21-05:00", 5.0],
  ["2026-10-21T07:11:19-05:00", 6.0],
  ["2026-10-21T07:12:28-05:00", 7.0],
  ["2026-10-21T07:12:29-05:00", 4.0],
  ["2026-10-21T07:12:31-05:00", 5.0],
  ["2026-10-21T07:17:27-05:00", 19.0],
  ["2026-10-21T07:17:28-05:00", 6.0],
];

/**
 * What the Historian's /data route returns for a window: the leading boundary
 * point (last sample before start), then every sample inside [start, end].
 */
function apiResponse(startIso, endIso, raw = RAW, name = TAG) {
  const s = T(startIso);
  const e = T(endIso);
  const before = raw.filter(([t]) => T(t) < s);
  const inside = raw.filter(([t]) => T(t) >= s && T(t) <= e);
  const d = [...before.slice(-1), ...inside].map(([t, v, q]) => ({ t, v, q: q ?? 192 }));
  return { s: startIso, e: endIso, tl: [{ t: { n: name, t: "System.Double" }, d }] };
}

const serialToMs = (serial) => Math.round((serial - _internal.EXCEL_UNIX_OFFSET_DAYS) * _internal.MS_PER_DAY);

/** Grid -> [{ at: "HH:MM" in GMT-05:00, v }] for the first tag column. */
function readColumn(grid) {
  return grid.slice(1).map((row) => {
    const d = new Date(serialToMs(row[0]) - 5 * 3600000);
    const hhmm = `${String(d.getUTCHours()).padStart(2, "0")}:${String(d.getUTCMinutes()).padStart(2, "0")}`;
    return [hhmm, row[1]];
  });
}

test("spec vector: Oct 20 17:30-18:00 at 5 min", () => {
  const grid = buildSampleGrid(apiResponse("2026-10-20T17:30:00-05:00", "2026-10-20T18:00:00-05:00"), {
    utc: true,
  });
  assert.deepEqual(grid[0], ["Timestamp", TAG]);
  assert.deepEqual(readColumn(grid), [
    ["17:30", 6],
    ["17:35", 5],
    ["17:40", 4],
    ["17:45", 4],
    ["17:50", 4],
    ["17:55", 12],
    ["18:00", 12],
  ]);
});

test("spec vector: Oct 21 07:00-07:20, seed carried from the day before", () => {
  const grid = buildSampleGrid(apiResponse("2026-10-21T07:00:00-05:00", "2026-10-21T07:20:00-05:00"), {
    utc: true,
  });
  assert.deepEqual(readColumn(grid), [
    ["07:00", 12], // carried from Oct 20 17:51:26
    ["07:05", 4],
    ["07:10", 5],
    ["07:15", 5],
    ["07:20", 6], // the one-second 19 at 07:17:27 is not a snapshot value
  ]);
});

test("a window that does not start on a clock mark starts at the next one", () => {
  const grid = buildSampleGrid(apiResponse("2026-10-21T07:03:12-05:00", "2026-10-21T07:20:00-05:00"), {
    utc: true,
  });
  assert.deepEqual(
    readColumn(grid).map(([t]) => t),
    ["07:05", "07:10", "07:15", "07:20"],
  );
});

test("a point exactly on a boundary counts for that boundary (<=)", () => {
  const b = T("2026-10-21T07:05:00.000-05:00");
  const pts = [
    { ms: T("2026-10-21T06:00:00-05:00"), v: 1 },
    { ms: b, v: 2 },
  ];
  assert.deepEqual(lastValueSamples(pts, [b - 300000, b, b + 300000]), [1, 2, 2]);
});

test("no data at or before a boundary leaves the cell blank, not 0", () => {
  const raw = [["2026-10-21T07:07:00-05:00", 3.0]];
  const grid = buildSampleGrid(
    apiResponse("2026-10-21T07:00:00-05:00", "2026-10-21T07:10:00-05:00", raw),
    { utc: true },
  );
  assert.deepEqual(readColumn(grid), [
    ["07:00", null],
    ["07:05", null],
    ["07:10", 3],
  ]);
});

test("a bad-quality point between boundaries is skipped; the previous good value holds", () => {
  const raw = [
    ["2026-10-21T06:50:00-05:00", 4.0],
    ["2026-10-21T07:02:00-05:00", 99.0, 0], // bad
    ["2026-10-21T07:07:00-05:00", 5.0],
  ];
  const resp = apiResponse("2026-10-21T07:00:00-05:00", "2026-10-21T07:10:00-05:00", raw);
  const good = buildSampleGrid(resp, { utc: true, qualityPredicate: predicateFor("good") });
  assert.deepEqual(readColumn(good), [
    ["07:00", 4],
    ["07:05", 4],
    ["07:10", 5],
  ]);
  // Including bad quality is an option, not the default.
  const any = buildSampleGrid(resp, { utc: true, qualityPredicate: predicateFor("any") });
  assert.deepEqual(readColumn(any)[1], ["07:05", 99]);
});

test("float state codes become integers by default; castInt false keeps the raw type", () => {
  const pts = [{ ms: 0, v: 6.0 }, { ms: 10, v: 6.7 }];
  assert.deepEqual(lastValueSamples(pts, [5, 15]), [6, 6]);
  assert.deepEqual(lastValueSamples(pts, [5, 15], { castInt: false }), [6, 6.7]);
  assert.ok(Number.isInteger(lastValueSamples(pts, [5])[0]));
});

test("text and boolean values are never cast", () => {
  const pts = [{ ms: 0, v: "Running" }, { ms: 10, v: true }];
  assert.deepEqual(lastValueSamples(pts, [5, 15]), ["Running", true]);
});

test("a rejected boundary point is reported as needing a seed, and a lookback seed fills it", () => {
  const raw = [
    ["2026-10-18T09:00:00-05:00", 3.0], // good, three days earlier
    ["2026-10-21T06:00:00-05:00", 8.0, 0], // bad: this is the API's boundary point
    ["2026-10-21T07:07:00-05:00", 5.0],
  ];
  const resp = apiResponse("2026-10-21T07:00:00-05:00", "2026-10-21T07:10:00-05:00", raw);
  const qualityPredicate = predicateFor("good");

  const series = resp.tl.map((e) => normalizeSeries(e, { qualityPredicate }));
  assert.deepEqual(seriesNeedingSeed(series, T("2026-10-21T07:00:00-05:00")), [TAG]);

  const unseeded = buildSampleGrid(resp, { utc: true, qualityPredicate });
  assert.equal(readColumn(unseeded)[0][1], null);

  const seedResponse = apiResponse("2026-10-14T07:00:00-05:00", "2026-10-21T07:00:00-05:00", raw);
  const seeded = buildSampleGrid(resp, { utc: true, qualityPredicate, seedResponse });
  assert.deepEqual(readColumn(seeded), [
    ["07:00", 3],
    ["07:05", 3],
    ["07:10", 5],
  ]);
});

test("multiple tags share one boundary grid, in the requested column order", () => {
  const a = apiResponse("2026-10-21T07:00:00-05:00", "2026-10-21T07:10:00-05:00", RAW, "A");
  const b = apiResponse("2026-10-21T07:00:00-05:00", "2026-10-21T07:10:00-05:00", [
    ["2026-10-21T07:06:00-05:00", 1.0],
  ], "B");
  const resp = { s: a.s, e: a.e, tl: [...a.tl, ...b.tl] };
  const grid = buildSampleGrid(resp, { utc: true, tagOrder: ["B", "A", "Missing"] });
  assert.deepEqual(grid[0], ["Timestamp", "B", "A", "Missing"]);
  assert.deepEqual(grid.slice(1).map((r) => r.slice(1)), [
    [null, 12, null],
    [null, 4, null],
    [1, 5, null],
  ]);
});

test("boundaries: minute intervals align to epoch multiples", () => {
  const start = T("2026-10-21T07:03:12Z");
  const out = sampleBoundaries(start, T("2026-10-21T07:31:00Z"), 15 * 60000);
  assert.deepEqual(out.map((ms) => new Date(ms).toISOString().slice(11, 16)), ["07:15", "07:30"]);
});

test("boundaries: hourly aligns to the local hour in a half-hour zone (UTC+05:30)", () => {
  const offsetMs = () => 5.5 * 3600000;
  const out = sampleBoundaries(T("2026-10-21T00:00:00Z"), T("2026-10-21T02:00:00Z"), 3600000, { offsetMs });
  // 00:30Z and 01:30Z are 06:00 and 07:00 local.
  assert.deepEqual(out.map((ms) => new Date(ms).toISOString().slice(11, 16)), ["00:30", "01:30"]);
});

test("boundaries: daily lands on local midnight, across a DST change", () => {
  // US Central: CDT (-5) until 2026-11-01 02:00 local, then CST (-6).
  const switchAt = T("2026-11-01T07:00:00Z");
  const offsetMs = (ms) => (ms < switchAt ? -5 : -6) * 3600000;
  const out = sampleBoundaries(T("2026-10-30T12:00:00Z"), T("2026-11-03T12:00:00Z"), 86400000, { offsetMs });
  assert.deepEqual(out.map((ms) => new Date(ms).toISOString()), [
    "2026-10-31T05:00:00.000Z", // Oct 31 00:00 CDT
    "2026-11-01T05:00:00.000Z", // Nov 1 00:00 CDT
    "2026-11-02T06:00:00.000Z", // Nov 2 00:00 CST
    "2026-11-03T06:00:00.000Z", // Nov 3 00:00 CST
  ]);
});

test("boundaries: end is inclusive, and the row cap refuses rather than truncates", () => {
  assert.equal(sampleBoundaries(0, 600000, 300000).length, 3);
  assert.throws(() => sampleBoundaries(0, 3600000, 60000, { maxCount: 10 }), /more than 10 rows/);
});

test("interval: a bare number is minutes, text takes a unit, blank is 5 minutes", () => {
  assert.equal(parseSampleInterval(undefined), 300000);
  assert.equal(parseSampleInterval(""), 300000);
  assert.equal(parseSampleInterval(5), 300000);
  assert.equal(parseSampleInterval("15"), 900000);
  assert.equal(parseSampleInterval("30s"), 30000);
  assert.equal(parseSampleInterval("1h"), 3600000);
  assert.equal(parseSampleInterval("1d"), 86400000);
  assert.throws(() => parseSampleInterval(0), RangeError);
  assert.throws(() => parseSampleInterval("5 minutes"), RangeError);
});

test("a no-value point never becomes the sample; it carries the value and flags it", async () => {
  const { sampleWithFlags } = await import("../src/lib/sample.js");
  const pts = [
    { ms: 0, v: 20, issue: false },
    { ms: 50, v: null, novalue: true, issue: true }, // quality-64 note
    { ms: 150, v: 30, issue: false },
  ];
  const { values, flags } = sampleWithFlags(pts, [10, 100, 200]);
  assert.deepEqual({ values, flags }, { values: [20, 20, 30], flags: [false, true, false] });
});

test("seriesNeedingSeed: a no-value boundary point does not count as a seed", () => {
  const series = [{ name: "A", points: [{ ms: 5, v: null, novalue: true }] }];
  assert.deepEqual(seriesNeedingSeed(series, 10), ["A"]);
});
