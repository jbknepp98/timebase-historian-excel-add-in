import { test } from "node:test";
import assert from "node:assert/strict";

import {
  normalizeSeries,
  splitBoundary,
  unionAxis,
  alignStepped,
  alignBucketed,
  buildGrid,
} from "../src/lib/align.js";
import { predicateFor } from "../src/lib/quality.js";
import { parseApiTimestamp, msToExcelSerial } from "../src/lib/time.js";

const T = (iso) => parseApiTimestamp(iso);

/** Build a TVQsResponse the way the API documents it. */
function response(startIso, endIso, tags) {
  return {
    s: startIso,
    e: endIso,
    tl: tags.map(({ name, points }) => ({
      t: { n: name, t: "System.Double" },
      d: points.map(([t, v, q]) => ({ t, v, q: q ?? 192 })),
    })),
  };
}

/** Read a grid back as { time, values } using UTC so assertions are stable. */
function readGrid(grid) {
  const [header, ...rows] = grid;
  return {
    header,
    rows: rows.map((r) => ({ serial: r[0], values: r.slice(1) })),
  };
}

test("normalizeSeries: parses timestamps and sorts", () => {
  const s = normalizeSeries({
    t: { n: "A" },
    d: [
      { t: "2026-08-10T00:00:02Z", v: 2, q: 192 },
      { t: "2026-08-10T00:00:01Z", v: 1, q: 192 },
    ],
  });
  assert.equal(s.name, "A");
  assert.deepEqual(
    s.points.map((p) => p.v),
    [1, 2],
    "late-arriving samples are sorted into place",
  );
});

test("normalizeSeries: drops samples with unusable timestamps rather than failing", () => {
  const s = normalizeSeries({
    t: { n: "A" },
    d: [
      { t: "2026-08-10T00:00:01Z", v: 1 },
      { t: "not a date", v: 999 },
      { t: null, v: 998 },
      { v: 997 },
    ],
  });
  assert.equal(s.points.length, 1);
  assert.equal(s.points[0].v, 1);
});

test("normalizeSeries: quality filter runs before alignment", () => {
  const s = normalizeSeries(
    {
      t: { n: "A" },
      d: [
        { t: "2026-08-10T00:00:01Z", v: 1, q: 192 },
        { t: "2026-08-10T00:00:02Z", v: 2, q: 0 },
        { t: "2026-08-10T00:00:03Z", v: 3, q: 64 },
      ],
    },
    { qualityPredicate: predicateFor("good") },
  );
  assert.deepEqual(
    s.points.map((p) => p.v),
    [1],
    "bad and uncertain samples never reach the aligner",
  );
});

test("splitBoundary: finds the pre-window sample by time, not by index", () => {
  const points = [
    { ms: T("2026-08-10T00:00:00Z"), v: "before" },
    { ms: T("2026-08-10T01:00:00Z"), v: "inside" },
  ];
  const { boundary, inside } = splitBoundary(points, T("2026-08-10T00:30:00Z"));
  assert.equal(boundary.v, "before");
  assert.equal(inside.length, 1);
  assert.equal(inside[0].v, "inside");
});

test("splitBoundary: a series with no earlier history keeps its first real sample", () => {
  // The bug this guards: assuming d[0] is always the boundary point would
  // silently discard the first in-window sample for a tag that only started
  // logging inside the window.
  const points = [
    { ms: T("2026-08-10T01:00:00Z"), v: "first real" },
    { ms: T("2026-08-10T02:00:00Z"), v: "second" },
  ];
  const { boundary, inside } = splitBoundary(points, T("2026-08-10T00:30:00Z"));
  assert.equal(boundary, null);
  assert.equal(inside.length, 2, "nothing was dropped");
  assert.equal(inside[0].v, "first real");
});

test("splitBoundary: a sample exactly at the window start is inside it", () => {
  const start = T("2026-08-10T00:30:00Z");
  const { boundary, inside } = splitBoundary([{ ms: start, v: "on the edge" }], start);
  assert.equal(boundary, null, "half-open window: start is included");
  assert.equal(inside.length, 1);
});

test("unionAxis: merges and de-duplicates timestamps across tags", () => {
  const axis = unionAxis([
    { points: [{ ms: 100 }, { ms: 300 }] },
    { points: [{ ms: 200 }, { ms: 300 }] },
  ]);
  assert.deepEqual(axis, [100, 200, 300]);
});

test("alignStepped: carries the last known value forward per tag", () => {
  // A is logged at 100 and 300; B only at 200. Every cell must be populated
  // from the most recent sample at or before the row's timestamp.
  const a = { points: [{ ms: 100, v: "a1" }, { ms: 300, v: "a2" }] };
  const b = { points: [{ ms: 200, v: "b1" }] };
  const rows = alignStepped([a, b], [100, 200, 300], [null, null]);
  assert.deepEqual(rows.map((r) => r.values), [
    ["a1", null],
    ["a1", "b1"],
    ["a2", "b1"],
  ]);
});

test("alignStepped: the boundary sample seeds the first row", () => {
  // This is why the boundary point matters. Without it, B's first cell would
  // be blank even though the tag held a good value for the whole window.
  const a = { points: [{ ms: 100, v: "a1" }] };
  const b = { points: [] };
  const rows = alignStepped([a, b], [100], [null, { ms: 50, v: "b-before" }]);
  assert.deepEqual(rows[0].values, ["a1", "b-before"]);
});

test("alignBucketed: buckets are half-open so a boundary sample is not double counted", () => {
  const s = {
    points: [
      { ms: 0, v: 1 },
      { ms: 1000, v: 2 }, // exactly on the second bucket's start
      { ms: 1500, v: 4 },
    ],
  };
  const rows = alignBucketed([s], [null], {
    startMs: 0,
    endMs: 2000,
    intervalMs: 1000,
    aggregate: "count",
  });
  assert.deepEqual(rows.map((r) => r.values[0]), [1, 2], "1 in [0,1000), 2 in [1000,2000)");
});

test("alignBucketed: aggregates over the bucket; avg is time-weighted", () => {
  const s = {
    points: [
      { ms: 0, v: 10 },
      { ms: 100, v: 20 },
      { ms: 200, v: 30 },
    ],
  };
  const opts = { startMs: 0, endMs: 1000, intervalMs: 1000 };
  const of = (kind) => alignBucketed([s], [null], { ...opts, aggregate: kind })[0].values[0];
  // 10 for 100ms, 20 for 100ms, 30 for the remaining 800ms of the bucket.
  assert.equal(of("avg"), (10 * 100 + 20 * 100 + 30 * 800) / 1000);
  assert.equal(of("min"), 10);
  assert.equal(of("max"), 30);
  assert.equal(of("first"), 10);
  assert.equal(of("last"), 30);
  assert.equal(of("count"), 3);
  assert.equal(of("sum"), 60);
  assert.equal(of("range"), 20);
});

test("alignBucketed: a bucket with no change carries the last known value", () => {
  // No sample in a bucket means the value did not change, not that it is unknown.
  const s = { points: [{ ms: 0, v: 5 }] };
  const opts = { startMs: 0, endMs: 2000, intervalMs: 1000 };
  const of = (kind) => alignBucketed([s], [null], { ...opts, aggregate: kind })[1].values[0];
  for (const kind of ["avg", "min", "max", "first", "last"]) assert.equal(of(kind), 5, kind);
  assert.equal(of("range"), 0);
  assert.equal(of("count"), 0, "count is of samples logged in the bucket");
  assert.equal(of("sum"), 0);
});

test("alignBucketed: a bucket starts from the value held at its start", () => {
  // Held 10 coming in; changes to 20 at 300ms into the 1000ms bucket.
  const s = { points: [{ ms: 300, v: 20 }] };
  const opts = { startMs: 0, endMs: 1000, intervalMs: 1000 };
  const of = (kind) => alignBucketed([s], [{ ms: -50, v: 10 }], { ...opts, aggregate: kind })[0].values[0];
  assert.equal(of("first"), 10, "value in force at the bucket start");
  assert.equal(of("last"), 20);
  assert.equal(of("min"), 10);
  assert.equal(of("max"), 20);
  assert.equal(of("avg"), (10 * 300 + 20 * 700) / 1000);
  assert.equal(of("count"), 1);
});

test("alignBucketed: a sample exactly on the bucket start replaces the held value", () => {
  const s = { points: [{ ms: 0, v: 7 }] };
  const opts = { startMs: 0, endMs: 1000, intervalMs: 1000 };
  const of = (kind) => alignBucketed([s], [{ ms: -50, v: 99 }], { ...opts, aggregate: kind })[0].values[0];
  assert.equal(of("first"), 7);
  assert.equal(of("max"), 7, "the held 99 was never in force inside this bucket");
  assert.equal(of("avg"), 7);
});

test("alignBucketed: time-weighting stops at the window end", () => {
  // Window ends at 1500ms inside the second 1000ms bucket.
  const s = { points: [{ ms: 1000, v: 2 }, { ms: 1250, v: 4 }] };
  const rows = alignBucketed([s], [{ ms: -1, v: 0 }], { startMs: 0, endMs: 1500, intervalMs: 1000, aggregate: "avg" });
  assert.equal(rows[1].values[0], (2 * 250 + 4 * 250) / 500);
});

test("alignBucketed: non-numeric values do not corrupt numeric aggregates", () => {
  const s = {
    points: [
      { ms: 0, v: 10 },
      { ms: 100, v: "Running" },
      { ms: 200, v: null },
      { ms: 300, v: 20 },
    ],
  };
  const rows = alignBucketed([s], [null], {
    startMs: 0,
    endMs: 1000,
    intervalMs: 1000,
    aggregate: "avg",
  });
  // 10 for 100ms, 20 for 700ms; the text and null segments carry no weight.
  assert.equal(rows[0].values[0], (10 * 100 + 20 * 700) / 800, "averages only the real numbers");
});

test("alignBucketed: string tags survive first/last", () => {
  const s = { points: [{ ms: 0, v: "Idle" }, { ms: 100, v: "Running" }] };
  const rows = alignBucketed([s], [null], {
    startMs: 0,
    endMs: 1000,
    intervalMs: 1000,
    aggregate: "last",
  });
  assert.equal(rows[0].values[0], "Running");
});

test("alignBucketed: rejects an interval or window that makes no sense", () => {
  const s = { points: [] };
  assert.throws(
    () => alignBucketed([s], [null], { startMs: 0, endMs: 1000, intervalMs: 0, aggregate: "avg" }),
    RangeError,
  );
  assert.throws(
    () => alignBucketed([s], [null], { startMs: 1000, endMs: 0, intervalMs: 100, aggregate: "avg" }),
    RangeError,
  );
});

test("buildGrid: two tags on different clocks become one aligned table", () => {
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:01:00Z", [
    {
      name: "TT-001",
      points: [
        ["2026-08-09T23:59:50Z", 50], // boundary
        ["2026-08-10T00:00:10Z", 51],
        ["2026-08-10T00:00:40Z", 52],
      ],
    },
    {
      name: "TT-002",
      points: [
        ["2026-08-09T23:59:55Z", 20], // boundary
        ["2026-08-10T00:00:25Z", 21],
      ],
    },
  ]);
  const { header, rows } = readGrid(buildGrid(r, { utc: true }));

  assert.deepEqual(header, ["Timestamp", "TT-001", "TT-002"]);
  assert.equal(rows.length, 4, "the window start, then the three in-window timestamps");
  assert.deepEqual(rows.map((x) => x.values), [
    [50, 20], // 00:00:00 -- the start row carries both boundary values
    [51, 20], // 00:00:10 -- TT-002 still holds its boundary value
    [51, 21], // 00:00:25
    [52, 21], // 00:00:40
  ]);
  assert.equal(
    rows[1].serial,
    msToExcelSerial(T("2026-08-10T00:00:10Z"), { utc: true }),
    "timestamps come back as Excel serials",
  );
});

test("buildGrid: raw mode starts AT the window start with the last known value", () => {
  // The user's rule: start at 12:00 with no change at 12:00, and the first row
  // is still 12:00, holding whatever the tag held going in.
  const r = response("2026-10-22T12:00:00Z", "2026-10-22T13:00:00Z", [
    { name: "A", points: [["2026-10-22T09:14:00Z", 4], ["2026-10-22T12:40:00Z", 5]] },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true }));
  assert.deepEqual(
    rows.map((x) => [x.serial, x.values[0]]),
    [
      [msToExcelSerial(T("2026-10-22T12:00:00Z"), { utc: true }), 4],
      [msToExcelSerial(T("2026-10-22T12:40:00Z"), { utc: true }), 5],
    ],
  );
});

test("buildGrid: no extra start row when a sample already lands on the start", () => {
  const r = response("2026-10-22T12:00:00Z", "2026-10-22T13:00:00Z", [
    { name: "A", points: [["2026-10-22T09:14:00Z", 4], ["2026-10-22T12:00:00Z", 6]] },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true }));
  assert.deepEqual(rows.map((x) => x.values[0]), [6]);
});

test("buildGrid: a lookback seed stands in for a missing or rejected boundary point", () => {
  // The Historian's boundary point here is a quality-64 status point with no
  // value; under the good filter it is dropped, and the seed read supplies 3.
  const r = response("2026-10-22T12:00:00Z", "2026-10-22T13:00:00Z", [
    { name: "A", points: [["2026-10-22T11:30:00Z", null, 64], ["2026-10-22T12:40:00Z", 5]] },
  ]);
  const seedResponse = response("2026-10-15T12:00:00Z", "2026-10-22T12:00:00Z", [
    { name: "A", points: [["2026-10-21T08:00:00Z", 3], ["2026-10-22T11:30:00Z", null, 64]] },
  ]);
  const qualityPredicate = predicateFor("good");
  const raw = readGrid(buildGrid(r, { utc: true, qualityPredicate, seedResponse }));
  assert.deepEqual(raw.rows.map((x) => x.values[0]), [3, 5]);
  const bucketed = readGrid(
    buildGrid(r, { utc: true, qualityPredicate, seedResponse, interval: 1800000, aggregate: "first" }),
  );
  assert.deepEqual(bucketed.rows.map((x) => x.values[0]), [3, 3], "12:00 and 12:30 slices both start at 3");
});

test("buildGrid: the pre-window sample can still get its own row, or none", () => {
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:01:00Z", [
    { name: "A", points: [["2026-08-09T23:59:50Z", 1], ["2026-08-10T00:00:10Z", 2]] },
  ]);
  assert.equal(readGrid(buildGrid(r, { utc: true, boundary: "drop" })).rows.length, 1);

  const kept = readGrid(buildGrid(r, { utc: true, boundary: "keep" }));
  assert.equal(kept.rows.length, 2);
  assert.equal(kept.rows[0].values[0], 1, "the pre-window value gets its own row");
  assert.ok(kept.rows[0].serial < kept.rows[1].serial, "and sorts first");
});

test("buildGrid: tagOrder fixes the column order and invents empty columns", () => {
  // The API may return tags in any order, and a tag with no data at all may be
  // missing from tl entirely. A report bound to column positions must not shift.
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:01:00Z", [
    { name: "B", points: [["2026-08-10T00:00:10Z", 2]] },
  ]);
  const { header, rows } = readGrid(
    buildGrid(r, { utc: true, tagOrder: ["A", "B", "C"] }),
  );
  assert.deepEqual(header, ["Timestamp", "A", "B", "C"]);
  assert.deepEqual(rows[0].values, [null, 2, null]);
});

test("buildGrid: refuses to silently truncate an oversized result", () => {
  const points = [];
  for (let i = 0; i < 500; i++) points.push([new Date(Date.UTC(2026, 7, 10, 0, 0, i)).toISOString(), i]);
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:10:00Z", [{ name: "A", points }]);

  assert.throws(
    () => buildGrid(r, { maxRows: 100 }),
    (err) => {
      assert.ok(err instanceof RangeError);
      assert.match(err.message, /500/, "says how many rows it actually is");
      assert.match(err.message, /interval/, "says how to fix it");
      return true;
    },
  );
});

test("buildGrid: a decimated pull fits where the raw one does not", () => {
  const points = [];
  for (let i = 0; i < 600; i++) points.push([new Date(Date.UTC(2026, 7, 10, 0, 0, i)).toISOString(), i]);
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:10:00Z", [{ name: "A", points }]);

  const grid = buildGrid(r, { utc: true, interval: 60000, aggregate: "avg", maxRows: 100 });
  const { rows } = readGrid(grid);
  assert.equal(rows.length, 10, "ten one-minute buckets over a ten-minute window");
  assert.equal(rows[0].values[0], 29.5, "average of 0..59");
});

test("buildGrid: an empty response still returns a usable array", () => {
  // A zero-length spilled array shows as #CALC! in Excel, which reads as a
  // broken formula rather than an empty range.
  const grid = buildGrid({ s: "2026-08-10T00:00:00Z", e: "2026-08-10T00:01:00Z", tl: [] });
  assert.ok(Array.isArray(grid) && grid.length >= 1);
  assert.ok(grid[0].length >= 1);
});

test("buildGrid: a tag present but with no samples yields a blank column, not an error", () => {
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:01:00Z", [
    { name: "A", points: [["2026-08-10T00:00:10Z", 1]] },
    { name: "B", points: [] },
  ]);
  const { header, rows } = readGrid(buildGrid(r, { utc: true }));
  assert.deepEqual(header, ["Timestamp", "A", "B"]);
  assert.deepEqual(rows[0].values, [1, null]);
});

test("buildGrid: headers can be turned off for appending to an existing table", () => {
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:01:00Z", [
    { name: "A", points: [["2026-08-10T00:00:10Z", 1]] },
  ]);
  const grid = buildGrid(r, { utc: true, headers: false });
  assert.equal(grid.length, 1);
  assert.equal(typeof grid[0][0], "number", "first cell is a timestamp serial, not the word Timestamp");
});

test("buildGrid: quality filtering flows through to the grid", () => {
  const r = response("2026-08-10T00:00:00Z", "2026-08-10T00:01:00Z", [
    {
      name: "A",
      points: [
        ["2026-08-10T00:00:10Z", 1, 192],
        ["2026-08-10T00:00:20Z", 999, 0],
        ["2026-08-10T00:00:30Z", 3, 192],
      ],
    },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true, qualityPredicate: predicateFor("good") }));
  assert.deepEqual(rows.map((x) => x.values[0]), [1, 3], "the bad sample is gone entirely");
});

test("buildGrid: bucketing without a reported window is an explicit error", () => {
  const r = { tl: [{ t: { n: "A" }, d: [{ t: "2026-08-10T00:00:00Z", v: 1, q: 192 }] }] };
  assert.throws(() => buildGrid(r, { interval: 1000 }), /window/i);
});

test("normalizeAggregate: avg is the time-weighted average, and its long names say so", async () => {
  const { normalizeAggregate } = await import("../src/lib/align.js");
  for (const name of [undefined, "", "avg", "AVG", "twa", "TimeWeightedAverage", "time-weighted average"]) {
    assert.equal(normalizeAggregate(name), "avg", String(name));
  }
  assert.equal(normalizeAggregate("Max"), "max");
  assert.throws(() => normalizeAggregate("median"), /time-weighted average/);
});

// --- points with no value (quality-64 "collector heartbeat lost" notes) -------

const FLAG = (v) => ({ flagged: v });

test("buildGrid: a no-value point keeps its row, carries the value, and is flagged", () => {
  const r = response("2026-10-21T16:00:00Z", "2026-10-21T19:00:00Z", [
    {
      name: "FL001.State",
      points: [
        ["2026-10-21T15:00:00Z", 10],
        ["2026-10-21T16:25:12Z", 20],
        ["2026-10-21T17:55:11Z", null, 64], // the Historian's note: no value
        ["2026-10-21T17:56:54Z", 30],
      ],
    },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true, flagCell: FLAG }));
  assert.deepEqual(rows.map((x) => x.values[0]), [10, 20, FLAG(20), 30], "start, change, carried + flagged, cleared");
});

test("buildGrid: novalue 'drop' leaves those rows out and carries straight through", () => {
  const r = response("2026-10-21T16:00:00Z", "2026-10-21T19:00:00Z", [
    {
      name: "A",
      points: [["2026-10-21T15:00:00Z", 10], ["2026-10-21T17:55:11Z", null, 64], ["2026-10-21T17:56:54Z", 30]],
    },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true, flagCell: FLAG, novalue: "drop" }));
  assert.deepEqual(rows.map((x) => x.values[0]), [10, 30]);
  assert.throws(() => buildGrid(r, { novalue: "hide" }), /flag" or "drop/);
});

test("buildGrid: a bad-quality point WITH a value is shown, flagged", () => {
  const r = response("2026-10-21T16:00:00Z", "2026-10-21T19:00:00Z", [
    { name: "A", points: [["2026-10-21T15:00:00Z", 10], ["2026-10-21T17:00:00Z", 99, 0], ["2026-10-21T18:00:00Z", 11]] },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true, flagCell: FLAG }));
  assert.deepEqual(rows.map((x) => x.values[0]), [10, FLAG(99), 11]);
});

test("buildGrid: a no-value boundary point still carries the last real value in, flagged", () => {
  // The Historian's boundary point is itself a quality-64 note; the real value
  // before it comes from the lookback read.
  const r = response("2026-10-22T12:00:00Z", "2026-10-22T13:00:00Z", [
    { name: "A", points: [["2026-10-22T11:30:00Z", null, 64], ["2026-10-22T12:40:00Z", 5]] },
  ]);
  const seedResponse = response("2026-10-15T12:00:00Z", "2026-10-22T12:00:00Z", [
    { name: "A", points: [["2026-10-21T08:00:00Z", 3], ["2026-10-22T11:30:00Z", null, 64]] },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true, flagCell: FLAG, seedResponse }));
  assert.deepEqual(rows.map((x) => x.values[0]), [FLAG(3), 5]);
});

test("buildGrid: a slice containing a quality issue is flagged; its aggregate ignores the empty point", () => {
  const r = response("2026-10-21T17:50:00Z", "2026-10-21T18:05:00Z", [
    {
      name: "A",
      points: [["2026-10-21T17:00:00Z", 20], ["2026-10-21T17:55:11Z", null, 64], ["2026-10-21T17:56:54Z", 30]],
    },
  ]);
  const { rows } = readGrid(buildGrid(r, { utc: true, flagCell: FLAG, interval: 300000, aggregate: "last" }));
  assert.deepEqual(rows.map((x) => x.values[0]), [20, FLAG(30), 30], "17:50 clean, 17:55 flagged, 18:00 clean");
});
