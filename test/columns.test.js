import { test } from "node:test";
import assert from "node:assert/strict";

import { parseColumns, worseQuality, statusText, rangeStats } from "../src/lib/columns.js";
import { buildGrid } from "../src/lib/align.js";
import { buildSampleGrid } from "../src/lib/sample.js";

function response(s, e, tags) {
  return {
    s,
    e,
    tl: tags.map(({ name, points }) => ({
      t: { n: name },
      d: points.map(([t, v, q]) => (v === null ? { t, q: q ?? 192 } : { t, v, q: q ?? 192 })),
    })),
  };
}

test("parseColumns: names, aliases, order kept, duplicates dropped", () => {
  assert.deepEqual(parseColumns(""), []);
  assert.deepEqual(parseColumns(undefined), []);
  assert.deepEqual(parseColumns("quality, samples, good%"), ["quality", "samples", "good%"]);
  assert.deepEqual(parseColumns([["Status"], ["count"]]), ["status", "samples"], "a range of names works too");
  assert.deepEqual(parseColumns("%good;quality;quality"), ["good%", "quality"]);
  assert.throws(() => parseColumns("median"), /quality, status, samples, good%/);
});

test("worseQuality: Bad < Uncertain < Good, ties to the lower code", () => {
  assert.equal(worseQuality(192, 64), 64);
  assert.equal(worseQuality(64, 0), 0);
  assert.equal(worseQuality(216, 192), 192, "both Good: the lower code");
  assert.equal(worseQuality(null, 192), 192);
  assert.equal(worseQuality(28, 64), 28, "28 is Bad (collector shut down), worse than Uncertain 64");
});

test("statusText: words, with the Historian's own meaning for known codes", () => {
  assert.equal(statusText(192), "Good");
  assert.equal(statusText(64), "Uncertain: store and forward (collector heartbeat lost for 10s)");
  assert.equal(statusText(0), "Bad");
  assert.equal(statusText(null), null);
});

test("rangeStats: the point in force counts toward worst quality, not toward samples", () => {
  const st = rangeStats({ q: 64 });
  st.add({ q: 192 });
  st.add({ q: 192 });
  assert.deepEqual({ worst: st.worst, count: st.count, good: st.good }, { worst: 64, count: 2, good: 2 });
});

// FL001.State-shaped data: a change, a quality-64 note, a change.
const raw = response("2026-10-21T16:00:00Z", "2026-10-21T18:05:00Z", [
  {
    name: "FL001.State",
    points: [
      ["2026-10-21T15:00:00Z", 10],
      ["2026-10-21T16:25:12Z", 20],
      ["2026-10-21T17:55:11Z", null, 64],
      ["2026-10-21T17:56:54Z", 30],
    ],
  },
]);

test("HISTORY raw: each tag's value is followed by its chosen columns", () => {
  const grid = buildGrid(raw, { utc: true, columns: ["quality", "status", "samples", "good%"] });
  assert.deepEqual(grid[0], [
    "Timestamp",
    "FL001.State",
    "FL001.State Quality",
    "FL001.State Status",
    "FL001.State Samples",
    "FL001.State % Good",
  ]);
  const body = grid.slice(1).map((r) => r.slice(1));
  assert.deepEqual(body, [
    [10, 192, "Good", 0, null], // start row: carried in, nothing logged in its range
    [20, 192, "Good", 1, 1],
    [20, 64, "Uncertain: store and forward (collector heartbeat lost for 10s)", 1, 0], // the note
    [30, 64, "Uncertain: store and forward (collector heartbeat lost for 10s)", 1, 1], // worst includes the note in force
  ]);
});

test("HISTORY slices: worst quality across the slice, samples and % Good within it", () => {
  const grid = buildGrid(raw, {
    utc: true,
    interval: 3600000, // 16:00, 17:00, 18:00
    aggregate: "last",
    columns: ["quality", "samples", "good%"],
    percentCell: (f) => Math.round(f * 100),
  });
  assert.deepEqual(
    grid.slice(1).map((r) => r.slice(1)),
    [
      [20, 192, 1, 100], // 16:00-17:00: the 16:25 change
      [30, 64, 2, 50], // 17:00-18:00: the note and the 17:56 change
      [30, 192, 0, null], // 18:00-18:05: nothing logged; 30 (good) in force
    ],
  );
});

test("SAMPLE: each row's range runs from the previous clock mark", () => {
  const r = response("2026-10-21T17:50:00Z", "2026-10-21T18:00:00Z", [
    {
      name: "FL001.State",
      points: [["2026-10-21T16:25:12Z", 20], ["2026-10-21T17:55:11Z", null, 64], ["2026-10-21T17:56:54Z", 30]],
    },
  ]);
  const grid = buildSampleGrid(r, { utc: true, columns: ["quality", "samples"] });
  assert.deepEqual(grid[0], ["Timestamp", "FL001.State", "FL001.State Quality", "FL001.State Samples"]);
  assert.deepEqual(
    grid.slice(1).map((row) => row.slice(1)),
    [
      [20, 192, 0], // 17:50: 20 in force, nothing new
      [20, 192, 0], // 17:55: still nothing (17:55:11 is after the mark)
      [30, 64, 2], // 18:00: the note and the change, worst is 64
    ],
  );
});

test("columns are off by default: output is unchanged", () => {
  assert.equal(buildGrid(raw, { utc: true })[0].length, 2);
});
