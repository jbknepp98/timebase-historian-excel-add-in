import { test } from "node:test";
import assert from "node:assert/strict";

import {
  parseRelative,
  isRelative,
  excelSerialToIsoLocal,
  parseApiTimestamp,
  msToExcelSerial,
  toQueryParam,
  parseInterval,
} from "../src/lib/time.js";

// Every example the API description gives, so the grammar is pinned to the
// documented behaviour rather than to my reading of it.
test("relative: accepts every documented example", () => {
  const cases = {
    "-1h": { clamp: null, offset: { sign: -1, count: 1, unit: "h" } },
    "-90m": { clamp: null, offset: { sign: -1, count: 90, unit: "m" } },
    "-5m": { clamp: null, offset: { sign: -1, count: 5, unit: "m" } },
    "8h": { clamp: null, offset: { sign: 1, count: 8, unit: "h" } },
    h: { clamp: "h", offset: null },
    d: { clamp: "d", offset: null },
    "d+6h": { clamp: "d", offset: { sign: 1, count: 6, unit: "h" } },
    "h-2h": { clamp: "h", offset: { sign: -1, count: 2, unit: "h" } },
  };
  for (const [text, expected] of Object.entries(cases)) {
    const got = parseRelative(text);
    assert.ok(got, `expected "${text}" to parse`);
    assert.equal(got.clamp, expected.clamp, `clamp for "${text}"`);
    assert.deepEqual(got.offset, expected.offset, `offset for "${text}"`);
  }
});

test("relative: month and minute are distinguished by case", () => {
  assert.equal(parseRelative("-1M").offset.unit, "M", "M is months");
  assert.equal(parseRelative("-1m").offset.unit, "m", "m is minutes");
  assert.equal(parseRelative("M").clamp, "M");
  assert.equal(parseRelative("m").clamp, "m");
});

test("relative: w is a clamp but not an offset unit", () => {
  // The API lists w under clamp only, not under the offset intervals.
  assert.equal(parseRelative("w").clamp, "w");
  assert.equal(parseRelative("-2w"), null, "weeks are not a valid offset unit");
});

test("relative: rejects things that are not relative expressions", () => {
  for (const bad of [
    "",
    "   ",
    "yesterday",
    "-1x",
    "1",
    "h+",
    "hh",
    "2026-08-10T00:00:00Z",
    "-1.5h",
    "z-1h",
    null,
    undefined,
    42,
  ]) {
    assert.equal(isRelative(bad), false, `expected ${JSON.stringify(bad)} to be rejected`);
  }
});

test("relative: an ISO timestamp is never mistaken for a relative expression", () => {
  // This is the ambiguity that matters: both can start with a digit.
  assert.equal(isRelative("2026-08-10T17:01:58Z"), false);
  assert.equal(toQueryParam("2026-08-10T17:01:58Z", "start").param, "start");
  assert.equal(toQueryParam("-24h", "start").param, "relativeStart");
});

test("excel serial: the epoch anchor is right", () => {
  // 25569 is 1970-01-01. If this drifts, every query is off by whole days.
  assert.equal(excelSerialToIsoLocal(25569), "1970-01-01T00:00:00");
});

test("excel serial: carries wall-clock time with no zone suffix", () => {
  // 45884.5 = 2025-08-14 12:00. The absence of Z is the point: adding one
  // would shift the query by the local UTC offset.
  const iso = excelSerialToIsoLocal(45883.5);
  assert.match(iso, /^\d{4}-\d{2}-\d{2}T12:00:00$/);
  assert.ok(!iso.endsWith("Z"), "must not claim UTC");
  assert.ok(!/[+-]\d{2}:\d{2}$/.test(iso), "must not carry an offset");
});

test("excel serial: keeps milliseconds when present and omits them when not", () => {
  assert.equal(excelSerialToIsoLocal(25569), "1970-01-01T00:00:00");
  const withMs = excelSerialToIsoLocal(25569 + 1.5 / 86400);
  assert.equal(withMs, "1970-01-01T00:00:01.500");
});

test("excel serial: refuses dates inside Excel's leap-year bug", () => {
  assert.throws(() => excelSerialToIsoLocal(59), RangeError);
  assert.throws(() => excelSerialToIsoLocal(0), RangeError);
  assert.throws(() => excelSerialToIsoLocal("45883"), RangeError);
});

test("api timestamp: handles all three documented shapes", () => {
  const utc = parseApiTimestamp("2026-08-10T17:01:58.832Z");
  assert.equal(utc, Date.UTC(2026, 7, 10, 17, 1, 58, 832));

  const offset = parseApiTimestamp("2026-08-10T17:01:58.832+00:00");
  assert.equal(offset, utc, "an explicit +00:00 offset equals Z");

  // A bare timestamp is local, so it must NOT equal the UTC reading unless the
  // machine happens to run at UTC.
  const bare = parseApiTimestamp("2026-08-10T17:01:58.832");
  const localOffsetMs = new Date(utc).getTimezoneOffset() * 60000;
  assert.equal(bare, utc + localOffsetMs);
});

test("api timestamp: tolerates sub-millisecond precision", () => {
  // The API examples carry six fractional digits.
  const six = parseApiTimestamp("2026-08-10T17:01:58.832941Z");
  const three = parseApiTimestamp("2026-08-10T17:01:58.832Z");
  assert.equal(six, three, "extra precision truncates rather than corrupting");
});

test("api timestamp: rejects junk", () => {
  for (const bad of ["", "  ", "not a date", null, 42]) {
    assert.throws(() => parseApiTimestamp(bad), RangeError);
  }
});

test("serial round trip: utc mode is exactly reversible", () => {
  const iso = "2026-08-10T17:01:58.000Z";
  const ms = parseApiTimestamp(iso);
  const serial = msToExcelSerial(ms, { utc: true });
  assert.equal(excelSerialToIsoLocal(serial), "2026-08-10T17:01:58");
});

test("serial: local mode shifts by the local offset", () => {
  const ms = Date.UTC(2026, 7, 10, 12, 0, 0);
  const utcSerial = msToExcelSerial(ms, { utc: true });
  const localSerial = msToExcelSerial(ms, { utc: false });
  const expectedShiftDays = (-new Date(ms).getTimezoneOffset() * 60000) / 86400000;
  assert.ok(
    Math.abs(localSerial - utcSerial - expectedShiftDays) < 1e-9,
    "local serial differs from UTC by exactly the zone offset",
  );
});

test("toQueryParam: picks the right parameter name per bound", () => {
  assert.equal(toQueryParam("-24h", "start").param, "relativeStart");
  assert.equal(toQueryParam("-1h", "end").param, "relativeEnd");
  assert.equal(toQueryParam(45883.5, "start").param, "start");
  assert.equal(toQueryParam(45883.5, "end").param, "end");
});

test("toQueryParam: blank input means 'omit this bound'", () => {
  for (const empty of [null, undefined, ""]) {
    assert.equal(toQueryParam(empty, "start"), null);
  }
});

test("toQueryParam: rejects a nonsense bound with a useful message", () => {
  assert.throws(() => toQueryParam(true, "start"), /Excel date|ISO timestamp|relative/);
  assert.throws(() => toQueryParam("last tuesday", "start"), RangeError);
});

test("interval: parses counts and units", () => {
  assert.equal(parseInterval("30s"), 30000);
  assert.equal(parseInterval("5m"), 300000);
  assert.equal(parseInterval("1h"), 3600000);
  assert.equal(parseInterval("1d"), 86400000);
  assert.equal(parseInterval("500ms"), 500);
  assert.equal(parseInterval("1.5m"), 90000);
  assert.equal(parseInterval(" 15m "), 900000);
});

test("interval: blank and zero mean raw alignment", () => {
  assert.equal(parseInterval(null), null);
  assert.equal(parseInterval(""), null);
  assert.equal(parseInterval(0), null);
});

test("interval: a bare number is seconds", () => {
  assert.equal(parseInterval(60), 60000);
});

test("interval: rejects junk and non-positive values", () => {
  assert.throws(() => parseInterval("soon"), RangeError);
  assert.throws(() => parseInterval("-5m"), RangeError);
  assert.throws(() => parseInterval("0s"), RangeError);
});

test("toQueryParam: a clamped expression means the user's today, resolved locally", async () => {
  const { resolveClampedLocal, parseRelative, toLocalIsoWithOffset } = await import("../src/lib/time.js");
  const now = new Date(2026, 9, 22, 14, 37, 12); // 14:37:12 local, a Thursday
  const at = (expr) => new Date(resolveClampedLocal(parseRelative(expr), now));
  const wall = (d) => [d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()];

  assert.deepEqual(wall(at("d")), [2026, 10, 22, 0, 0, 0]);
  assert.deepEqual(wall(at("d+12h")), [2026, 10, 22, 12, 0, 0], "12:00 today on this machine's clock");
  assert.deepEqual(wall(at("d-1d")), [2026, 10, 21, 0, 0, 0]);
  assert.deepEqual(wall(at("h")), [2026, 10, 22, 14, 0, 0]);
  assert.deepEqual(wall(at("m-5m")), [2026, 10, 22, 14, 32, 0]);
  assert.deepEqual(wall(at("w")), [2026, 10, 18, 0, 0, 0], "weeks start on Sunday");
  assert.deepEqual(wall(at("M")), [2026, 10, 1, 0, 0, 0]);
  assert.deepEqual(wall(at("M-1M")), [2026, 9, 1, 0, 0, 0]);
  assert.deepEqual(wall(at("y")), [2026, 1, 1, 0, 0, 0]);

  // Sent as an absolute start carrying the offset, so the Historian cannot reinterpret it.
  const q = toQueryParam("d+12h", "start", { now: () => now });
  assert.equal(q.param, "start");
  assert.equal(q.value, toLocalIsoWithOffset(at("d+12h").getTime()));
  assert.match(q.value, /^2026-10-22T12:00:00[+-]\d{2}:\d{2}$/);
});

test("toQueryParam: a pure offset is still left to the server", () => {
  assert.deepEqual(toQueryParam("-24h", "start"), { param: "relativeStart", value: "-24h" });
  assert.deepEqual(toQueryParam("-90m", "end"), { param: "relativeEnd", value: "-90m" });
});

test("timestampCell: a real date that reads mm/dd/yy hh:mm:ss.000", async () => {
  const { timestampCell, TIMESTAMP_FORMAT } = await import("../src/lib/time.js");
  assert.equal(TIMESTAMP_FORMAT, "mm/dd/yy hh:mm:ss.000");
  const serial = msToExcelSerial(Date.UTC(2026, 9, 22, 12, 18, 36, 668), { utc: true });
  assert.deepEqual(timestampCell(serial), { type: "FormattedNumber", basicValue: serial, numberFormat: TIMESTAMP_FORMAT });
  assert.equal(timestampCell(""), "", "non-numbers pass through");
});
