import { test } from "node:test";
import assert from "node:assert/strict";

import { pickerValueToIso, isoToPickerValue, addTags, removeTag } from "../src/lib/picker.js";
import { toQueryParam, parseApiTimestamp } from "../src/lib/time.js";

const eastern = () => 240; // getTimezoneOffset() for UTC-04:00
const india = () => -330; // UTC+05:30

test("pickerValueToIso: writes the workstation offset, never a bare local time", () => {
  assert.equal(pickerValueToIso("2026-10-22T07:00", eastern), "2026-10-22T07:00:00-04:00");
  assert.equal(pickerValueToIso("2026-10-22T07:00:30", eastern), "2026-10-22T07:00:30-04:00");
  assert.equal(pickerValueToIso("2026-10-22T07:00", india), "2026-10-22T07:00:00+05:30");
  assert.equal(pickerValueToIso("2026-10-22T07:00", () => 0), "2026-10-22T07:00:00+00:00");
});

test("pickerValueToIso: the result names the instant the user picked", () => {
  const iso = pickerValueToIso("2026-10-22T07:00", eastern);
  assert.equal(new Date(parseApiTimestamp(iso)).toISOString(), "2026-10-22T11:00:00.000Z");
  // and goes to the API as an absolute start, not a relative expression
  assert.deepEqual(toQueryParam(iso, "start"), { param: "start", value: iso });
});

test("pickerValueToIso: the offset is the picked day's, across DST", () => {
  // US Eastern: EDT (-4) until 2026-11-01, EST (-5) after.
  const offsetFor = (d) => (d < new Date(2026, 10, 1, 2) ? 240 : 300);
  assert.equal(pickerValueToIso("2026-10-31T09:00", offsetFor), "2026-10-31T09:00:00-04:00");
  assert.equal(pickerValueToIso("2026-11-02T09:00", offsetFor), "2026-11-02T09:00:00-05:00");
});

test("pickerValueToIso: blank is blank; junk is refused", () => {
  assert.equal(pickerValueToIso(""), "");
  assert.throws(() => pickerValueToIso("yesterday"), RangeError);
});

test("isoToPickerValue: opens on the timestamp already in the box, else on now", () => {
  const now = new Date(2026, 9, 22, 14, 7, 59);
  const typed = pickerValueToIso("2026-10-20T17:30");
  assert.equal(isoToPickerValue(typed, now), "2026-10-20T17:30");
  assert.equal(isoToPickerValue("-24h", now), "2026-10-22T14:07");
  assert.equal(isoToPickerValue("", now), "2026-10-22T14:07");
});

test("addTags / removeTag: no duplicates, order kept", () => {
  let list = addTags([], ["FL001.State", "FL002.State"]);
  list = addTags(list, ["FL002.State", " FL003.State ", ""]);
  assert.deepEqual(list, ["FL001.State", "FL002.State", "FL003.State"]);
  assert.deepEqual(removeTag(list, "FL002.State"), ["FL001.State", "FL003.State"]);
});
