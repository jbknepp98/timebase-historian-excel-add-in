import { test } from "node:test";
import assert from "node:assert/strict";

import { flattenTagList, requireScalar, isExcelError, rejectIfError } from "../src/lib/args.js";

test("flattenTagList: a bare scalar becomes a one-item list", () => {
  assert.deepEqual(flattenTagList("131-TT-001.PV"), ["131-TT-001.PV"]);
});

test("flattenTagList: a one-column range (array of one-cell rows)", () => {
  assert.deepEqual(flattenTagList([["A"], ["B"], ["C"]]), ["A", "B", "C"]);
});

test("flattenTagList: a one-row range (single row of cells)", () => {
  assert.deepEqual(flattenTagList([["A", "B", "C"]]), ["A", "B", "C"]);
});

test("flattenTagList: a full 2-D block reads row-major", () => {
  assert.deepEqual(flattenTagList([["A", "B"], ["C", "D"]]), ["A", "B", "C", "D"]);
});

test("flattenTagList: blank cells (empty string) are dropped, not kept as gaps", () => {
  assert.deepEqual(flattenTagList([["A"], [""], ["B"]]), ["A", "B"]);
});

test("flattenTagList: null/undefined cells are dropped the same way", () => {
  assert.deepEqual(flattenTagList([["A"], [null], [undefined], ["B"]]), ["A", "B"]);
});

test("flattenTagList: numeric cells are stringified", () => {
  assert.deepEqual(flattenTagList([[1], [2.5]]), ["1", "2.5"]);
});

test("flattenTagList: whitespace-only cells are dropped", () => {
  assert.deepEqual(flattenTagList([["A"], ["   "], ["B"]]), ["A", "B"]);
});

test("flattenTagList: no argument at all is an empty list, not an error", () => {
  assert.deepEqual(flattenTagList(null), []);
  assert.deepEqual(flattenTagList(undefined), []);
  assert.deepEqual(flattenTagList(""), []);
});

test("requireScalar: a single value passes through", () => {
  assert.equal(requireScalar("MyDataset", "Dataset"), "MyDataset");
  assert.equal(requireScalar([["MyDataset"]], "Dataset"), "MyDataset");
});

test("requireScalar: empty is rejected with the field name in the message", () => {
  assert.throws(() => requireScalar("", "Dataset"), /Dataset/);
  assert.throws(() => requireScalar(null, "Dataset"), /Dataset/);
});

test("requireScalar: a multi-cell range is rejected rather than silently truncated", () => {
  assert.throws(() => requireScalar([["A"], ["B"]], "Dataset"), /single cell/);
});

test("isExcelError / rejectIfError: recognises Excel's error-value shape", () => {
  const err = { type: "Error", errorType: "#N/A" };
  assert.equal(isExcelError(err), true);
  assert.equal(isExcelError("just a string"), false);
  assert.equal(isExcelError(null), false);
  assert.throws(() => rejectIfError(err, "Start time"), /Start time/);
  assert.equal(rejectIfError("fine", "Start time"), "fine");
});

test("blankNulls: null and undefined become empty strings; 0 and false stay", async () => {
  const { blankNulls } = await import("../src/lib/args.js");
  assert.deepEqual(blankNulls([["Timestamp", "A"], [1, null], [2, undefined], [3, 0], [4, false]]), [
    ["Timestamp", "A"],
    [1, ""],
    [2, ""],
    [3, 0],
    [4, false],
  ]);
});
