import { test } from "node:test";
import assert from "node:assert/strict";

import {
  classifyUnits,
  unitLabel,
  decimalsFromFormat,
  formatValue,
  columnHeading,
  isEnumerated,
} from "../src/lib/format.js";

test("classifyUnits: single positive key is a suffix unit", () => {
  assert.equal(classifyUnits({ 1: "kW" }), "suffix");
});

test("classifyUnits: single negative key is a prefix unit", () => {
  assert.equal(classifyUnits({ "-1": "US$" }), "prefix");
});

test("classifyUnits: exactly the keys 0 and 1 is boolean, not a two-member enum", () => {
  assert.equal(classifyUnits({ 0: "Stopped", 1: "Running" }), "boolean");
});

test("classifyUnits: three or more keys is an enumeration", () => {
  assert.equal(classifyUnits({ 1: "Apple", 2: "Orange", 4: "Lemon" }), "enum");
});

test("classifyUnits: two keys that are not {0,1} is an enum, not boolean", () => {
  assert.equal(classifyUnits({ 1: "Idle", 2: "Running" }), "enum");
});

test("classifyUnits: empty, missing or non-numeric-key maps are none", () => {
  assert.equal(classifyUnits(null), "none");
  assert.equal(classifyUnits({}), "none");
  assert.equal(classifyUnits({ x: "?" }), "none");
});

test("unitLabel: only scalar (prefix/suffix) maps produce a label", () => {
  assert.equal(unitLabel({ 1: "kW" }), "kW");
  assert.equal(unitLabel({ "-1": "US$" }), "US$");
  assert.equal(unitLabel({ 0: "Stopped", 1: "Running" }), "", "boolean has no scalar unit");
  assert.equal(unitLabel({ 1: "A", 2: "B", 3: "C" }), "", "enum has no scalar unit");
  assert.equal(unitLabel(null), "");
});

test("decimalsFromFormat: .NET standard numeric formats", () => {
  assert.equal(decimalsFromFormat("F2"), 2);
  assert.equal(decimalsFromFormat("f0"), 0);
  assert.equal(decimalsFromFormat("N1"), 1);
});

test("decimalsFromFormat: custom picture formats", () => {
  assert.equal(decimalsFromFormat("0.00"), 2);
  assert.equal(decimalsFromFormat("#,##0.0"), 1);
  assert.equal(decimalsFromFormat("0.0000"), 4);
  assert.equal(decimalsFromFormat("#,##0"), 0);
});

test("decimalsFromFormat: unknown or absent format is null, not a guess", () => {
  assert.equal(decimalsFromFormat(""), null);
  assert.equal(decimalsFromFormat(undefined), null);
  assert.equal(decimalsFromFormat("yyyy-MM-dd"), null);
});

test("formatValue: suffix unit follows the number with a space", () => {
  assert.equal(formatValue(57.957364, { format: "0.00", units: { 1: "°F" } }), "57.96 °F");
});

test("formatValue: prefix unit has no space, per the currency example", () => {
  assert.equal(formatValue(42, { format: "0", units: { "-1": "US$" } }), "US$42");
});

test("formatValue: boolean tag maps 1/0 through the label pair", () => {
  const tag = { units: { 0: "Stopped", 1: "Running" } };
  assert.equal(formatValue(1, tag), "Running");
  assert.equal(formatValue(0, tag), "Stopped");
  assert.equal(formatValue(true, tag), "Running", "boolean DataType values also map");
  assert.equal(formatValue(false, tag), "Stopped");
});

test("formatValue: enum tag maps the numeric code to its label", () => {
  const tag = { units: { 1: "Apple", 2: "Orange", 4: "Lemon" } };
  assert.equal(formatValue(2, tag), "Orange");
});

test("formatValue: an enum code with no matching label shows the raw code, not blank", () => {
  const tag = { units: { 1: "Apple", 2: "Orange", 4: "Lemon" } };
  assert.equal(formatValue(99, tag), "99");
});

test("formatValue: null and undefined render as empty, not the string 'null'", () => {
  assert.equal(formatValue(null, {}), "");
  assert.equal(formatValue(undefined, {}), "");
});

test("formatValue: a string-valued tag passes through", () => {
  assert.equal(formatValue("Arthur Dent", {}), "Arthur Dent");
});

test("columnHeading: appends the unit only for scalar tags", () => {
  assert.equal(columnHeading({ name: "TT-001", units: { 1: "°F" } }), "TT-001 (°F)");
  assert.equal(columnHeading({ name: "FL001.State", units: { 0: "Idle", 1: "Running" } }), "FL001.State");
  assert.equal(columnHeading({ name: "Plain" }), "Plain");
});

test("isEnumerated: true for boolean and enum tags, false for scalars", () => {
  assert.equal(isEnumerated({ units: { 0: "Stopped", 1: "Running" } }), true);
  assert.equal(isEnumerated({ units: { 1: "Apple", 2: "Orange", 4: "Lemon" } }), true);
  assert.equal(isEnumerated({ units: { 1: "kW" } }), false);
  assert.equal(isEnumerated({}), false);
});
