import { test } from "node:test";
import assert from "node:assert/strict";

import {
  quote,
  tagsArg,
  buildHistoryFormula,
  buildValueFormula,
  buildDatasetsFormula,
  buildTagsFormula,
} from "../src/lib/formula.js";

test("quote: doubles interior quotes per Excel string-literal syntax", () => {
  assert.equal(quote("plain"), '"plain"');
  assert.equal(quote('Say "hi"'), '"Say ""hi"""');
});

test("tagsArg: a single tag is a plain quoted string", () => {
  assert.equal(tagsArg("131-TT-001.PV"), '"131-TT-001.PV"');
  assert.equal(tagsArg(["131-TT-001.PV"]), '"131-TT-001.PV"');
});

test("tagsArg: multiple tags become an Excel array constant", () => {
  assert.equal(tagsArg(["A", "B", "C"]), '{"A","B","C"}');
});

test("tagsArg: empty input is null (caller decides how to fail)", () => {
  assert.equal(tagsArg([]), null);
  assert.equal(tagsArg([""]), null);
});

test("buildHistoryFormula: minimal call omits every optional trailing argument", () => {
  assert.equal(
    buildHistoryFormula({ dataset: "MyDataset", tags: "131-TT-001.PV" }),
    '=TB.HISTORY("MyDataset","131-TT-001.PV")',
  );
});

test("buildHistoryFormula: multiple tags render as an array constant", () => {
  const f = buildHistoryFormula({ dataset: "D", tags: ["A", "B"] });
  assert.equal(f, '=TB.HISTORY("D",{"A","B"})');
});

test("buildHistoryFormula: a reference is passed through unquoted when flagged", () => {
  const f = buildHistoryFormula({ dataset: "D", tags: "A1:A20", tagsAreReference: true });
  assert.equal(f, "=TB.HISTORY(\"D\",A1:A20)");
});

test("buildHistoryFormula: start and end fill in without pulling in later defaults", () => {
  const f = buildHistoryFormula({ dataset: "D", tags: "T", start: "-24h" });
  assert.equal(f, '=TB.HISTORY("D","T","-24h")');
});

test("buildHistoryFormula: a later argument forces earlier defaults to render", () => {
  // aggregate is non-default, so interval (blank) must still appear as "" so
  // Excel's positional parsing lands aggregate in the right slot.
  const f = buildHistoryFormula({ dataset: "D", tags: "T", aggregate: "max" });
  assert.equal(f, '=TB.HISTORY("D","T","","","","max")');
});

test("buildHistoryFormula: quality alone still fills start, end, interval and aggregate", () => {
  // The blank fill-ins are "", not the literal word "avg": HISTORY's own
  // default coercion already treats a blank aggregate argument as avg, so an
  // empty placeholder is correct and simpler than re-deriving the default here.
  const f = buildHistoryFormula({ dataset: "D", tags: "T", quality: "good" });
  assert.equal(f, '=TB.HISTORY("D","T","","","","","good")');
});

test("buildHistoryFormula: default aggregate (avg) and default quality (any) are omitted", () => {
  const f = buildHistoryFormula({ dataset: "D", tags: "T", interval: "5m", aggregate: "avg", quality: "any" });
  assert.equal(f, '=TB.HISTORY("D","T","","","5m")');
});

test("buildHistoryFormula: a relative expression is quoted like any other string", () => {
  const f = buildHistoryFormula({ dataset: "D", tags: "T", start: "-24h", end: "now" });
  assert.equal(f, '=TB.HISTORY("D","T","-24h","now")');
});

test("buildHistoryFormula: requires a dataset and at least one tag", () => {
  assert.throws(() => buildHistoryFormula({ dataset: "", tags: "T" }), RangeError);
  assert.throws(() => buildHistoryFormula({ dataset: "D", tags: [] }), RangeError);
});

test("buildValueFormula: default mode is omitted", () => {
  assert.equal(buildValueFormula({ dataset: "D", tag: "T" }), '=TB.VALUE("D","T")');
});

test("buildValueFormula: a non-default mode is included", () => {
  assert.equal(buildValueFormula({ dataset: "D", tag: "T", mode: "display" }), '=TB.VALUE("D","T","display")');
});

test("buildDatasetsFormula: no arguments", () => {
  assert.equal(buildDatasetsFormula(), "=TB.DATASETS()");
});

test("buildTagsFormula: with and without a query", () => {
  assert.equal(buildTagsFormula({ dataset: "D" }), '=TB.TAGS("D")');
  assert.equal(buildTagsFormula({ dataset: "D", query: "name:Pump_*" }), '=TB.TAGS("D","name:Pump_*")');
});

test("a dataset name containing a quote round-trips safely", () => {
  const f = buildHistoryFormula({ dataset: 'My "Special" Dataset', tags: "T" });
  assert.equal(f, '=TB.HISTORY("My ""Special"" Dataset","T")');
});

test("buildSampleFormula: defaults are omitted", async () => {
  const { buildSampleFormula } = await import("../src/lib/formula.js");
  assert.equal(
    buildSampleFormula({ dataset: "Solar", tags: ["INV_11A01.Status"], start: "-24h" }),
    '=TB.SAMPLE("Solar","INV_11A01.Status","-24h")',
  );
  assert.equal(
    buildSampleFormula({ dataset: "Solar", tags: ["A", "B"], start: "-24h", interval: "5m", quality: "good" }),
    '=TB.SAMPLE("Solar",{"A","B"},"-24h")',
  );
});

test("buildSampleFormula: a later argument forces blanks for earlier ones", async () => {
  const { buildSampleFormula } = await import("../src/lib/formula.js");
  assert.equal(
    buildSampleFormula({ dataset: "S", tags: "A", start: "-1d", interval: "15", quality: "any" }),
    '=TB.SAMPLE("S","A","-1d","",15,"","any")',
  );
  assert.equal(
    buildSampleFormula({ dataset: "S", tags: "A", start: "-1d", interval: "1h", integers: false }),
    '=TB.SAMPLE("S","A","-1d","","1h","","",FALSE)',
  );
  assert.throws(() => buildSampleFormula({ dataset: "S", tags: "A" }), /start is required/);
});

test("tags from a worksheet range go in unquoted, so the cells drive the tag list", async () => {
  const { buildSampleFormula, buildHistoryFormula, looksLikeReference } = await import("../src/lib/formula.js");
  assert.equal(
    buildSampleFormula({ dataset: "state-example", tags: "Sheet1!A2:A10", tagsAreReference: true, start: "-24h" }),
    '=TB.SAMPLE("state-example",Sheet1!A2:A10,"-24h")',
  );
  assert.equal(
    buildHistoryFormula({ dataset: "state-example", tags: "'Tag list'!$B$2:$B$40", tagsAreReference: true, start: "-1h" }),
    "=TB.HISTORY(\"state-example\",'Tag list'!$B$2:$B$40,\"-1h\")",
  );
  for (const ok of ["A2:A10", "$A$2:$A$10", "Sheet1!A2:A10", "'Tag list'!B2:B40", "B7"]) assert.ok(looksLikeReference(ok), ok);
  for (const bad of ["FL001.State", "", "Sheet1"]) assert.ok(!looksLikeReference(bad), bad);
});

test("buildHistoryFormula: novalue 'drop' is written explicitly; 'flag' is the default", async () => {
  const { buildHistoryFormula } = await import("../src/lib/formula.js");
  assert.equal(
    buildHistoryFormula({ dataset: "S", tags: "A", start: "-24h", novalue: "drop" }),
    '=TB.HISTORY("S","A","-24h","","","","","","drop")',
  );
  assert.equal(buildHistoryFormula({ dataset: "S", tags: "A", start: "-24h", novalue: "flag" }), '=TB.HISTORY("S","A","-24h")');
});

test("companion columns are written as one list at the end", async () => {
  const { buildHistoryFormula, buildSampleFormula } = await import("../src/lib/formula.js");
  assert.equal(
    buildHistoryFormula({ dataset: "S", tags: "A", start: "-24h", columns: ["quality", "good%"] }),
    '=TB.HISTORY("S","A","-24h","","","","","","","quality,good%")',
  );
  assert.equal(
    buildSampleFormula({ dataset: "S", tags: "A", start: "-24h", columns: ["samples"] }),
    '=TB.SAMPLE("S","A","-24h","","","","","","","samples")',
  );
  assert.equal(buildHistoryFormula({ dataset: "S", tags: "A", start: "-24h", columns: [] }), '=TB.HISTORY("S","A","-24h")');
});
