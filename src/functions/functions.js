/* global CustomFunctions */
// Worksheet custom functions. Each one is a thin adapter: parse Excel's
// arguments with lib/args.js, call the tested lib/*.js logic, and turn any
// thrown error into the one-line message Excel shows on #VALUE!.
//
// Every function id below MUST have a matching entry in functions.json with
// the same id, or Excel will not register it.
//
// Runs in the custom-functions runtime (functions.html), separate from the
// task pane, so each function loads the saved settings before using them.

import { getClient, requireConfigured, getSettings, loadSettings } from "../runtime/session.js";
import { flattenTagList, requireScalar, rejectIfError, blankNulls } from "../lib/args.js";
import { toQueryParam, parseInterval, parseApiTimestamp, msToExcelSerial, timestampCell } from "../lib/time.js";
import { buildGrid, normalizeSeries, normalizeAggregate } from "../lib/align.js";
import {
  buildSampleGrid,
  parseSampleInterval,
  sampleBoundaries,
  seriesNeedingSeed,
  SAMPLE_MODES,
  DEFAULT_LOOKBACK_MS,
} from "../lib/sample.js";
import { predicateFor } from "../lib/quality.js";
import { formatValue } from "../lib/format.js";
import { parseColumns } from "../lib/columns.js";
import { toCellMessage } from "../lib/errors.js";

/**
 * How a quality issue shows in the sheet: the carried value, in red.
 *
 * A custom function cannot format the cell it lives in, but it can return a
 * FormattedNumber -- a real number that carries its own number format -- so the
 * cell stays numeric for charts and further formulas while reading red. Text
 * and blank cells have no number format to carry and are returned unchanged.
 */
export const QUALITY_ISSUE_FORMAT = "[Red]General";

function flagCell(value) {
  if (typeof value !== "number" || !isFinite(value)) return value;
  return { type: "FormattedNumber", basicValue: value, numberFormat: QUALITY_ISSUE_FORMAT };
}

/** % Good as a real fraction that reads as a percentage (0.75 shows 75%). */
function percentCell(fraction) {
  return { type: "FormattedNumber", basicValue: fraction, numberFormat: "0%" };
}

function parseNoValue(value) {
  const v = (value ?? "").toString().trim().toLowerCase();
  if (v === "" || v === "flag" || v === "keep") return "flag";
  if (v === "drop") return "drop";
  throw new RangeError(`novalue must be "flag" (keep the row, carry the value, show it red) or "drop".`);
}

function parseLookback(value) {
  return value === undefined || value === null || value === "" ? DEFAULT_LOOKBACK_MS : parseSampleInterval(value);
}

/**
 * The carried-in value for every tag, when the response cannot supply it.
 *
 * The Historian returns, as the first point of each series, the last sample
 * before the window -- normally all that carry-forward needs. It is missing for
 * a tag with no earlier history, and unusable when the quality filter rejects
 * it (a quality-64 status point carries no value at all). For just those tags,
 * one bounded read over [start - lookback, start] finds the last good value.
 * `firstNeededAt(startMs)` is the first instant a value is needed: the window
 * start for TB.HISTORY, the first clock mark for TB.SAMPLE.
 */
async function readSeedIfNeeded(client, ds, tagNames, response, qualityPredicate, lookbackMs, firstNeededAt) {
  if (!response?.s) return null;
  const startMs = parseApiTimestamp(response.s);
  const series = (response.tl ?? []).map((e) => normalizeSeries(e, { qualityPredicate }));
  const returned = new Set(series.map((s) => s.name));
  const missing = [...new Set([...seriesNeedingSeed(series, firstNeededAt(startMs)), ...tagNames.filter((n) => !returned.has(n))])];
  if (!missing.length) return null;
  return client.readData(
    ds,
    missing,
    { param: "start", value: new Date(startMs - lookbackMs).toISOString() },
    { param: "end", value: new Date(startMs).toISOString() },
  );
}

/**
 * Read time-aligned history for one or more tags.
 *
 * @customfunction HISTORY
 * @param {any} dataset Dataset name.
 * @param {any[][]} tags One tag name, a range of tag names, or an array such as {"a","b"}. Declared as a matrix: as a scalar, Excel would call the function once per tag and return #CALC!.
 * @param {any} [start] Start of the window: an Excel date, an ISO timestamp, or a relative expression like "-24h". Blank means "just now".
 * @param {any} [end] End of the window. Blank means "now".
 * @param {string} [interval] Bucket width such as "30s", "5m", "1h". Blank means every raw sample, aligned by last-known-value.
 * @param {string} [aggregate] avg (time-weighted average; also "twa"), min, max, first, last, count, sum, range. Only used with interval. Defaults to avg. Every slice starts from the value held at its start.
 * @param {string} [quality] One of any, good, notbad. Defaults to any.
 * @param {any} [lookback] How far before start to search for the carried-in value if the Historian's prior-value point is missing or rejected. Defaults to "7d".
 * @param {string} [novalue] Points with no value (quality notes such as 64, "collector heartbeat lost"): "flag" (default) keeps the row, carries the last value through it and shows it red; "drop" leaves those rows out.
 * @param {any} [columns] Extra columns after each tag: any of "quality" (worst OPC code in the row's range), "status" (the same in words), "samples" (points logged), "good%" (share that were Good), e.g. "quality,samples,good%".
 * @returns {any[][]} A table: Timestamp column, then one column per tag. With a start, the first row is the start itself, carrying each tag's last known value. Values whose latest point is not Good quality show red.
 */
async function HISTORY(dataset, tags, start, end, interval, aggregate, quality, lookback, novalue, columns) {
  try {
    await loadSettings();
    requireConfigured();
    const ds = requireScalar(rejectIfError(dataset, "Dataset"), "Dataset");
    const tagNames = flattenTagList(rejectIfError(tags, "Tags"));
    if (tagNames.length === 0) {
      throw new RangeError("At least one tag name is required.");
    }

    const startParam = toQueryParam(rejectIfError(start, "Start"), "start");
    const endParam = toQueryParam(rejectIfError(end, "End"), "end");
    const intervalMs = parseInterval(rejectIfError(interval, "Interval"));
    const aggKind = normalizeAggregate(rejectIfError(aggregate, "Aggregate"));
    const qualityPredicate = predicateFor(quality);
    const lookbackMs = parseLookback(rejectIfError(lookback, "Lookback"));
    const noValueMode = parseNoValue(rejectIfError(novalue, "novalue"));
    const extraColumns = parseColumns(rejectIfError(columns, "Columns"));

    const client = getClient();
    const response = await client.readData(ds, tagNames, startParam, endParam);
    // Carry-forward needs the value in force at the start; only fetch further
    // back for the tags whose own boundary point is missing or rejected.
    const seedResponse = startParam
      ? await readSeedIfNeeded(client, ds, tagNames, response, qualityPredicate, lookbackMs, (startMs) => startMs)
      : null;

    const { maxRows } = getSettings();
    return blankNulls(buildGrid(response, {
      interval: intervalMs,
      aggregate: aggKind,
      qualityPredicate,
      tagOrder: tagNames,
      maxRows,
      seedResponse,
      // No start means "latest value only": nothing to carry forward to.
      boundary: startParam ? "start" : "drop",
      novalue: noValueMode,
      flagCell,
      columns: extraColumns,
      percentCell,
      timeCell: timestampCell,
    }));
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, toCellMessage(err));
  }
}

/**
 * Sample-and-hold ("last value") for state tags: one row per clock-aligned
 * boundary, each holding the last raw value at or before that instant.
 *
 * @customfunction SAMPLE
 * @param {any} dataset Dataset name.
 * @param {any[][]} tags One tag name, a range of tag names, or an array such as {"a","b"}. Declared as a matrix: as a scalar, Excel would call the function once per tag and return #CALC!.
 * @param {any} start Start of the window: an Excel date, an ISO timestamp, or a relative expression like "-24h".
 * @param {any} [end] End of the window, inclusive. Blank means "now".
 * @param {any} [interval] Minutes between rows (5), or text with a unit ("30s", "15m", "1h", "1d"). Defaults to 5.
 * @param {string} [mode] "last" (sample-and-hold). The only mode today.
 * @param {string} [quality] One of good (default), notbad, any.
 * @param {boolean} [integers] TRUE (default) truncates numbers to integers, so 6.0 reads as 6. FALSE keeps the stored type.
 * @param {any} [lookback] How far before start to look for a seed value if the Historian's own prior-value point is missing or rejected. Defaults to "7d".
 * @param {any} [columns] Extra columns after each tag: any of "quality", "status", "samples", "good%" -- over the span since the previous row.
 * @returns {any[][]} A table: Timestamp column, then one column per tag.
 */
async function SAMPLE(dataset, tags, start, end, interval, mode, quality, integers, lookback, columns) {
  try {
    await loadSettings();
    requireConfigured();
    const ds = requireScalar(rejectIfError(dataset, "Dataset"), "Dataset");
    const tagNames = flattenTagList(rejectIfError(tags, "Tags"));
    if (tagNames.length === 0) {
      throw new RangeError("At least one tag name is required.");
    }

    const startParam = toQueryParam(rejectIfError(start, "Start"), "start");
    if (!startParam) throw new RangeError("Start is required: sampling needs a window.");
    const endParam = toQueryParam(rejectIfError(end, "End"), "end");
    const intervalMs = parseSampleInterval(rejectIfError(interval, "Interval"));
    const modeKind = (mode ?? "last").toString().trim().toLowerCase() || "last";
    if (!SAMPLE_MODES.includes(modeKind)) {
      throw new RangeError(`Unknown mode "${mode}". Use one of ${SAMPLE_MODES.join(", ")}.`);
    }
    const qualityPredicate = predicateFor(quality === undefined || quality === null || quality === "" ? "good" : quality);
    const castInt = integers === undefined || integers === null || integers === "" ? true : Boolean(integers);
    const lookbackMs = parseLookback(rejectIfError(lookback, "Lookback"));
    const extraColumns = parseColumns(rejectIfError(columns, "Columns"));

    const client = getClient();
    const response = await client.readData(ds, tagNames, startParam, endParam);
    const { maxRows } = getSettings();

    // The leading boundary point the API returns is the seed for the first
    // row; only when it is missing or rejected is a second read made.
    const seedResponse = await readSeedIfNeeded(
      client, ds, tagNames, response, qualityPredicate, lookbackMs,
      (startMs) => sampleBoundaries(startMs, startMs + intervalMs, intervalMs)[0],
    );

    return blankNulls(buildSampleGrid(response, {
      intervalMs,
      qualityPredicate,
      castInt,
      tagOrder: tagNames,
      seedResponse,
      maxRows,
      flagCell,
      columns: extraColumns,
      percentCell,
      timeCell: timestampCell,
    }));
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, toCellMessage(err));
  }
}

/**
 * The current (or last known) value of a single tag.
 *
 * @customfunction VALUE
 * @param {any} dataset Dataset name.
 * @param {any} tag Tag name.
 * @param {string} [mode] "raw" (default) returns the number/string as stored. "display" returns the Historian's formatted text, with units and enum labels resolved.
 * @returns {any} The current value.
 */
async function VALUE(dataset, tag, mode) {
  try {
    await loadSettings();
    requireConfigured();
    const ds = requireScalar(rejectIfError(dataset, "Dataset"), "Dataset");
    const tagName = requireScalar(rejectIfError(tag, "Tag"), "Tag");
    const wantDisplay = (mode ?? "raw").toString().trim().toLowerCase() === "display";

    const client = getClient();
    const response = await client.readData(ds, [tagName], null, null);
    const envelope = (response.tl ?? []).find((e) => e?.t?.n === tagName);
    if (!envelope || !Array.isArray(envelope.d) || envelope.d.length === 0) {
      throw new RangeError(`No value is available for "${tagName}" in dataset "${ds}".`);
    }
    const latest = envelope.d[envelope.d.length - 1];

    if (!wantDisplay) return latest.v ?? null;

    const tags = await client.listTags(ds, { contains: tagName });
    const full = tags.find((t) => t.name === tagName) ?? { name: tagName };
    return formatValue(latest.v, full);
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, toCellMessage(err));
  }
}

/**
 * List the datasets on this Historian.
 *
 * @customfunction DATASETS
 * @returns {any[][]} Name, tag count, size (GB), data age (days).
 */
async function DATASETS() {
  try {
    await loadSettings();
    requireConfigured();
    const client = getClient();
    const datasets = await client.listDatasets();
    const rows = datasets.map((d) => [d.name, d.tagCount, d.sizeGb, d.dataAgeDays]);
    return blankNulls([["Dataset", "Tags", "Size (GB)", "Age (days)"], ...rows]);
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, toCellMessage(err));
  }
}

/**
 * Search tags in a dataset using the Historian's filter language
 * (name:, description:, datatype:, type:, format:, AND/OR/NOT, wildcards).
 *
 * @customfunction TAGS
 * @param {any} dataset Dataset name.
 * @param {string} [query] Filter query, e.g. "name:Pump_* AND datatype:Double".
 * @returns {any[][]} Name, Description, Data type, Unit, Current value, Current time.
 */
async function TAGS(dataset, query) {
  try {
    await loadSettings();
    requireConfigured();
    const ds = requireScalar(rejectIfError(dataset, "Dataset"), "Dataset");
    const client = getClient();
    const tags = await client.listTags(ds, { query: query || null });
    const rows = tags.map((t) => [
      t.name,
      t.description,
      t.dataType,
      formatUnitCell(t.units),
      t.last ? t.last.value : null,
      t.last && t.last.time ? tagTimeCell(t.last.time) : null,
    ]);
    return blankNulls([["Name", "Description", "Data Type", "Unit", "Current Value", "Current Time"], ...rows]);
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, toCellMessage(err));
  }
}

/** A tag's last-value time (ISO text from the API) as a formatted Excel date. */
function tagTimeCell(iso) {
  try {
    return timestampCell(msToExcelSerial(parseApiTimestamp(iso)));
  } catch {
    return iso;
  }
}

function formatUnitCell(units) {
  if (!units || typeof units !== "object") return "";
  const keys = Object.keys(units);
  if (keys.length === 1) return String(units[keys[0]] ?? "");
  return keys.length ? "(enumerated)" : "";
}

CustomFunctions.associate("HISTORY", HISTORY);
CustomFunctions.associate("SAMPLE", SAMPLE);
CustomFunctions.associate("VALUE", VALUE);
CustomFunctions.associate("DATASETS", DATASETS);
CustomFunctions.associate("TAGS", TAGS);
