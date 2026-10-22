// Sample-and-hold ("last value", LOCF) onto clock-aligned boundaries.
//
// This is the right read for a discrete or state tag: an integer enum stored on
// change. Averaging such a tag invents states that do not exist (4.6 is not a
// valid inverter status), and interpolating it pretends a state ramps from 4 to
// 5. At each boundary T this returns the last raw value whose timestamp is
// <= T, and nothing else.
//
// It differs from TB.HISTORY's interval + "last" aggregate in two ways that
// matter for a report:
//
//   - Rows land on clock marks (07:00, 07:05, ...), not on offsets from
//     whenever the query window happened to start (07:03:12, 07:08:12, ...).
//   - Each row is a snapshot AT its timestamp. A bucket's "last" is the value at
//     the END of the bucket labelled with its START, which is five minutes
//     ahead of what the row's timestamp says.
//
// A snapshot misses a state that starts and ends between two boundaries (a
// one-second fault at 07:17:27 never appears in a 5-minute table). That is
// inherent to sampling, not a bug; the task pane says so next to the option.
//
// The Historian's /data route has no interpolation, aggregation, or interval
// parameter (checked against its OpenAPI v1 document), so this is done here,
// on top of the raw read. The route does return a leading boundary point --
// the last sample before the window -- which is what seeds the first row.

import { normalizeSeries } from "./align.js";
import { rangeStats, columnHeaders, columnCells } from "./columns.js";
import { parseApiTimestamp, msToExcelSerial } from "./time.js";

const MINUTE_MS = 60000;
const HOUR_MS = 60 * MINUTE_MS;

export const SAMPLE_MODES = ["last"];
export const DEFAULT_SAMPLE_INTERVAL_MS = 5 * MINUTE_MS;
export const DEFAULT_LOOKBACK_MS = 7 * 24 * HOUR_MS;

/** Offset of local wall-clock time from UTC at instant `ms`, in milliseconds. */
function localOffsetMs(ms) {
  return -new Date(ms).getTimezoneOffset() * MINUTE_MS;
}

/**
 * Parse TB.SAMPLE's interval argument.
 *
 * A bare number is MINUTES here (the spec's `intervalMinutes`), unlike
 * TB.HISTORY where a bare number is seconds -- a state report is thought of in
 * minutes, and "5" meaning five seconds would be a nasty surprise. Text takes a
 * unit: "30s", "5m", "1h", "1d".
 */
export function parseSampleInterval(input) {
  if (input === null || input === undefined || input === "") return DEFAULT_SAMPLE_INTERVAL_MS;
  if (typeof input === "number") {
    if (!isFinite(input) || input <= 0) throw new RangeError(`Interval must be a positive number of minutes, got ${input}.`);
    return Math.round(input * MINUTE_MS);
  }
  const s = String(input).trim();
  if (/^\d+(\.\d+)?$/.test(s)) return parseSampleInterval(Number(s));
  const m = /^(\d+(?:\.\d+)?)\s*(s|m|h|d)$/.exec(s);
  if (!m) {
    throw new RangeError(`Not an interval: "${input}". Use minutes (5) or a count and unit ("30s", "5m", "1h", "1d").`);
  }
  const ms = Math.round(Number(m[1]) * { s: 1000, m: MINUTE_MS, h: HOUR_MS, d: 24 * HOUR_MS }[m[2]]);
  if (ms <= 0) throw new RangeError(`Interval must be greater than zero: "${input}"`);
  return ms;
}

/**
 * Every clock-aligned boundary from ceil(start) to end, inclusive.
 *
 * Under an hour: multiples of the interval since the Unix epoch. For 1, 5, 10,
 * 15 and 30 minutes that is the same as local clock marks in every whole- and
 * half-hour time zone.
 *
 * An hour or more: aligned on the LOCAL wall clock, so an hourly row lands on
 * the local hour in a :30 zone and a daily row lands on local midnight rather
 * than UTC midnight. Boundaries are stepped in wall-clock time and converted
 * back, so a daily row stays on midnight across a DST change (that day is 23
 * or 25 hours long).
 *
 * `offsetMs(ms)` is injectable for tests; it defaults to this machine's zone,
 * which is the zone every other timestamp this add-in writes is shown in.
 */
export function sampleBoundaries(startMs, endMs, intervalMs, { offsetMs = localOffsetMs, maxCount = Infinity } = {}) {
  if (!(intervalMs > 0)) throw new RangeError("The interval must be greater than zero.");
  if (!(endMs >= startMs)) throw new RangeError("The end of the window must not be before its start.");

  const out = [];
  const push = (b) => {
    if (out.length >= maxCount) {
      throw new RangeError(
        `This query samples to more than ${maxCount.toLocaleString("en-US")} rows. Use a longer interval or a shorter time range.`,
      );
    }
    out.push(b);
  };

  if (intervalMs < HOUR_MS) {
    for (let b = Math.ceil(startMs / intervalMs) * intervalMs; b <= endMs; b += intervalMs) push(b);
    return out;
  }

  // Wall-clock -> UTC: one correction pass handles every instant except the
  // hour skipped by a spring-forward change, which the strictly-increasing
  // guard below drops rather than emitting a duplicate row.
  const toUtc = (wall) => {
    const guess = wall - offsetMs(wall);
    return wall - offsetMs(guess);
  };
  const startWall = startMs + offsetMs(startMs);
  let wall = Math.ceil(startWall / intervalMs) * intervalMs;
  let b = toUtc(wall);
  // A boundary computed in wall time can fall a DST hour before startMs.
  while (b < startMs) {
    wall += intervalMs;
    b = toUtc(wall);
  }
  for (; b <= endMs; wall += intervalMs, b = toUtc(wall)) {
    if (out.length && b <= out[out.length - 1]) continue;
    push(b);
  }
  return out;
}

/** 6.0 -> 6 for a state code. Text and booleans are left exactly as stored. */
function castValue(v, castInt) {
  if (!castInt || typeof v !== "number" || !isFinite(v)) return v;
  return Math.trunc(v);
}

/**
 * The reference algorithm from the feature spec: one pass, O(points + rows).
 *
 * `points` is [{ms, v}] already quality-filtered, and must include the seed
 * (the last point before the first boundary) if one exists. Returns one value
 * per boundary; null where the tag had no value yet, which Excel shows as a
 * blank cell rather than a misleading 0.
 */
export function lastValueSamples(points, boundaries, { castInt = true } = {}) {
  return sampleWithFlags(points, boundaries, { castInt }).values;
}

/**
 * The same walk, also reporting per boundary whether the latest point at or
 * before it was a quality issue. A point with no value (a quality-64 note)
 * never becomes the sampled value -- the last real value carries through it --
 * but it does flag the sample until the next good point.
 */
export function sampleWithFlags(points, boundaries, { castInt = true, rangeStartMs = null } = {}) {
  const pts = points.slice().sort((a, b) => a.ms - b.ms);
  const values = new Array(boundaries.length);
  const flags = new Array(boundaries.length);
  const stats = new Array(boundaries.length);
  let i = 0;
  let last = null;
  let issue = false;
  let inForce = null;

  // Row k's range for the quality columns is (previous mark, this mark],
  // starting from the point in force at the previous mark; the first row's
  // range starts at the window start.
  const consume = (p) => {
    const novalue = p.novalue ?? (p.v === undefined || p.v === null);
    if (!novalue) last = p.v;
    issue = !!p.issue;
    inForce = p;
  };
  if (rangeStartMs !== null) {
    while (i < pts.length && pts[i].ms <= rangeStartMs) consume(pts[i++]);
  }

  for (let k = 0; k < boundaries.length; k++) {
    const b = boundaries[k];
    const st = rangeStats(inForce);
    while (i < pts.length && pts[i].ms <= b) {
      st.add(pts[i]);
      consume(pts[i++]);
    }
    values[k] = last === null ? null : castValue(last, castInt);
    flags[k] = issue;
    stats[k] = st;
  }
  return { values, flags, stats };
}

/**
 * Series that still need a seed: no surviving point at or before the first
 * boundary. Usually empty, because the API's leading boundary point is the
 * seed. It is non-empty when that point was rejected by the quality filter, or
 * the tag has no history before the window at all.
 */
export function seriesNeedingSeed(seriesList, firstBoundaryMs) {
  if (firstBoundaryMs === undefined || firstBoundaryMs === null) return [];
  // Only a point with a value can seed; a quality-64 note before the window
  // (which the Historian may return as the boundary point) does not count.
  return seriesList
    .filter((s) => !s.points.some((p) => p.ms <= firstBoundaryMs && !p.novalue && p.v !== null && p.v !== undefined))
    .map((s) => s.name);
}

/**
 * A TVQsResponse in, a 2-D array for Excel out.
 *
 * Options:
 *   intervalMs        boundary spacing (default 5 minutes)
 *   qualityPredicate  per-sample filter; rejected samples never become a held value
 *   castInt           truncate numeric values to integers (default true)
 *   tagOrder          column order
 *   seedResponse      an optional second TVQsResponse covering a lookback
 *                     window, merged in for series that lacked a seed
 *   maxRows           refuse rather than truncate past this
 *   offsetMs          time-zone hook for tests
 *   utc               write UTC wall-clock instead of local
 */
export function buildSampleGrid(response, options = {}) {
  const {
    intervalMs = DEFAULT_SAMPLE_INTERVAL_MS,
    qualityPredicate = () => true,
    castInt = true,
    tagOrder = null,
    seedResponse = null,
    maxRows = 100000,
    offsetMs = localOffsetMs,
    utc = false,
    flagCell = (v) => v,
    columns = [],
    percentCell = (f) => f,
    timeCell = (serial) => serial,
  } = options;

  if (response?.s === undefined || response?.s === null || response?.e === undefined || response?.e === null) {
    throw new Error("Sampling needs the query window, which the response did not report in its s and e fields.");
  }
  const startMs = parseApiTimestamp(response.s);
  const endMs = parseApiTimestamp(response.e);
  const boundaries = sampleBoundaries(startMs, endMs, intervalMs, { offsetMs, maxCount: maxRows });

  const envelopes = Array.isArray(response.tl) ? response.tl : [];
  let seriesList = envelopes.map((e) => normalizeSeries(e, { qualityPredicate }));
  if (tagOrder && tagOrder.length) {
    const byName = new Map(seriesList.map((s) => [s.name, s]));
    seriesList = tagOrder.map((name) => byName.get(name) ?? { name, tag: { n: name }, points: [] });
  }

  if (seedResponse && Array.isArray(seedResponse.tl)) {
    const seeds = new Map(
      seedResponse.tl.map((e) => normalizeSeries(e, { qualityPredicate })).map((s) => [s.name, s.points]),
    );
    seriesList = seriesList.map((s) => {
      const extra = (seeds.get(s.name) ?? []).filter((p) => p.ms < startMs);
      return extra.length ? { ...s, points: [...extra, ...s.points] } : s;
    });
  }

  const sampled = seriesList.map((s) => sampleWithFlags(s.points, boundaries, { castInt, rangeStartMs: startMs }));

  const grid = [["Timestamp", ...seriesList.flatMap((s) => [s.name, ...columnHeaders(s.name, columns)])]];
  for (let k = 0; k < boundaries.length; k++) {
    grid.push([
      timeCell(msToExcelSerial(boundaries[k], { utc })),
      ...sampled.flatMap((c) => [
        c.flags[k] ? flagCell(c.values[k]) : c.values[k],
        ...columnCells(c.stats[k], columns, { percentCell }),
      ]),
    ]);
  }
  return grid;
}
