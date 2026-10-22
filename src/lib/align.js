// Turning the Historian's per-tag series into a rectangular sheet.
//
// The API returns one independent series per tag: `tl[i].d` is that tag's own
// timestamps, and two tags logged by different collectors will never share a
// clock. A spreadsheet needs one timestamp column and one column per tag, so
// the series have to be aligned onto a common axis.
//
// Two modes:
//
//   raw       The axis is the union of every tag's timestamps. Each cell holds
//             the tag's last known value at or before that instant. This is
//             lossless -- every logged sample appears -- and is the right
//             default for engineering work.
//
//   interval  The axis is fixed buckets. Each cell aggregates the samples that
//             fall in the bucket. This is how a month of 1-second data fits in
//             a sheet at all.
//
// Both modes lean on the boundary point: the API returns, as the first point of
// each series, the sample immediately BEFORE the requested window. Without it
// the first rows of a stepped alignment would be blank even though the tag had
// a perfectly good value the whole time.

import { parseApiTimestamp, msToExcelSerial } from "./time.js";
import { classify } from "./quality.js";
import { rangeStats, columnHeaders, columnCells } from "./columns.js";

export const AGGREGATES = ["avg", "min", "max", "first", "last", "count", "sum", "range"];

/**
 * Accepted spellings for an aggregate. "avg" is a TIME-WEIGHTED average (each
 * value counts for as long as it was in force), so the long names say so.
 */
const AGGREGATE_ALIASES = {
  twa: "avg",
  twavg: "avg",
  timeweightedaverage: "avg",
  "time-weighted average": "avg",
  "time weighted average": "avg",
  average: "avg",
};

/** Normalise an aggregate argument: blank -> "avg", aliases resolved, unknowns refused. */
export function normalizeAggregate(input) {
  const raw = (input ?? "").toString().trim().toLowerCase();
  if (raw === "") return "avg";
  const kind = AGGREGATE_ALIASES[raw] ?? raw;
  if (!AGGREGATES.includes(kind)) {
    throw new RangeError(
      `Unknown aggregate "${input}". Use avg (time-weighted average), min, max, first, last, count, sum or range.`,
    );
  }
  return kind;
}

/**
 * Normalise one TagDataEnvelope into { name, points } with numeric timestamps,
 * dropping samples the quality filter rejects.
 *
 * Quality filtering happens here, before alignment, so a rejected sample does
 * not become a carried-forward value in a later row.
 */
export function normalizeSeries(envelope, { qualityPredicate = () => true } = {}) {
  const tag = envelope?.t ?? {};
  const name = tag.n ?? "";
  const raw = Array.isArray(envelope?.d) ? envelope.d : [];

  const points = [];
  for (const p of raw) {
    if (!p || p.t === undefined || p.t === null) continue;
    let ms;
    try {
      ms = parseApiTimestamp(p.t);
    } catch {
      continue; // A sample we cannot place in time is worse than no sample.
    }
    if (!qualityPredicate(p.q)) continue;
    // A point can carry no value at all: the Historian logs quality 64
    // ("collector heartbeat lost") as a bare timestamp + quality. Such a point
    // never replaces the carried value; it only marks a quality issue.
    const novalue = p.v === undefined || p.v === null;
    // Missing quality is read as Good: the API always sends one, and synthetic
    // or hand-built series should not light up red.
    const issue = novalue || (p.q !== undefined && p.q !== null && classify(p.q) !== "good");
    points.push({ ms, v: novalue ? null : p.v, q: p.q, novalue, issue });
  }
  // The API returns data in order, but sorting makes the alignment independent
  // of that promise, and late-data backfill can legitimately arrive unordered.
  points.sort((a, b) => a.ms - b.ms);
  return { name, tag, points };
}

/**
 * Split a series at the window start into the boundary sample and the samples
 * inside the window.
 *
 * Identified by timestamp rather than by array position: the boundary point is
 * absent when the tag has no earlier history, and taking `d[0]` on faith would
 * then silently drop the first real sample.
 */
export function splitBoundary(points, windowStartMs) {
  if (windowStartMs === null || windowStartMs === undefined) {
    return { boundary: null, inside: points };
  }
  let cut = 0;
  while (cut < points.length && points[cut].ms < windowStartMs) cut++;
  return {
    boundary: cut > 0 ? points[cut - 1] : null,
    inside: points.slice(cut),
  };
}

/**
 * The value and quality status carried into a window from the points before it.
 *
 * The value is the last point that HAS one; the status is the last point's,
 * value or not. So a tag whose last pre-window point is a quality-64 note
 * enters the window holding its last real value, flagged as a quality issue.
 */
export function carryIn(pointsBefore) {
  if (!pointsBefore || pointsBefore.length === 0) return null;
  let lastValued = null;
  for (let i = pointsBefore.length - 1; i >= 0; i--) {
    if (!pointsBefore[i].novalue) {
      lastValued = pointsBefore[i];
      break;
    }
  }
  const last = pointsBefore[pointsBefore.length - 1];
  // `q` is the in-force point's quality: the last point's, value or not.
  return { ms: (lastValued ?? last).ms, v: lastValued ? lastValued.v : null, issue: !!last.issue, q: last.q };
}

/** Union of every series' timestamps, sorted and de-duplicated. */
export function unionAxis(seriesList) {
  const seen = new Set();
  for (const s of seriesList) {
    for (const p of s.points) seen.add(p.ms);
  }
  return Array.from(seen).sort((a, b) => a - b);
}

/**
 * Last-known-value alignment onto an arbitrary axis.
 *
 * Each series is walked once with a moving cursor rather than searched per
 * timestamp, so the cost is linear in total samples instead of
 * axis-length times series-length.
 */
export function alignStepped(seriesList, axis, boundaries) {
  const cursors = seriesList.map(() => 0);
  const held = boundaries.map((b) => (b ? b.v : null));
  const issue = boundaries.map((b) => !!(b && b.issue));
  // The point in force, for the quality columns: each row's range is the span
  // since the previous row, starting from whatever was in force then.
  const inForce = boundaries.map((b) => b ?? null);

  const rows = [];
  for (const t of axis) {
    const row = new Array(seriesList.length);
    const flags = new Array(seriesList.length);
    const stats = new Array(seriesList.length);
    for (let i = 0; i < seriesList.length; i++) {
      const pts = seriesList[i].points;
      let c = cursors[i];
      const st = rangeStats(inForce[i]);
      while (c < pts.length && pts[c].ms <= t) {
        if (!pts[c].novalue) held[i] = pts[c].v;
        issue[i] = !!pts[c].issue;
        st.add(pts[c]);
        inForce[i] = pts[c];
        c++;
      }
      cursors[i] = c;
      row[i] = held[i];
      flags[i] = issue[i];
      stats[i] = st;
    }
    rows.push({ ms: t, values: row, flags, stats });
  }
  return rows;
}

const isNum = (v) => typeof v === "number" && isFinite(v);

/**
 * One bucket's aggregate, with the value carried in from before the bucket.
 *
 * Carry-forward is the rule throughout: a tag holds its last known value until
 * it changes. So a bucket starts at `held` (the value in force at `from`)
 * unless a sample lands exactly on `from`, and an empty bucket is not "no
 * data" but "unchanged": every value aggregate reports the held value.
 *
 *   first  value in force at the bucket start
 *   last   value in force at the bucket end
 *   min / max / range   over the held value (while it lasts) and the samples
 *   avg    TIME-WEIGHTED: each value counts for as long as it was in force,
 *          over the part of the bucket inside the query window
 *   count / sum   statistics of the samples actually logged in the bucket
 *
 * Non-numeric values (text states, nulls) take part in first/last only.
 */
function aggregateBucket(kind, samples, held, from, to) {
  const startsOnSample = samples.length > 0 && samples[0].ms === from;
  const heldCounts = held !== null && held !== undefined && !startsOnSample;

  switch (kind) {
    case "first":
      return startsOnSample ? samples[0].v : heldCounts ? held : samples.length ? samples[0].v : null;
    case "last":
      return samples.length ? samples[samples.length - 1].v : held ?? null;
    case "count":
      return samples.length;
    case "sum":
      return samples.reduce((a, p) => (isNum(p.v) ? a + p.v : a), 0);
    case "min":
    case "max":
    case "range": {
      const nums = samples.map((p) => p.v).filter(isNum);
      if (heldCounts && isNum(held)) nums.push(held);
      if (!nums.length) return null;
      const lo = Math.min(...nums);
      const hi = Math.max(...nums);
      return kind === "min" ? lo : kind === "max" ? hi : hi - lo;
    }
    case "avg": {
      // Step segments: [from, s0) holds `held`, [s_i, s_i+1) holds s_i, last to `to`.
      const segments = [];
      if (heldCounts) segments.push({ v: held, a: from, b: samples.length ? samples[0].ms : to });
      samples.forEach((p, i) => segments.push({ v: p.v, a: p.ms, b: i + 1 < samples.length ? samples[i + 1].ms : to }));
      let weighted = 0;
      let duration = 0;
      for (const seg of segments) {
        const len = seg.b - seg.a;
        if (len > 0 && isNum(seg.v)) {
          weighted += seg.v * len;
          duration += len;
        }
      }
      if (duration > 0) return weighted / duration;
      // A sample in a zero-length tail (logged exactly at the window end) still counts.
      const nums = samples.map((p) => p.v).filter(isNum);
      return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : isNum(held) ? held : null;
    }
    default:
      throw new RangeError(`Unknown aggregate "${kind}". Use one of ${AGGREGATES.join(", ")}.`);
  }
}

/**
 * Bucket every series onto a fixed grid.
 *
 * Buckets are half-open [start, start+interval) and labelled by their start,
 * which is what makes consecutive buckets tile without double-counting a
 * sample that lands exactly on a boundary. Each bucket starts from the value
 * carried in from before it -- the boundary point for the first bucket -- so a
 * bucket with no change reports the held value rather than a blank.
 */
export function alignBucketed(seriesList, boundaries, { startMs, endMs, intervalMs, aggregate: kind }) {
  if (!(intervalMs > 0)) throw new RangeError("Bucketing needs an interval greater than zero.");
  if (!(endMs > startMs)) throw new RangeError("The window end must be after its start.");
  if (!AGGREGATES.includes(kind)) {
    throw new RangeError(`Unknown aggregate "${kind}". Use one of ${AGGREGATES.join(", ")}.`);
  }

  const bucketCount = Math.ceil((endMs - startMs) / intervalMs);
  const cursors = seriesList.map(() => 0);
  const held = boundaries.map((b) => (b ? b.v : null));
  const issue = boundaries.map((b) => !!(b && b.issue));
  const inForce = boundaries.map((b) => b ?? null);

  const rows = [];
  for (let b = 0; b < bucketCount; b++) {
    const from = startMs + b * intervalMs;
    // The last bucket may run past the window end; time-weighting stops at the
    // end, because nothing is known about the value after it.
    const to = Math.min(from + intervalMs, endMs);
    const row = new Array(seriesList.length);
    const flags = new Array(seriesList.length);
    const stats = new Array(seriesList.length);

    for (let i = 0; i < seriesList.length; i++) {
      const pts = seriesList[i].points;
      let c = cursors[i];
      const inBucket = [];
      const logged = [];
      while (c < pts.length && pts[c].ms < from) {
        // Before the window's first bucket: just carry it.
        if (!pts[c].novalue) held[i] = pts[c].v;
        issue[i] = !!pts[c].issue;
        inForce[i] = pts[c];
        c++;
      }
      let bucketIssue = issue[i]; // a quality issue in force at the bucket start counts
      while (c < pts.length && pts[c].ms < from + intervalMs) {
        const p = pts[c];
        if (p.issue) bucketIssue = true;
        if (!p.novalue) inBucket.push(p);
        logged.push(p);
        c++;
      }
      // Quality over the slice: the point in force at its start (unless a
      // point lands exactly on the start and replaces it), then every point in it.
      const st = rangeStats(logged.length && logged[0].ms === from ? null : inForce[i]);
      for (const p of logged) st.add(p);
      stats[i] = st;
      if (logged.length) inForce[i] = logged[logged.length - 1];
      cursors[i] = c;
      row[i] = aggregateBucket(kind, inBucket, held[i], from, to);
      flags[i] = bucketIssue;
      if (inBucket.length) held[i] = inBucket[inBucket.length - 1].v;
      // Status carried into the next bucket: the last point in this one, value or not.
      if (c > 0 && pts[c - 1].ms >= from) issue[i] = !!pts[c - 1].issue;
    }
    rows.push({ ms: from, values: row, flags, stats });
  }
  return rows;
}

/**
 * The whole job: a TVQsResponse in, a 2-D array ready for Excel out.
 *
 * Options:
 *   interval        ms per bucket, or null for raw union alignment
 *   aggregate       aggregate name, used only when interval is set
 *   qualityFilter   predicate applied per sample before alignment
 *   headers         prepend a header row
 *   boundary        raw mode only:
 *                     "start" (default) -- a row AT the window start carrying
 *                       each tag's last known value, unless a sample already
 *                       lands exactly there. A 12:00 start reads 12:00 first.
 *                     "keep" -- a row at the pre-window sample's own time
 *                     "drop" -- no seed row (used when no start was given)
 *   seedResponse    an optional TVQsResponse over a lookback window, used for
 *                   tags whose boundary point was missing or rejected by the
 *                   quality filter, so the carried-in value still exists
 *   novalue         points that carry no value (quality-64 notes):
 *                     "flag" (default) -- keep the row, carry the last value
 *                       through it, and flag the cell as a quality issue
 *                     "drop" -- leave such points out entirely
 *   flagCell        (value) => cell, applied to every flagged cell (a value
 *                   whose latest point is not Good). Identity by default; the
 *                   worksheet function turns it into red-formatted output.
 *   columns         companion columns after each tag's value, from
 *                   columns.js: any of "quality", "status", "samples", "good%"
 *   percentCell     (fraction) => cell for "good%"; identity by default
 *   timeCell        (excelSerial) => cell for the Timestamp column; identity by
 *                   default, a formatted date in the worksheet
 *   maxRows         refuse rather than silently truncate past this
 *   utc             emit UTC wall-clock instead of local
 *   tagOrder        column order; defaults to response order
 */
export function buildGrid(response, options = {}) {
  const {
    interval = null,
    aggregate: aggKind = "avg",
    qualityPredicate = () => true,
    headers = true,
    boundary = "start",
    maxRows = 100000,
    utc = false,
    tagOrder = null,
    seedResponse = null,
    novalue = "flag",
    flagCell = (v) => v,
    columns = [],
    percentCell = (f) => f,
    timeCell = (serial) => serial,
  } = options;
  if (novalue !== "flag" && novalue !== "drop") {
    throw new RangeError(`novalue must be "flag" or "drop", not "${novalue}".`);
  }
  const keep = novalue === "drop" ? (p) => !p.novalue : () => true;
  const normalize = (e) => {
    const ser = normalizeSeries(e, { qualityPredicate });
    return { ...ser, points: ser.points.filter(keep) };
  };

  const envelopes = Array.isArray(response?.tl) ? response.tl : [];
  let seriesList = envelopes.map(normalize);

  if (tagOrder && tagOrder.length) {
    const byName = new Map(seriesList.map((s) => [s.name, s]));
    seriesList = tagOrder.map(
      (name) => byName.get(name) ?? { name, tag: { n: name }, points: [] },
    );
  }

  const startMs = response?.s !== undefined && response?.s !== null ? parseApiTimestamp(response.s) : null;
  const endMs = response?.e !== undefined && response?.e !== null ? parseApiTimestamp(response.e) : null;

  const seeds = new Map();
  if (seedResponse && Array.isArray(seedResponse.tl) && startMs !== null) {
    for (const e of seedResponse.tl) {
      const ser = normalize(e);
      seeds.set(ser.name, ser.points.filter((p) => p.ms < startMs));
    }
  }

  // Split every series at the window start. What it carries in comes from all
  // its pre-window points -- the response's boundary point plus any lookback
  // seed -- so a boundary point with no value still yields the last real one.
  const boundaries = [];
  const trimmed = seriesList.map((s) => {
    if (startMs === null) {
      boundaries.push(null);
      return s;
    }
    const before = s.points.filter((p) => p.ms < startMs);
    const inside = s.points.filter((p) => p.ms >= startMs);
    const merged = [...(seeds.get(s.name) ?? []), ...before].sort((a, b) => a.ms - b.ms);
    boundaries.push(carryIn(merged));
    return { ...s, points: inside };
  });

  let rows;
  if (interval) {
    if (startMs === null || endMs === null) {
      throw new Error(
        "Bucketing needs the query window, which the response did not report in its s and e fields.",
      );
    }
    rows = alignBucketed(trimmed, boundaries, {
      startMs,
      endMs,
      intervalMs: interval,
      aggregate: aggKind,
    });
  } else {
    const axis = unionAxis(trimmed);
    // Carry-forward from the start: the first row is the window start itself,
    // holding each tag's last known value, then every change after it.
    if (boundary === "start" && startMs !== null && axis[0] !== startMs && boundaries.some(Boolean)) {
      axis.unshift(startMs);
    }
    rows = alignStepped(trimmed, axis, boundaries);

    if (boundary === "keep") {
      // One extra row at the earliest boundary timestamp, carrying whatever
      // each tag held going into the window.
      const boundaryTimes = boundaries.filter(Boolean).map((b) => b.ms);
      if (boundaryTimes.length) {
        rows.unshift({
          ms: Math.min(...boundaryTimes),
          values: boundaries.map((b) => (b ? b.v : null)),
          flags: boundaries.map((b) => !!(b && b.issue)),
        });
      }
    }
  }

  if (rows.length > maxRows) {
    throw new RangeError(
      `This query aligns to ${rows.length.toLocaleString("en-US")} rows, over the ${maxRows.toLocaleString("en-US")} row limit. ` +
        `The Historian API has no server-side aggregation, so narrow the time range, or pass an interval such as "1m" to bucket the data.`,
    );
  }

  const grid = [];
  if (headers) grid.push(["Timestamp", ...trimmed.flatMap((s) => [s.name, ...columnHeaders(s.name, columns)])]);
  for (const r of rows) {
    grid.push([
      timeCell(msToExcelSerial(r.ms, { utc })),
      ...r.values.flatMap((v, i) => {
        const value = v === undefined ? null : v;
        const cell = r.flags && r.flags[i] ? flagCell(value) : value;
        return [cell, ...columnCells(r.stats ? r.stats[i] : null, columns, { percentCell })];
      }),
    ]);
  }

  // A spilled array must never be empty: Excel shows #CALC! for a zero-length
  // result, which reads as a broken formula rather than "no data in range".
  if (grid.length === 0) return [["Timestamp"]];
  return grid;
}
