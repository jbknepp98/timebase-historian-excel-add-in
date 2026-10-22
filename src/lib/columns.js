// Optional per-tag companion columns: quality, status, sample count, % Good.
//
// Every output row covers a RANGE of time -- an interval slice, the span since
// the previous clock mark (TB.SAMPLE), or the span since the previous row (raw
// history) -- and these columns describe the points behind that row:
//
//   quality   the WORST OPC quality code in the range, counting the point in
//             force at the range start (a value carried from a bad point is
//             still bad). Bad < Uncertain < Good; ties go to the lower code.
//   status    the same worst quality, in words.
//   samples   how many points the Historian logged inside the range.
//   good%     the share of those points with Good quality; blank when none
//             were logged (nothing to take a share of).
//
// Missing quality on a point reads as Good (192), as it does everywhere else.

import { classify, COLLECTOR_STATUS_CODES } from "./quality.js";

export const COLUMN_KINDS = ["quality", "status", "samples", "good%"];

const ALIASES = {
  quality: "quality",
  q: "quality",
  status: "status",
  qualitytext: "status",
  samples: "samples",
  count: "samples",
  "good%": "good%",
  "%good": "good%",
  goodpct: "good%",
  pctgood: "good%",
  "percent good": "good%",
};

const HEADERS = { quality: "Quality", status: "Status", samples: "Samples", "good%": "% Good" };

export const GOOD_CODE = 192;

/**
 * Parse the `columns` argument: blank, a comma/space separated string, or a
 * range/array of names. Order is the caller's; duplicates are dropped.
 */
export function parseColumns(input) {
  const parts = [];
  const visit = (x) => {
    if (x === null || x === undefined) return;
    if (Array.isArray(x)) return x.forEach(visit);
    for (const p of String(x).split(/[,;]/)) parts.push(p);
  };
  visit(input);
  const out = [];
  for (const raw of parts) {
    const key = raw.trim().toLowerCase();
    if (!key) continue;
    const kind = ALIASES[key];
    if (!kind) {
      throw new RangeError(`Unknown column "${raw.trim()}". Use any of quality, status, samples, good%.`);
    }
    if (!out.includes(kind)) out.push(kind);
  }
  return out;
}

const SEVERITY = { bad: 0, unknown: 0, uncertain: 1, good: 2 };

/** The worse of two quality codes (null/undefined means "nothing yet"). */
export function worseQuality(a, b) {
  if (a === null || a === undefined) return b ?? null;
  if (b === null || b === undefined) return a;
  const sa = SEVERITY[classify(a)];
  const sb = SEVERITY[classify(b)];
  if (sa !== sb) return sa < sb ? a : b;
  return Math.min(a, b);
}

/** A point's quality code, reading a missing one as Good. */
export function codeOf(point) {
  const q = point && point.q;
  return typeof q === "number" && isFinite(q) ? q : GOOD_CODE;
}

/**
 * Accumulates one cell's range statistics. Seed it with the point in force at
 * the range start, then add every point logged inside the range.
 */
export function rangeStats(inForce = null) {
  return {
    worst: inForce ? codeOf(inForce) : null,
    count: 0,
    good: 0,
    add(point) {
      const code = codeOf(point);
      this.worst = worseQuality(this.worst, code);
      this.count++;
      if (classify(code) === "good") this.good++;
    },
  };
}

/** "Good", "Uncertain: store and forward (collector heartbeat lost for 10s)", "Bad". */
export function statusText(code) {
  if (code === null || code === undefined) return null;
  const word = { good: "Good", uncertain: "Uncertain", bad: "Bad", unknown: "Unknown" }[classify(code)];
  const known = COLLECTOR_STATUS_CODES[code];
  return known ? `${word}: ${known.charAt(0).toLowerCase()}${known.slice(1)}` : word;
}

/** Header cells for one tag's companion columns. */
export function columnHeaders(tagName, columns) {
  return columns.map((c) => `${tagName} ${HEADERS[c]}`);
}

/**
 * Cells for one tag's companion columns. `percentCell(fraction)` formats the
 * % Good value (0..1); identity by default, a "0%" number in the worksheet.
 */
export function columnCells(stats, columns, { percentCell = (f) => f } = {}) {
  return columns.map((c) => {
    if (!stats) return null;
    switch (c) {
      case "quality":
        return stats.worst;
      case "status":
        return statusText(stats.worst);
      case "samples":
        return stats.count;
      case "good%":
        return stats.count ? percentCell(stats.good / stats.count) : null;
      default:
        return null;
    }
  });
}
