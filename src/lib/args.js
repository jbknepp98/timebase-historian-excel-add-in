// Custom-function argument shaping.
//
// Excel passes a single cell as a scalar but a range as a 2-D array, even when
// the range is one column or one row. A function that takes "one or more tag
// names" has to accept =TB.HISTORY("A1") and =TB.HISTORY(A1:A20) identically,
// so every list-shaped parameter goes through here first.

/**
 * Flatten a scalar-or-2D-array argument into a list of non-blank strings.
 *
 * Excel represents a blank cell inside a range as "" (never null/undefined),
 * so "" is filtered out here rather than left to trip up a URL builder later.
 */
export function flattenTagList(arg) {
  if (arg === null || arg === undefined) return [];

  if (Array.isArray(arg)) {
    const out = [];
    for (const row of arg) {
      const cells = Array.isArray(row) ? row : [row];
      for (const cell of cells) {
        if (cell === null || cell === undefined) continue;
        const s = String(cell).trim();
        if (s !== "") out.push(s);
      }
    }
    return out;
  }

  const s = String(arg).trim();
  return s === "" ? [] : [s];
}

/**
 * Same shape, for a single scalar-or-range argument that must resolve to
 * exactly one value (a dataset name, for instance). Throws if the caller
 * passed a range with more than one non-blank cell, rather than silently
 * taking the first.
 */
export function requireScalar(arg, label) {
  const list = flattenTagList(arg);
  if (list.length === 0) {
    throw new RangeError(`${label} is required.`);
  }
  if (list.length > 1) {
    throw new RangeError(`${label} must be a single cell, not a range of ${list.length} cells.`);
  }
  return list[0];
}

/** Excel's #N/A, #VALUE! etc. arrive as {type: "Error", ...}, not a JS Error. */
export function isExcelError(value) {
  return !!value && typeof value === "object" && value.type === "Error";
}

/** Reject an Excel error object with a message pointing at which argument. */
export function rejectIfError(value, label) {
  if (isExcelError(value)) {
    throw new RangeError(`${label} contains an error value (${value.errorType ?? "#ERROR"}) instead of data.`);
  }
  return value;
}

/**
 * Make a result table safe to spill.
 *
 * Excel shows a `null` inside a custom function's returned array as 0, not as
 * an empty cell -- so a tag with no value yet, or a sample with no value
 * (quality 64 "collector heartbeat lost" points carry none), would read as a
 * real state 0. An empty string is what renders blank.
 */
export function blankNulls(grid) {
  return grid.map((row) => row.map((v) => (v === null || v === undefined ? "" : v)));
}
