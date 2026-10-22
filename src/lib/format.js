// Units of measure and display formatting.
//
// The Tag schema's `u` dictionary is overloaded, and the rules come straight
// from the API description:
//
//   single key, negative   the unit goes BEFORE the number   { "-1": "US$" }
//   single key, positive   the unit goes AFTER the number    { "1": "kW" }
//   keys 1 and 0           boolean labels                    { "1": "Running", "0": "Stopped" }
//   multiple keys          an enumeration keyed by value      { "1": "Apple", "2": "Orange" }
//
// Note the collision: { "1": "kW" } and a two-key boolean map are both "keys
// including 1". Only the number of entries separates a suffix unit from a state
// enumeration, so the count is checked before anything else.
//
// Custom functions deliberately return raw numbers, not formatted strings, so
// Excel can chart and aggregate them. This module exists for the task pane, for
// column headers, and for the explicit "display" mode of TB.VALUE.

/** "prefix" | "suffix" | "boolean" | "enum" | "none" */
export function classifyUnits(units) {
  if (!units || typeof units !== "object") return "none";
  const keys = Object.keys(units);
  if (keys.length === 0) return "none";

  if (keys.length === 1) {
    const n = Number(keys[0]);
    if (!isFinite(n)) return "none";
    return n < 0 ? "prefix" : "suffix";
  }

  const numeric = keys.map(Number).filter((n) => isFinite(n));
  if (numeric.length === 2 && numeric.includes(0) && numeric.includes(1)) return "boolean";
  return "enum";
}

/** The bare unit text for a scalar tag, or "" when the tag is enumerated. */
export function unitLabel(units) {
  const kind = classifyUnits(units);
  if (kind !== "prefix" && kind !== "suffix") return "";
  return String(units[Object.keys(units)[0]] ?? "");
}

/**
 * Number of decimal places implied by a .NET format string.
 *
 * Only the shapes that actually appear on Historian tags are handled -- "0.00",
 * "#,##0.0", "F2", "N1". Anything else returns null, which callers read as
 * "leave the number alone" rather than guessing wrong.
 */
export function decimalsFromFormat(format) {
  if (typeof format !== "string" || format.trim() === "") return null;
  const f = format.trim();

  const std = /^([FfNn])(\d+)$/.exec(f);
  if (std) return Number(std[2]);

  const custom = /^[#,0]*\.(0+|#+)$/.exec(f);
  if (custom) return custom[1].length;

  if (/^[#,0]+$/.test(f)) return 0;

  return null;
}

/**
 * Render a value for display, applying enumeration labels and units.
 *
 * Used for the task pane preview and TB.VALUE(..., "display"). Where the API
 * already gives a server-rendered string in `l.vs`, prefer that instead: it
 * resolves the tag's own format and unit exactly as the Historian would.
 */
export function formatValue(value, tag = {}) {
  if (value === null || value === undefined) return "";

  const units = tag.units ?? null;
  const kind = classifyUnits(units);

  if (kind === "boolean") {
    // A boolean tag may arrive as true/false or as 1/0 depending on its
    // declared DataType, so both are mapped through the same label pair.
    const key = value === true || value === 1 || value === "1" ? "1" : "0";
    return String(units[key] ?? value);
  }

  if (kind === "enum") {
    const key = String(typeof value === "number" ? Math.trunc(value) : value);
    // An unmapped state is shown as the raw code rather than blank: a number
    // with no label is still information, and blank looks like missing data.
    return units[key] !== undefined ? String(units[key]) : String(value);
  }

  let text;
  if (typeof value === "number") {
    const decimals = decimalsFromFormat(tag.format);
    text = decimals === null ? String(value) : value.toFixed(decimals);
  } else {
    text = String(value);
  }

  const unit = unitLabel(units);
  if (!unit) return text;
  return kind === "prefix" ? `${unit}${text}` : `${text} ${unit}`;
}

/** Column heading for a tag: its name, plus its unit when it has a scalar one. */
export function columnHeading(tag = {}) {
  const name = tag.name ?? "";
  const unit = unitLabel(tag.units);
  return unit ? `${name} (${unit})` : name;
}

/** Is this tag a state enumeration rather than a measurement? */
export function isEnumerated(tag = {}) {
  const kind = classifyUnits(tag.units);
  return kind === "enum" || kind === "boolean";
}
