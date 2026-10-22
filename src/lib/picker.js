// Helpers behind the Insert tab's date/time pickers and tag list.
//
// The pickers are <input type="datetime-local">, whose value is a wall-clock
// string with no zone ("2026-10-22T07:00"). Written into a formula as-is it
// would be ambiguous in the worst way: the Historian reads a bare ISO string
// as ITS local time, and a containerised Historian usually runs on UTC, so
// 07:00 picked in Eastern time would query 03:00. The picked time is therefore
// written with this workstation's UTC offset ("2026-10-22T07:00:00-04:00"),
// which every reader -- the Historian, time.js, a person -- resolves the same.

const pad = (n) => String(Math.abs(Math.trunc(n))).padStart(2, "0");

/** "+HH:MM" / "-HH:MM" for a JS getTimezoneOffset() value (minutes WEST of UTC). */
function offsetSuffix(timezoneOffsetMinutes) {
  const east = -timezoneOffsetMinutes;
  const sign = east >= 0 ? "+" : "-";
  return `${sign}${pad(east / 60)}:${pad(east % 60)}`;
}

/**
 * A datetime-local value to an ISO timestamp with this machine's offset.
 *
 * `offsetMinutesFor(date)` defaults to the real zone and is injectable for
 * tests; it is evaluated for the picked instant, so a time picked across a
 * DST change gets that day's offset, not today's.
 */
export function pickerValueToIso(value, offsetMinutesFor = (d) => d.getTimezoneOffset()) {
  if (typeof value !== "string" || value.trim() === "") return "";
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/.exec(value.trim());
  if (!m) throw new RangeError(`Not a date and time: "${value}"`);
  const [, y, mo, d, h, mi, s = "00"] = m;
  const local = new Date(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), Number(s));
  return `${y}-${mo}-${d}T${h}:${mi}:${s}${offsetSuffix(offsetMinutesFor(local))}`;
}

/**
 * The reverse, for opening a picker on whatever the text box holds: an
 * absolute timestamp becomes that local wall-clock time; anything else
 * (blank, a relative expression like "-24h") opens on `now`, rounded down to
 * the minute.
 */
export function isoToPickerValue(text, now = new Date()) {
  let date = null;
  if (typeof text === "string" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(text.trim())) {
    const ms = Date.parse(text.trim().replace(/(\.\d{3})\d+/, "$1"));
    if (!Number.isNaN(ms)) date = new Date(ms);
  }
  if (!date) date = now;
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
    `T${pad(date.getHours())}:${pad(date.getMinutes())}`
  );
}

/** Append names not already present, keeping first-seen order. */
export function addTags(list, names) {
  const out = [...list];
  for (const n of names) {
    const name = String(n ?? "").trim();
    if (name && !out.includes(name)) out.push(name);
  }
  return out;
}

export function removeTag(list, name) {
  return list.filter((n) => n !== name);
}
