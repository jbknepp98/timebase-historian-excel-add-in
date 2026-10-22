// Time handling between Excel, JavaScript and the Historian API.
//
// Three representations meet here and each has a trap:
//
//   Excel   serial number, days since 1899-12-30, wall-clock with NO timezone
//   ISO     what the API speaks; a bare "2026-08-10T17:01:58" is LOCAL time,
//           the same string with "Z" is UTC
//   relative the API's own "-1h" / "d+6h" grammar
//
// A pure offset ("-24h", "-90m") means the same instant in every time zone, so
// it is passed through for the server to resolve: no clock-skew disagreement
// between workstation and Historian, and the response echoes the resolved
// window back as `s` and `e`.
//
// An expression with a CLAMP ("d", "d+12h", "w", "M-1M") is different: "today"
// depends on whose today. The Historian resolves it on its own clock, and a
// containerised Historian runs on UTC -- so "d+12h" meant 12:00 UTC, 08:00 in
// New York. Clamped expressions are therefore resolved here, in the
// workstation's time zone (the zone every timestamp in the sheet is shown in),
// and sent as an absolute timestamp with the UTC offset.

/** Days between the Excel epoch (1899-12-30) and the Unix epoch. */
const EXCEL_UNIX_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;

/** Clamp letters accepted at the start of a relative expression. */
const CLAMP_UNITS = ["y", "M", "w", "d", "h", "m"];
/** Offset letters accepted in the trailing offset of a relative expression. */
const OFFSET_UNITS = ["y", "M", "d", "h", "m", "s"];

/**
 * Is this text one of the API's relative time expressions?
 *
 * Grammar, from the API description:
 *   [clamp][offset]
 *   clamp  = one of y M w d h m          (start of the current period)
 *   offset = [-]<count><unit>, unit one of y M d h m s
 *
 * The leading character disambiguates: a digit or sign starts a bare offset,
 * a clamp letter starts a clamped expression. Case is significant, because
 * `M` is months and `m` is minutes.
 */
export function parseRelative(text) {
  if (typeof text !== "string") return null;
  const s = text.trim();
  if (s === "") return null;

  const first = s[0];
  const startsWithOffset = first === "-" || first === "+" || (first >= "0" && first <= "9");

  let clamp = null;
  let rest = s;

  if (!startsWithOffset) {
    if (!CLAMP_UNITS.includes(first)) return null;
    clamp = first;
    rest = s.slice(1);
    if (rest === "") return { clamp, offset: null, text: s };
  }

  const m = /^([-+]?)(\d+)([yMdhms])$/.exec(rest);
  if (!m) return null;
  const [, sign, count, unit] = m;
  if (!OFFSET_UNITS.includes(unit)) return null;

  return {
    clamp,
    offset: { sign: sign === "-" ? -1 : 1, count: Number(count), unit },
    text: s,
  };
}

export function isRelative(text) {
  return parseRelative(text) !== null;
}

const pad2 = (n) => String(Math.abs(Math.trunc(n))).padStart(2, "0");

/** An instant as local wall-clock ISO with its UTC offset: "2026-10-22T12:00:00-04:00". */
export function toLocalIsoWithOffset(ms) {
  const d = new Date(ms);
  const east = -d.getTimezoneOffset();
  const sign = east >= 0 ? "+" : "-";
  return (
    `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}` +
    `T${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}` +
    `${sign}${pad2(east / 60)}:${pad2(east % 60)}`
  );
}

/**
 * Resolve a CLAMPED relative expression in local time, returning epoch ms.
 *
 * Clamp: y = Jan 1, M = the 1st, w = Sunday (as the API defines weeks),
 * d = midnight, h = top of the hour, m = top of the minute -- all local. The
 * offset is then applied on the local calendar for y/M/d (so "d-1d" is
 * yesterday's midnight even across a DST change) and as elapsed time for h/m/s.
 */
export function resolveClampedLocal(parsed, now = new Date()) {
  if (!parsed || !parsed.clamp) throw new RangeError("resolveClampedLocal needs a clamped expression.");
  const d = new Date(now.getTime());
  switch (parsed.clamp) {
    case "y":
      d.setMonth(0, 1);
      d.setHours(0, 0, 0, 0);
      break;
    case "M":
      d.setDate(1);
      d.setHours(0, 0, 0, 0);
      break;
    case "w":
      d.setDate(d.getDate() - d.getDay());
      d.setHours(0, 0, 0, 0);
      break;
    case "d":
      d.setHours(0, 0, 0, 0);
      break;
    case "h":
      d.setMinutes(0, 0, 0);
      break;
    case "m":
      d.setSeconds(0, 0);
      break;
    default:
      throw new RangeError(`Unknown clamp "${parsed.clamp}".`);
  }
  if (parsed.offset) {
    const n = parsed.offset.sign * parsed.offset.count;
    switch (parsed.offset.unit) {
      case "y":
        d.setFullYear(d.getFullYear() + n);
        break;
      case "M":
        d.setMonth(d.getMonth() + n);
        break;
      case "d":
        d.setDate(d.getDate() + n);
        break;
      case "h":
        return d.getTime() + n * 3600000;
      case "m":
        return d.getTime() + n * 60000;
      case "s":
        return d.getTime() + n * 1000;
    }
  }
  return d.getTime();
}

/**
 * Excel serial number to an ISO string with no offset suffix.
 *
 * No suffix is correct rather than lazy: an Excel serial carries wall-clock
 * time with no zone, and the API documents a bare ISO string as local time.
 * Appending "Z" here would silently shift every query by the UTC offset.
 */
export function excelSerialToIsoLocal(serial) {
  if (typeof serial !== "number" || !isFinite(serial)) {
    throw new RangeError(`Not an Excel date serial number: ${serial}`);
  }
  // Excel on Windows treats 1900 as a leap year, so serials 1..60 are off by
  // one against the real calendar. Dates that old are meaningless for a
  // historian, so reject them rather than return a silently wrong day.
  if (serial < 61) {
    throw new RangeError(
      `Excel date serial ${serial} predates 1900-03-01, where Excel's leap-year bug makes the date ambiguous.`,
    );
  }
  const ms = Math.round((serial - EXCEL_UNIX_OFFSET_DAYS) * MS_PER_DAY);
  const d = new Date(ms);
  // Read back through the UTC accessors: `ms` was built from a wall-clock
  // serial, so the UTC fields hold exactly the wall-clock fields we want.
  const p2 = (n) => String(n).padStart(2, "0");
  const p3 = (n) => String(n).padStart(3, "0");
  const base =
    `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}` +
    `T${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())}`;
  const msPart = d.getUTCMilliseconds();
  return msPart === 0 ? base : `${base}.${p3(msPart)}`;
}

/**
 * Parse an API timestamp to epoch milliseconds.
 *
 * Handles the three documented shapes -- bare local, trailing Z, explicit
 * offset -- and the sub-millisecond precision the Historian emits (the API
 * examples carry six fractional digits, which Date would otherwise mis-read
 * on some engines).
 */
export function parseApiTimestamp(text) {
  if (typeof text !== "string" || text.trim() === "") {
    throw new RangeError(`Not a timestamp: ${text}`);
  }
  const s = text.trim();
  // Truncate fractional seconds to milliseconds, keeping any zone suffix.
  const normalized = s.replace(/(\.\d{3})\d+/, "$1");
  const ms = Date.parse(normalized);
  if (Number.isNaN(ms)) throw new RangeError(`Unrecognised timestamp: ${text}`);
  return ms;
}

/**
 * Epoch milliseconds to an Excel serial number.
 *
 * `utc: false` (the default) converts to the workstation's local wall clock,
 * because that is what a person reading the sheet expects to see and Excel has
 * no way to represent a zone. `utc: true` keeps UTC wall-clock instead.
 */
export function msToExcelSerial(ms, { utc = false } = {}) {
  const shifted = utc ? ms : ms - new Date(ms).getTimezoneOffset() * 60000;
  return shifted / MS_PER_DAY + EXCEL_UNIX_OFFSET_DAYS;
}

/**
 * Turn whatever Excel handed us into the right query parameter for the API.
 *
 * Returns { param, value } where param is one of start/end, relativeStart/
 * relativeEnd, or unixStart/unixEnd, ready to go straight onto the query
 * string. `which` is "start" or "end".
 */
export function toQueryParam(input, which, { now = () => new Date() } = {}) {
  if (which !== "start" && which !== "end") {
    throw new Error(`toQueryParam expects "start" or "end", got ${which}`);
  }
  if (input === null || input === undefined || input === "") return null;

  // Excel hands a real date cell to a custom function as a serial number.
  if (typeof input === "number") {
    return { param: which, value: excelSerialToIsoLocal(input) };
  }

  if (input instanceof Date) {
    return { param: which, value: input.toISOString() };
  }

  if (typeof input === "string") {
    const s = input.trim();
    if (s === "") return null;
    const rel = parseRelative(s);
    if (rel && rel.clamp) {
      // "Today" is the user's today, not the Historian's: resolve it here.
      return { param: which, value: toLocalIsoWithOffset(resolveClampedLocal(rel, now())) };
    }
    if (rel) {
      return { param: which === "start" ? "relativeStart" : "relativeEnd", value: s };
    }
    // Anything else must be an absolute timestamp. Validate now so a typo
    // surfaces here rather than as an opaque API error.
    parseApiTimestamp(s);
    return { param: which, value: s };
  }

  throw new TypeError(
    `A time bound must be an Excel date, an ISO timestamp, or a relative expression such as "-24h". Got ${typeof input}.`,
  );
}

/** Parse an interval such as "30s", "5m", "1h", "1d" to milliseconds. */
export function parseInterval(text) {
  if (text === null || text === undefined || text === "") return null;
  if (typeof text === "number") {
    // A bare number is read as seconds; 0 means "no bucketing".
    if (!isFinite(text) || text < 0) throw new RangeError(`Not an interval: ${text}`);
    return text === 0 ? null : Math.round(text * 1000);
  }
  const m = /^(\d+(?:\.\d+)?)\s*(ms|s|m|h|d)$/.exec(String(text).trim());
  if (!m) {
    throw new RangeError(
      `Not an interval: "${text}". Use a count and a unit, such as "30s", "5m", "1h" or "1d".`,
    );
  }
  const n = Number(m[1]);
  const factor = { ms: 1, s: 1000, m: 60000, h: 3600000, d: MS_PER_DAY }[m[2]];
  const ms = Math.round(n * factor);
  if (ms <= 0) throw new RangeError(`Interval must be greater than zero: "${text}"`);
  return ms;
}

/**
 * How every timestamp the add-in writes is displayed: month/day/year and time
 * to the millisecond. In Excel's format codes "mm" right after "hh" means
 * minutes, so this is month first, minutes in the time.
 */
export const TIMESTAMP_FORMAT = "mm/dd/yy hh:mm:ss.000";

/**
 * An Excel date serial as a cell that carries its own display format. It is
 * still a real date (a FormattedNumber), so sorting, charts and date
 * arithmetic work on it; only how it reads changes.
 */
export function timestampCell(serial) {
  if (typeof serial !== "number" || !isFinite(serial)) return serial;
  return { type: "FormattedNumber", basicValue: serial, numberFormat: TIMESTAMP_FORMAT };
}

export const _internal = { EXCEL_UNIX_OFFSET_DAYS, MS_PER_DAY };
