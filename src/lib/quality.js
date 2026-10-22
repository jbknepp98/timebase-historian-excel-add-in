// OPC quality codes.
//
// The API documents quality as "represented as per OPC (e.g. 192 is Good)".
// OPC DA encodes the coarse verdict in the top two bits of the low byte, so the
// test is a mask, not equality with 192. Treating only 192 as good would
// discard every good-with-substatus code the Historian passes through from a
// data source.
//
//   0xC0  Good        192..255
//   0x40  Uncertain    64..127
//   0x00  Bad           0..63

export const QUALITY_MASK = 0xc0;
export const GOOD = 0xc0;
export const UNCERTAIN = 0x40;
export const BAD = 0x00;

/** "good" | "uncertain" | "bad" | "unknown" */
export function classify(q) {
  if (typeof q !== "number" || !isFinite(q)) return "unknown";
  switch (q & QUALITY_MASK) {
    case GOOD:
      return "good";
    case UNCERTAIN:
      return "uncertain";
    case BAD:
      return "bad";
    default:
      // 0x80 is not defined by OPC DA as a coarse verdict on its own.
      return "unknown";
  }
}

export function isGood(q) {
  return classify(q) === "good";
}

/**
 * Build a predicate from a filter name.
 *
 *   "any" (default) keeps everything
 *   "good"          keeps Good only
 *   "notbad"        keeps Good and Uncertain
 */
export function predicateFor(filter) {
  const f = String(filter ?? "any").trim().toLowerCase();
  switch (f) {
    case "":
    case "any":
    case "all":
      return () => true;
    case "good":
      return (q) => classify(q) === "good";
    case "notbad":
    case "not-bad":
      return (q) => {
        const c = classify(q);
        return c === "good" || c === "uncertain";
      };
    default:
      throw new RangeError(`Unknown quality filter "${filter}". Use any, good or notbad.`);
  }
}

/**
 * Codes the Historian itself writes to a dataset status tag, per the docs.
 * Surfaced so a report can tell "no data because the collector was down" apart
 * from "no data because nothing changed".
 */
export const COLLECTOR_STATUS_CODES = {
  64: "Store and forward (collector heartbeat lost for 10s)",
  28: "Collector shut down",
  24: "Collector up, but lost communication with its data source",
};

export function describe(q) {
  if (typeof q !== "number" || !isFinite(q)) return "unknown";
  const known = COLLECTOR_STATUS_CODES[q];
  const base = classify(q);
  return known ? `${base} (${known})` : base;
}
