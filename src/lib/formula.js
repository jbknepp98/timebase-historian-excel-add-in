// Excel formula text generation for the task pane's "Insert" tab and its tag
// browser's "insert selected" button.
//
// This exists as its own tested module because a formula the task pane
// inserts has to be syntactically perfect on the first try -- there is no
// compiler to catch a stray quote before it lands in the user's cell. Two
// Excel-specific rules matter:
//
//   - A string literal doubles its internal quotes: "Say ""hi""" is the
//     three-character string  Say "hi".
//   - Excel's array-constant syntax lets a list of strings be written inline
//     as {"a","b","c"}, which TB.HISTORY's `tags` parameter accepts exactly
//     like a range, so multiple tags never require a helper range of cells.

/** Excel string-literal escaping: wrap in quotes, double any interior quote. */
export function quote(text) {
  return `"${String(text).replace(/"/g, '""')}"`;
}

/**
 * Render one argument for a formula: a blank/null value becomes null (meaning
 * "omit"), a string becomes a quoted literal (unless it already looks like a
 * cell/range reference or a formula fragment, in which case it is passed
 * through verbatim), and anything else is passed through as-is.
 */
function renderArg(value, { asRefIfLooksLikeOne = false } = {}) {
  if (value === null || value === undefined) return null;
  if (typeof value === "string") {
    const s = value.trim();
    if (s === "") return null;
    if (asRefIfLooksLikeOne && looksLikeReference(s)) return s;
    return quote(s);
  }
  return String(value);
}

/** A1, $A$1, A1:B20, Sheet1!A1 -- rough enough to distinguish from a literal. */
export function looksLikeReference(s) {
  return /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(s) || /^[^!]+![A-Za-z$]/.test(s);
}

/** One or more tag names as an Excel argument: a single string, or {"a","b"}. */
export function tagsArg(tags) {
  const list = (Array.isArray(tags) ? tags : [tags]).map((t) => String(t).trim()).filter(Boolean);
  if (list.length === 0) return null;
  if (list.length === 1) return quote(list[0]);
  return `{${list.map(quote).join(",")}}`;
}

/** Companion columns as one quoted list, e.g. "quality,samples,good%"; null when none. */
function columnsArg(columns) {
  const list = (Array.isArray(columns) ? columns : []).filter(Boolean);
  return list.length ? quote(list.join(",")) : null;
}

/** Drop trailing nulls, since Excel does not need (or want) dangling commas. */
function trimTrailingNulls(args) {
  const out = [...args];
  while (out.length && out[out.length - 1] === null) out.pop();
  return out;
}

/**
 * =TB.HISTORY(dataset, tags, start, end, interval, aggregate, quality, lookback, novalue, columns)
 *
 * `novalue` is written only when it is "drop"; "flag" (keep the row, carry the
 * value, show it red) is the function's default.
 *
 * `tags` may be a string, an array of tag names, or a raw cell/range reference
 * such as "A1:A20" (pass `tagsAreReference: true` for that case, since a
 * reference must not be quoted).
 */
export function buildHistoryFormula({
  dataset,
  tags,
  tagsAreReference = false,
  start,
  end,
  interval,
  aggregate,
  quality,
  lookback,
  novalue,
  columns,
}) {
  const datasetArg = renderArg(dataset, { asRefIfLooksLikeOne: true });
  if (!datasetArg) throw new RangeError("A dataset is required.");

  const tagArg = tagsAreReference ? renderArg(tags, { asRefIfLooksLikeOne: true }) : tagsArg(tags);
  if (!tagArg) throw new RangeError("At least one tag is required.");

  const args = [
    datasetArg,
    tagArg,
    renderArg(start),
    renderArg(end),
    renderArg(interval),
    aggregate && aggregate !== "avg" ? renderArg(aggregate) : null,
    quality && quality !== "any" ? renderArg(quality) : null,
    renderArg(lookback),
    novalue === "drop" ? quote("drop") : null,
    columnsArg(columns),
  ];

  // A later positional argument being present forces every earlier one to be
  // rendered too, even when it is the default -- Excel has no named
  // arguments, so a gap cannot be left in the middle of the list.
  for (let i = args.length - 2; i >= 2; i--) {
    if (args[i + 1] !== null && args[i] === null) args[i] = '""';
  }

  return `=TB.HISTORY(${trimTrailingNulls(args).join(",")})`;
}

/**
 * =TB.SAMPLE(dataset, tags, start, end, interval, mode, quality, integers, lookback, columns)
 *
 * Defaults are omitted: interval 5 (minutes), mode "last", quality "good",
 * integers TRUE. A numeric-looking interval ("5", "15") is written as a bare
 * number, which TB.SAMPLE reads as minutes.
 */
export function buildSampleFormula({
  dataset,
  tags,
  tagsAreReference = false,
  start,
  end,
  interval,
  quality,
  integers = true,
  columns,
}) {
  const datasetArg = renderArg(dataset, { asRefIfLooksLikeOne: true });
  if (!datasetArg) throw new RangeError("A dataset is required.");

  const tagArg = tagsAreReference ? renderArg(tags, { asRefIfLooksLikeOne: true }) : tagsArg(tags);
  if (!tagArg) throw new RangeError("At least one tag is required.");

  const startArg = renderArg(start);
  if (!startArg) throw new RangeError("A start is required for sampling, for example -24h.");

  const intervalText = interval === null || interval === undefined ? "" : String(interval).trim();
  let intervalArg = null;
  if (/^\d+(\.\d+)?$/.test(intervalText)) intervalArg = Number(intervalText) === 5 ? null : intervalText;
  else if (intervalText !== "" && intervalText !== "5m") intervalArg = quote(intervalText);

  const args = [
    datasetArg,
    tagArg,
    startArg,
    renderArg(end),
    intervalArg,
    null, // mode: "last" is the default and only mode
    quality && quality !== "good" ? renderArg(quality) : null,
    integers === false ? "FALSE" : null,
    null, // lookback: default
    columnsArg(columns),
  ];

  for (let i = args.length - 2; i >= 3; i--) {
    if (args[i + 1] !== null && args[i] === null) args[i] = '""';
  }

  return `=TB.SAMPLE(${trimTrailingNulls(args).join(",")})`;
}

/** =TB.VALUE(dataset, tag, mode) */
export function buildValueFormula({ dataset, tag, mode }) {
  const datasetArg = renderArg(dataset, { asRefIfLooksLikeOne: true });
  if (!datasetArg) throw new RangeError("A dataset is required.");
  const tagArg = renderArg(tag, { asRefIfLooksLikeOne: true });
  if (!tagArg) throw new RangeError("A tag is required.");

  const args = [datasetArg, tagArg, mode && mode !== "raw" ? renderArg(mode) : null];
  return `=TB.VALUE(${trimTrailingNulls(args).join(",")})`;
}

/** =TB.DATASETS() */
export function buildDatasetsFormula() {
  return "=TB.DATASETS()";
}

/** =TB.TAGS(dataset, query) */
export function buildTagsFormula({ dataset, query }) {
  const datasetArg = renderArg(dataset, { asRefIfLooksLikeOne: true });
  if (!datasetArg) throw new RangeError("A dataset is required.");
  const args = [datasetArg, renderArg(query)];
  return `=TB.TAGS(${trimTrailingNulls(args).join(",")})`;
}
