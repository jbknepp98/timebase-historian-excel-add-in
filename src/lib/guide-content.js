// The getting-started guide, as data.
//
// src/taskpane/guide.js writes this into a workbook through Office.js, so the
// TB.* example formulas are entered by Excel itself (spilling correctly) and
// the guide is rebuilt from source rather than hand-maintained as a binary.
// test/guide.test.js checks every example formula against functions.json and
// the names defined on the Setup sheet, so the guide cannot drift from the
// functions it teaches.
//
// Block types, written top to bottom in column B (results go to column D on):
//   { title }                      large heading
//   { subtitle }                   one line under the title
//   { heading }                    section heading
//   { text }                       a paragraph (wrapped)
//   { bullets: [...] }             one bullet per row
//   { table: { header, rows, links? } }   a grid starting in column B
//   { input: { label, name, values, note } }  named input cells (Setup)
//   { example: { label, formula, at, note } } a live formula at `at` (e.g. "D6"),
//                                  with its text shown above it via FORMULATEXT

export const GUIDE_VERSION = "1.0";

/** Named cells on the Setup sheet that the examples refer to. */
export const GUIDE_NAMES = [
  "Guide_Dataset",
  "Guide_StateTags",
  "Guide_AnalogTag",
  "Guide_Start",
  "Guide_Query",
  "Guide_TagList",
];

const SHEET = {
  welcome: "Welcome",
  setup: "Setup",
  functions: "Functions",
  value: "1 Current value",
  tags: "2 Find tags",
  raw: "3 Raw history",
  agg: "4 Aggregates",
  sample: "5 State sampling",
  quality: "6 Quality",
  cells: "7 Tags from cells",
  time: "Time expressions",
  trouble: "Troubleshooting",
  more: "Learn more",
};

export const GUIDE_SHEETS = [
  {
    name: SHEET.welcome,
    blocks: [
      { title: "Timebase Historian for Excel" },
      { subtitle: `Getting-started guide ${GUIDE_VERSION} -- examples, reference and troubleshooting` },
      {
        text:
          "This add-in puts Timebase Historian data straight into worksheet formulas: current values, tag " +
          "searches, time-aligned history, interval aggregates and state sampling. It only reads: it never " +
          "writes a value, creates a tag or deletes anything on the Historian.",
      },
      { heading: "Start here" },
      {
        bullets: [
          "1. Open this workbook with the add-in installed. The Timebase Historian task pane opens on the right.",
          "2. In the task pane's Connect tab, check the Historian address and click Test connection.",
          "3. On the Setup sheet, replace the example dataset and tag names with your own.",
          "4. Work through the numbered example sheets. Every formula is live and updates from Setup.",
          "5. To add a formula to your own sheet, use the task pane's Insert tab.",
        ],
      },
      { heading: "In this guide" },
      {
        table: {
          header: ["Sheet", "What it covers"],
          rows: [
            [SHEET.setup, "Connecting, and the dataset and tag names every example uses"],
            [SHEET.functions, "All five worksheet functions and every argument"],
            [SHEET.value, "TB.VALUE -- the latest value, raw or formatted with units and labels"],
            [SHEET.tags, "TB.TAGS -- search a dataset by name, description, data type"],
            [SHEET.raw, "TB.HISTORY -- every change, carried forward from your start time"],
            [SHEET.agg, "TB.HISTORY with an interval -- time-weighted average, min, max, first, last"],
            [SHEET.sample, "TB.SAMPLE -- the value at each 5-minute mark, for state and enum tags"],
            [SHEET.quality, "Quality, status, sample count and % Good columns; red quality issues"],
            [SHEET.cells, "Driving a formula from a list of tag names in cells"],
            [SHEET.time, "Start and end: relative expressions, dates, the date picker"],
            [SHEET.trouble, "What an error or odd result means, and the fix"],
            [SHEET.more, "Documentation, tools and where to go next"],
          ],
          links: 0, // column 0 cells link to the sheet of that name
        },
      },
      { heading: "Four ideas that explain most results" },
      {
        bullets: [
          "Values carry forward. A tag holds its last known value until it changes, so a 12:00 start shows the value in force at 12:00 even if nothing changed then.",
          "avg is a time-weighted average: each value counts for as long as it was in force in the slice.",
          "Red means a quality issue: the value's latest point was not Good quality (for example a quality-64 'collector disconnected' note). The value itself is still the last known one.",
          "Timestamps are real Excel dates shown as mm/dd/yy hh:mm:ss.000, in this computer's time zone.",
        ],
      },
    ],
  },

  {
    name: SHEET.setup,
    blocks: [
      { title: "Setup" },
      { subtitle: "Connect once, then tell the examples which data to use" },
      { heading: "1. Connect" },
      {
        bullets: [
          "Recommended: the add-in server holds the Pulse credentials (node tools/set-credentials.js). The Connect tab then says 'Authentication is handled by the server' and Excel never stores a secret.",
          "Otherwise: in the Connect tab enter the Historian address, the Pulse client id and secret, then Test connection and Save.",
          "Test it here: the TB.DATASETS formula on the right should list your datasets.",
        ],
      },
      { example: { label: "Datasets on this Historian", formula: "=TB.DATASETS()", at: "F6" } },
      { heading: "2. Your data (replace the shaded cells)" },
      { text: "Every example in this guide reads these named cells. Type your own dataset and tags over the examples." },
      { input: { label: "Dataset", name: "Guide_Dataset", values: ["state-example"], note: "Exact dataset name, as TB.DATASETS lists it" } },
      {
        input: {
          label: "State tags (up to 3)",
          name: "Guide_StateTags",
          values: ["FL001.State", "FL002.State", "FL003.State"],
          note: "Discrete or enumerated tags: machine state, mode, status",
        },
      },
      { input: { label: "Analog tag", name: "Guide_AnalogTag", values: ["131-TT-001.PV"], note: "A continuous measurement: temperature, flow, pressure" } },
      { input: { label: "Default start", name: "Guide_Start", values: ["-24h"], note: "Relative (-24h, d, d+6h) or a date; see Time expressions" } },
      { heading: "3. Check the tags exist" },
      { text: "The second list on the right is TB.TAGS for your dataset. If a name above is not in it, the examples for that tag stay blank." },
      { example: { label: "Tags in your dataset", formula: "=TB.TAGS(Guide_Dataset)", at: "L6" } },
    ],
  },

  {
    name: SHEET.functions,
    blocks: [
      { title: "Worksheet functions" },
      { subtitle: "Arguments in [brackets] are optional. Leave a gap with two commas, e.g. =TB.HISTORY(ds, tag, \"-1h\",, \"5m\")" },
      {
        table: {
          header: ["Function", "Returns", "Use it for"],
          rows: [
            ["TB.DATASETS()", "Every dataset: name, tag count, size, data age", "Checking the connection; finding dataset names"],
            ["TB.TAGS(dataset, [query])", "Matching tags: name, description, type, unit, current value and time", "Finding tag names"],
            ["TB.VALUE(dataset, tag, [mode])", "One value: the latest", "A live KPI cell; mode \"display\" adds units and state labels"],
            ["TB.HISTORY(dataset, tags, [start], [end], [interval], [aggregate], [quality], [lookback], [novalue], [columns])", "Timestamp column, then one column per tag", "Raw history, or interval aggregates"],
            ["TB.SAMPLE(dataset, tags, start, [end], [interval], [mode], [quality], [integers], [lookback], [columns])", "Timestamp at each clock mark, then one column per tag", "State and enum tags: the value at 07:00, 07:05, ..."],
          ],
        },
      },
      { heading: "TB.HISTORY arguments" },
      {
        table: {
          header: ["Argument", "Default", "Meaning"],
          rows: [
            ["dataset", "(required)", "Dataset name"],
            ["tags", "(required)", "One name, an array {\"a\",\"b\"}, or a range of cells holding names"],
            ["start", "latest value only", "Window start. The first row is the start itself, carrying each tag's last known value"],
            ["end", "now", "Window end"],
            ["interval", "raw (every change)", "Slice width: \"30s\", \"5m\", \"1h\", \"1d\""],
            ["aggregate", "avg", "avg (time-weighted average, also \"twa\"), min, max, first, last, count, sum, range"],
            ["quality", "any", "any, notbad, or good. Filters points before anything else"],
            ["lookback", "\"7d\"", "How far back to look for the carried-in value if the Historian's prior value is missing"],
            ["novalue", "flag", "Points with no value (quality notes): flag keeps the row in red; drop leaves them out"],
            ["columns", "(none)", "Extra columns after each tag: quality, status, samples, good%"],
          ],
        },
      },
      { heading: "TB.SAMPLE arguments (where they differ)" },
      {
        table: {
          header: ["Argument", "Default", "Meaning"],
          rows: [
            ["start", "(required)", "Rows start at the first clock mark at or after it"],
            ["interval", "5 (minutes)", "A bare number is MINUTES here; or \"30s\", \"15m\", \"1h\", \"1d\""],
            ["mode", "last", "Sample-and-hold: the last value at or before each mark"],
            ["quality", "good", "Non-good points are skipped and the previous good value holds"],
            ["integers", "TRUE", "6.0 reads as 6. FALSE keeps the stored type"],
          ],
        },
      },
    ],
  },

  {
    name: SHEET.value,
    blocks: [
      { title: "1  Current value" },
      { subtitle: "TB.VALUE returns the latest value of one tag" },
      { text: "Raw (the default) returns the stored number, so charts and further formulas work on it." },
      { example: { label: "Raw", formula: "=TB.VALUE(Guide_Dataset, INDEX(Guide_StateTags, 1))", at: "F6" } },
      { text: "mode \"display\" returns the Historian's formatted text: units appended, and state codes turned into their labels (20 becomes Running)." },
      { example: { label: "Display", formula: "=TB.VALUE(Guide_Dataset, INDEX(Guide_StateTags, 1), \"display\")", at: "F10" } },
      { example: { label: "An analog tag, with its unit", formula: "=TB.VALUE(Guide_Dataset, Guide_AnalogTag, \"display\")", at: "F14" } },
      { heading: "Pointers" },
      {
        bullets: [
          "Recalculate (Formulas > Calculate Now) to refresh. TB.VALUE does not stream live updates.",
          "A tag with no unit or state labels returns the raw value in display mode too.",
        ],
      },
    ],
  },

  {
    name: SHEET.tags,
    blocks: [
      { title: "2  Find tags" },
      { subtitle: "TB.TAGS searches a dataset with the Historian's own filter language" },
      { text: "Type a query in the shaded cell; the list on the right follows it. Leave it blank to list every tag." },
      { input: { label: "Query", name: "Guide_Query", values: ["name:*State"], note: "Try: datatype:Double   or   description:Filler   or   NOT name:*State" } },
      { example: { label: "Matching tags", formula: "=TB.TAGS(Guide_Dataset, Guide_Query)", at: "F6" } },
      { heading: "Query language" },
      {
        table: {
          header: ["Query", "Finds"],
          rows: [
            ["name:FL001*", "Names starting FL001"],
            ["description:temperature", "Descriptions containing 'temperature'"],
            ["datatype:Double", "Floating-point tags"],
            ["name:*State AND description:Filler*", "Both conditions"],
            ["NOT name:Process.*", "Everything except the Historian's own process tags"],
          ],
        },
      },
      { text: "The task pane's Insert tab has the same search under + Add tags, with checkboxes to build a tag list." },
    ],
  },

  {
    name: SHEET.raw,
    blocks: [
      { title: "3  Raw history" },
      { subtitle: "Every change, carried forward from your start time" },
      {
        bullets: [
          "The first row is your start time itself, holding the value in force then -- even if nothing changed at that moment.",
          "After that, one row per change.",
          "A red value had a quality issue at that time; the value shown is the last known one carried through it.",
        ],
      },
      { example: { label: "State tag, from the Setup start", formula: "=TB.HISTORY(Guide_Dataset, INDEX(Guide_StateTags, 1), Guide_Start)", at: "F6" } },
      { heading: "Try" },
      {
        bullets: [
          "Change Guide_Start on Setup to d+6h (06:00 today) and watch the first row move.",
          "Several tags: replace INDEX(Guide_StateTags, 1) with Guide_StateTags. Every tag's changes interleave on one time axis.",
          "Analog tags change constantly: add an interval (sheet 4) before pulling a day of one.",
        ],
      },
    ],
  },

  {
    name: SHEET.agg,
    blocks: [
      { title: "4  Aggregates" },
      { subtitle: "An interval slices time; the aggregate decides what each slice shows" },
      {
        bullets: [
          "avg is time-weighted: a value held 4 minutes weighs four times one held 1 minute.",
          "Every slice starts from the value held at its start, so a slice with no change shows that value, not a blank.",
          "first is the value at the slice start; last is the value at the slice end.",
        ],
      },
      { example: { label: "Hourly time-weighted average", formula: "=TB.HISTORY(Guide_Dataset, Guide_AnalogTag, Guide_Start,, \"1h\", \"avg\")", at: "F6" } },
      { example: { label: "Hourly minimum", formula: "=TB.HISTORY(Guide_Dataset, Guide_AnalogTag, Guide_Start,, \"1h\", \"min\")", at: "I6" } },
      { example: { label: "Hourly maximum", formula: "=TB.HISTORY(Guide_Dataset, Guide_AnalogTag, Guide_Start,, \"1h\", \"max\")", at: "L6" } },
      { example: { label: "Samples per hour", formula: "=TB.HISTORY(Guide_Dataset, Guide_AnalogTag, Guide_Start,, \"1h\", \"count\")", at: "O6" } },
      { heading: "Pointers" },
      {
        bullets: [
          "The Historian has no server-side aggregation: the add-in fetches raw points and slices them. Wide ranges on fast tags take longer.",
          "A row limit (Connect tab) stops a formula that would produce too many rows; add or widen the interval.",
        ],
      },
    ],
  },

  {
    name: SHEET.sample,
    blocks: [
      { title: "5  State sampling" },
      { subtitle: "TB.SAMPLE: the value at each clock mark, for state and enum tags" },
      {
        bullets: [
          "Averaging a state tag invents states (4.6 is not a valid state). TB.SAMPLE reports, at 07:00, 07:05, ..., the last value at or before that instant.",
          "Rows land on clock marks, never 07:03:12. Hourly and daily rows land on this computer's local hour and midnight.",
          "Values are whole numbers by default (6.0 reads 6). Bad-quality points are skipped and the previous good value holds.",
          "A state shorter than the interval can start and end between two marks and never appear. That is what snapshot sampling means.",
        ],
      },
      { example: { label: "Every 5 minutes, all state tags", formula: "=TB.SAMPLE(Guide_Dataset, Guide_StateTags, Guide_Start)", at: "F6" } },
      { heading: "Try" },
      {
        bullets: [
          "15-minute rows: add ,,15 to the arguments -- =TB.SAMPLE(Guide_Dataset, Guide_StateTags, Guide_Start,, 15)",
          "In the Insert tab, choose Aggregate: Last value (step) to build these formulas.",
        ],
      },
    ],
  },

  {
    name: SHEET.quality,
    blocks: [
      { title: "6  Quality" },
      { subtitle: "Companion columns, red quality issues, and filtering" },
      {
        table: {
          header: ["Column", "Shows, over each row's range"],
          rows: [
            ["quality", "The worst OPC quality code (Bad < Uncertain < Good), counting the point in force at the range start"],
            ["status", "The same in words: Good / Uncertain: ... / Bad"],
            ["samples", "Points the Historian logged"],
            ["good%", "Share of those points with Good quality; blank when none were logged"],
          ],
        },
      },
      {
        example: {
          label: "Hourly last value with all four columns",
          formula: "=TB.HISTORY(Guide_Dataset, INDEX(Guide_StateTags, 1), Guide_Start,, \"1h\", \"last\",,,, \"quality,status,samples,good%\")",
          at: "F6",
        },
      },
      { heading: "Quality 64 notes" },
      {
        bullets: [
          "The Historian writes a value-less quality-64 point on every tag a collector feeds when it loses the connection to that collector. Data buffered by the collector arrives afterwards as Good.",
          "Such a point keeps its row, carries the last value and shows red. Tick 'Drop rows that have no value' in the Insert tab, or pass \"drop\" as novalue, to leave them out.",
          "quality \"good\" filters out every non-good point before alignment, so the quality columns then always read Good.",
        ],
      },
    ],
  },

  {
    name: SHEET.cells,
    blocks: [
      { title: "7  Tags from cells" },
      { subtitle: "Point a formula at a range of tag names; edit the cells to change the tags" },
      { text: "The shaded cells below copy the Setup state tags. Type any other tag names over them, or clear one -- blank cells are skipped." },
      { input: { label: "Tag list", name: "Guide_TagList", values: ["=INDEX(Guide_StateTags,1)", "=INDEX(Guide_StateTags,2)", "=INDEX(Guide_StateTags,3)"], note: "One tag name per cell" } },
      { example: { label: "Current state of each listed tag, every 5 minutes", formula: "=TB.SAMPLE(Guide_Dataset, Guide_TagList, \"-1h\")", at: "F6" } },
      { text: "In the Insert tab, choose Tags: Worksheet range, select the cells and click Use selected cells to build such a formula." },
    ],
  },

  {
    name: SHEET.time,
    blocks: [
      { title: "Time expressions" },
      { subtitle: "start and end accept a relative expression, a date, or a cell holding a date" },
      {
        table: {
          header: ["You write", "Means"],
          rows: [
            ["-24h", "24 hours ago"],
            ["-90m", "90 minutes ago (m = minutes, M = months)"],
            ["d", "Midnight today"],
            ["d+6h", "06:00 today"],
            ["d-1d", "Midnight yesterday"],
            ["h", "The top of the current hour"],
            ["w", "Midnight last Sunday"],
            ["M", "Midnight on the 1st of this month"],
            ["M-1M", "Midnight on the 1st of last month"],
            ["y", "Midnight on January 1st"],
            ["2026-10-22T07:00:00-04:00", "An exact instant, with its UTC offset"],
            ["A cell holding a date", "That date and time, on this computer's clock"],
          ],
        },
      },
      {
        bullets: [
          "'Today', 'this week' and so on are this computer's, worked out by the add-in -- not the Historian server's, which often runs on UTC.",
          "In the Insert tab, the calendar button next to Start and End picks a date and time and writes the exact instant with your UTC offset.",
          "Leave end blank for 'now'. Leave start blank in TB.HISTORY for just the latest value.",
        ],
      },
    ],
  },

  {
    name: SHEET.trouble,
    blocks: [
      { title: "Troubleshooting" },
      {
        table: {
          header: ["You see", "Likely cause", "Fix"],
          rows: [
            ["#VALUE! 'The Historian address is not set'", "Not connected yet", "Connect tab: enter the address (and credentials unless the server holds them), Test connection, Save"],
            ["#VALUE! 'Pulse did not issue a token' or a 401", "Wrong client id/secret, or the client's audiences do not include the Historian", "Fix the client in Pulse, then node tools/set-credentials.js (or the Connect tab)"],
            ["#VALUE! 'Could not reach the Historian'", "The add-in server (tools/serve.js) or the Historian is not running", "Start them; Test connection in the Connect tab"],
            ["#BUSY! that never clears", "The add-in's functions page did not start", "Close and reopen this workbook. Persisting: clear Excel's add-in cache (see README)"],
            ["#CALC!", "A list of tags passed where Excel expects one", "Use the latest add-in; pass tags as {\"a\",\"b\"} or a range"],
            ["#SPILL!", "Cells below or right of the formula are not empty", "Clear the cells in the way, or move the formula"],
            ["A row-limit error", "The result would exceed the Connect tab's row limit", "Add or widen the interval, or narrow the range"],
            ["A column blank all the way down", "Wrong tag name, or the tag has no data in the window", "Check the name with TB.TAGS (sheet 2)"],
            ["Values in red", "A quality issue at that time (often a quality-64 collector-disconnected note)", "Expected. Use novalue \"drop\" or quality \"good\" to hide them"],
            ["Bursts of quality 64 on a laptop", "The computer slept; the Historian marked the collector disconnected on waking", "Keep the host awake during demos (caffeinate -s on a Mac)"],
            ["'This add-in is no longer available'", "Excel cannot match this workbook to the installed add-in", "Reinstall with the current dist/manifest.xml; open the dist/Timebase Historian.xlsx launcher"],
            ["##########", "The column is too narrow for the timestamp", "Widen the column"],
          ],
        },
      },
    ],
  },

  {
    name: SHEET.more,
    blocks: [
      { title: "Learn more" },
      {
        table: {
          header: ["Where", "What"],
          rows: [
            ["README.md", "Install, sideload (Windows, Mac, web), every function with examples, troubleshooting"],
            ["docs/security.md", "Where the Pulse secret lives, and the server-held credentials mode"],
            ["docs/decisions.md", "Why the add-in is built the way it is (runtimes, bundling, no ribbon button)"],
            ["docs/api-contract.md", "Every Historian API behaviour the add-in depends on, verified live"],
            ["node tools/serve.js", "Runs the add-in and the same-origin proxy to the Historian and Pulse"],
            ["node tools/set-credentials.js", "Stores the Pulse client credentials for the server, out of Excel"],
            ["node tools/package.js <host>", "Builds dist/: the manifest, web files, this guide and the launcher"],
          ],
        },
      },
      { text: "Connect tab > Getting started > Add the getting-started guide recreates these sheets in any workbook." },
    ],
  },
];

/** Every example formula in the guide, with where it lives -- for the tests. */
export function guideFormulas() {
  const out = [];
  for (const sheet of GUIDE_SHEETS) {
    for (const b of sheet.blocks) {
      if (b.example) out.push({ sheet: sheet.name, ...b.example });
    }
  }
  return out;
}
