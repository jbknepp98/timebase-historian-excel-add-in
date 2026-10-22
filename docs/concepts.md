# Concepts: how the add-in turns history into rows

The Historian stores each tag as its own series of points -- a timestamp, a
value and an OPC quality -- and logs a point only when the value or the
quality changes. Between two points the tag is unchanged. A worksheet needs one
timestamp column and one column per tag. These are
the rules the add-in follows to get from one to the other. They explain almost
every "why does it show that?".

## Carry-forward: a tag holds its value until it changes

A tag's value at any instant is the **last value logged at or before that
instant**. Nothing is interpolated and nothing is invented.

- **Raw history** (`TB.HISTORY` with a start and no interval): the first row is
  **the start time itself**, holding each tag's value in force then, even if
  nothing changed at that moment. Then one row per change of any tag; the
  other tags carry their values into that row.
- **Interval slices**: every slice starts from the value held at its start. A
  slice in which nothing changed shows the held value, not a blank.
- **`TB.SAMPLE`**: each clock mark shows the last value at or before it.

**Where the starting value comes from.** The Historian returns, with every
query, the last point before the window. If that point is missing (the tag
has no earlier history) or unusable (it failed the quality filter, or it is a
quality-64 note with no value), the add-in makes one more request for just
those tags, looking back up to `lookback` (default 7 days) for the last real
value. If there is none, the cell is blank -- never 0.

## Aggregates, and the time-weighted average

With an `interval`, each tag's points are cut into half-open slices
`[start, start + interval)`, labelled by their start. A point exactly on a
slice boundary belongs to the slice it starts.

| Aggregate | Per slice |
|---|---|
| `avg` (default; also `twa`) | **Time-weighted average**: each value counts for as long as it was in force in the slice. Held 10 for 3 min then 20 for 2 min: (10x3 + 20x2) / 5 = 14 |
| `min`, `max`, `range` | Over the value held at the slice start and every change in the slice |
| `first` | The value in force at the slice start |
| `last` | The value in force at the slice end |
| `count`, `sum` | Of the points actually logged in the slice |

Time-weighting stops at the window end (nothing is known beyond it). Text and
missing values carry no weight in `avg`, `min`, `max`, `sum`; they take part in
`first` and `last`.

Slices start at the window start: a `"-24h"` window at 14:49 gives 14:49,
15:49, ... Use `TB.SAMPLE` for rows on clock marks.

## State sampling (`TB.SAMPLE`)

For discrete and enumerated tags -- machine state, mode, status -- averaging
invents states that do not exist (4.6 is not a state) and interpolation
pretends a state ramps. `TB.SAMPLE` reports, at each clock-aligned mark, the
last value at or before it (sample-and-hold):

- **Clock-aligned rows.** Under an hour, marks are multiples of the interval
  (07:00, 07:05, ...), which match local clock marks in whole- and half-hour
  time zones. An hour or more aligns to the **local** clock, so daily rows land
  on local midnight, across daylight-saving changes too.
- **Whole numbers.** `6.0` reads `6` unless `integers` is `FALSE`.
- **Good quality only** by default: a bad point is skipped and the previous
  good value holds.
- **Short states can be missed.** A state that starts and ends between two
  marks (a one-second fault at 07:17:27) never appears in a 5-minute table.
  That is what snapshot sampling means; use raw history to see every change.

## Quality

Every point carries an OPC quality code. The add-in classifies by the code's
top two bits, so sub-status variants count the same as their class:

| Class | Codes | Examples |
|---|---|---|
| Good | 192-255 | 192; 193 (Good, low-limited) |
| Uncertain | 64-127 | 64 |
| Bad | 0-63 | 0; 24, 28 (collector status codes) |

- **`quality` argument** (`any`, `notbad`, `good`) filters points **before**
  anything else, so a filtered-out point never becomes a carried value.
- **Red values**: a value whose latest point is not Good is returned as a
  number with a `[Red]` number format -- it reads red but stays a real number.
  It clears at the next good point. For a slice, red means a non-good point
  inside it or in force at its start.
- **Companion columns** (`columns` argument; *Extra columns per tag* in the
  task pane), over each row's range -- the slice, the span since the previous
  clock mark, or the span since the previous row:

  | Column | Shows |
  |---|---|
  | `quality` | The **worst** code in the range (Bad < Uncertain < Good; ties to the lower code), counting the point in force at the range start |
  | `status` | The same in words |
  | `samples` | Points logged in the range -- each one a value or quality change |
  | `good%` | Share of those that were Good; blank when none |

### Points with no value: dropped sessions and deleted tags

Most points carry a value. Some carry only a timestamp and a quality, with a
**null** value. The Historian writes these to mark a break in a tag's data,
when there is no real value to record:

- **A dropped session.** When the Historian loses its connection to a
  collector, it writes a null point on every tag that collector feeds (quality
  64, below). The tag's data resumes with the next real point, typically the
  collector's buffered data once it reconnects.
- **A deleted tag.** When a tag is deleted, or stops being collected, its
  series ends with a null point. Nothing follows it.

Because the Historian logs only on change, a null point is the only way the
data can say "from here, there is no value" rather than "unchanged". The
add-in represents it like this:

| Where | A null point shows as |
|---|---|
| Raw history | A row at the null point's timestamp, holding the **last real value, in red** (a quality issue). Following rows stay red until the next real point; for a deleted tag, that is every row after it |
| Interval slices | The slice is red. The null point adds nothing to `avg`, `min`, `max` or `sum`, and does not change `first`/`last` |
| `TB.SAMPLE` | Skipped by the default `good` filter. With quality `any`, the last real value carries through it, in red |
| Quality columns | Counted in `samples`, not as Good in `good%`; its quality (64 for a dropped session) can be the range's worst |
| A null before the window start | Never used as the starting value: the add-in looks further back for the last real one, and that first row is red |
| A tag with no real value at all | A blank cell -- never 0 |

So **red values are where to look for these breaks**: a run of red rows after a
dropped session clears when data resumes; red to the end of the table usually
means the tag was deleted or is no longer collected. To leave the null points
out entirely, use `novalue "drop"` (*Drop rows that have no value* in the
Insert tab), or quality `"good"`.

### Quality 64: collector disconnected

The Historian writes a **value-less** point with quality 64 on every tag a
collector feeds when it loses its connection to that collector (its
`CollectorStatusMonitor` logs `-> Disconnected`). It is a connection event,
not the quality of a sample: data the collector buffered while disconnected
arrives afterwards as Good.

The add-in keeps showing these by default -- the row stays, the last value
carries through it in red, and it counts as 64 in the quality columns -- so an
outage stays visible. To leave them out, tick **Drop rows that have no value**
in the Insert tab (formulas then end with `"drop"` as `novalue`), or use
quality `"good"`.

On a laptop running Timebase in Docker, they arrive in bursts each time the
computer wakes from sleep: the frozen containers' connections time out and
the Historian marks the collector disconnected. Keep the host awake during
demos (`caffeinate -s` on a Mac).

## Time

- **Timestamps** are real Excel dates, shown as `mm/dd/yy hh:mm:ss.000`, in
  this computer's time zone. Sorting, charts and date arithmetic work on them.
- **Relative expressions.** A pure offset (`-24h`, `-90m`) means the same
  instant everywhere and is resolved by the Historian. An expression anchored
  to a calendar boundary -- `d` (midnight), `d+6h`, `w` (Sunday), `M` (the
  1st), `y` -- means *this computer's* today, so the add-in resolves it locally
  and sends the exact instant with its UTC offset. (The Historian would
  otherwise resolve it on its own clock, often UTC.)
- **Dates you pick** in the Insert tab are written with your UTC offset, e.g.
  `2026-10-22T07:00:00-04:00`, for the same reason.

## Limits

- The Historian has no server-side aggregation and no row cap. The add-in
  fetches raw points and does the work itself, so a wide range on a fast tag
  takes time; the **row limit** in the Connect tab stops a formula that would
  produce more rows than you allow.
- Formulas are request/response: values update when Excel recalculates, not
  by streaming.
- The add-in only reads. It never writes values or changes tags.
