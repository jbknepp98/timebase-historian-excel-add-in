# Timebase Historian for Excel

An Excel add-in that puts [Timebase](https://www.flow-software.com/timebase)
Historian data straight into worksheet formulas: current values, tag search,
time-aligned history, time-weighted aggregates, state sampling and quality
columns. It only reads -- it never writes a value, creates a tag or deletes
anything on the Historian.

```
=TB.HISTORY("state-example", "131-TT-001.PV", "-24h", , "1h", "avg")
=TB.SAMPLE("state-example", {"FL001.State","FL002.State"}, "d+6h")
=TB.VALUE("state-example", "FL001.State", "display")          -> Running
```

## What you get

| Function | Returns |
|---|---|
| `TB.DATASETS()` | Every dataset: name, tag count, size, data age |
| `TB.TAGS(dataset, [query])` | Matching tags, with the Historian's filter language |
| `TB.VALUE(dataset, tag, [mode])` | The latest value; `"display"` adds units and state labels |
| `TB.HISTORY(dataset, tags, [start], [end], [interval], [aggregate], [quality], [lookback], [novalue], [columns])` | Raw history carried forward from your start, or interval slices (`avg` is time-weighted) |
| `TB.SAMPLE(dataset, tags, start, [end], [interval], [mode], [quality], [integers], [lookback], [columns])` | The value at each clock mark (07:00, 07:05, ...), for state and enum tags |

And a task pane:

- **Connect** -- the Historian address and sign-in, and **Getting started**,
  which adds a guide workbook of live examples to any workbook.
- **Insert** -- build a formula without typing it: search and pick tags (or
  point at a range of cells), date/time pickers, aggregates, quality columns.

Values carry forward (a 12:00 start shows the value in force at 12:00),
quality issues show red, and timestamps are real Excel dates shown as
`mm/dd/yy hh:mm:ss.000`. [Concepts](docs/concepts.md) explains the rules.

## Quick start

You need Timebase Historian and Pulse running, Node.js 22+, and Microsoft 365
Excel. [Installation](docs/installation.md) has every step in detail,
including creating the Pulse client if you have never used Pulse.

```bash
git clone https://github.com/jbknepp98/timebase-historian-excel-add-in.git
cd timebase-historian-excel-add-in
npx --yes office-addin-dev-certs install                      # trust a localhost certificate (once)
mkdir -p .certs && docker cp <historian-container>:/historian/certificates/pulse-ca.cer .certs/
node tools/set-credentials.js                                 # the Pulse client id + secret for Excel
npm run build && npm start                                    # leave running
npm run sideload:mac                                          # Mac; Windows: Upload My Add-in, dist\manifest.xml
```

Then open `dist/Timebase Historian Guide.xlsx`.

## Documentation

| | |
|---|---|
| [Installation](docs/installation.md) | From a Timebase install to working formulas, including the Pulse client, certificates, Mac/Windows/web, upgrading |
| [User guide](docs/user-guide.md) | Every function and argument, worked examples, time expressions, troubleshooting (the same content as the guide workbook) |
| [Concepts](docs/concepts.md) | Carry-forward, time-weighted averages, state sampling, quality and red values, quality-64 notes, time zones |
| [Security](docs/security.md) | Where the Pulse secret lives; the server-held credentials mode |
| [Development](docs/development.md) | Layout, tests, debugging in Excel, rebuilding the guide, releasing |
| [Design decisions](docs/decisions.md) | Why it is built this way |
| [API contract](docs/api-contract.md) | The Historian API behaviour the add-in relies on, verified live |

## How it works

```
Excel  --->  https://localhost:3443  (tools/serve.js)  --->  Historian  :4512
                                                       --->  Pulse      :4542
```

`tools/serve.js` serves the add-in and forwards Excel's requests from one
origin, because the Historian and Pulse send no CORS headers. With server-held
credentials it also signs in to Pulse, so Excel never stores a secret. The
`TB.*` functions run in Excel's custom-functions runtime; the task pane is
separate. All the logic is plain JavaScript in `src/lib/`, unit-tested with no
Office dependency.

## Status

Version 1.0.3. Exercised live on **Excel for Mac 16.105** against Timebase
Historian and Pulse **2.0-gc4**: every function, the task pane, multi-tag
formulas, quality columns, the guide workbook. 210 Node tests cover the logic.

Known limits:

- **No ribbon button**, because of a Microsoft sideloading regression
  (OfficeDev/office-js#6973, from the 24-25 Sept 2026 Office update). The add-in starts
  from a workbook linked to it -- the guide, `dist/Timebase Historian.xlsx`,
  or any workbook saved with the task pane open. `--ribbon` restores the button.
- **Windows and Excel on the web** are supported by the same build but have
  not been exercised as thoroughly as Mac.
- **`TB.HISTORY` interval slices** start at the window start, not on clock
  marks; use `TB.SAMPLE` for clock-aligned rows.
