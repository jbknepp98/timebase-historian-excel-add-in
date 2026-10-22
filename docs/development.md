# Development

Everything here runs on plain Node.js 22+, with no npm dependencies. The files
Excel loads are the files on disk, except the custom-functions bundle (below).

## Commands

| Command | Does |
|---|---|
| `npm test` | Every unit test (`test/*.test.js`): time handling, alignment, sampling, quality columns, formula text, settings, credentials, the guide's formulas, the functions bundle |
| `npm run check` | `tools/check.js` (manifest and `functions.json` consistency) and the user-guide freshness check |
| `npm start` | `tools/serve.js`: serves the add-in on https://localhost:3443 and proxies `/api` and `/auth` |
| `npm run build` | `tools/package.js https://localhost:3443`: writes `dist/` |
| `npm run package -- <host> [--ribbon]` | Build for another host; `--ribbon` restores the Home-tab button |
| `npm run sideload:mac` | Install `dist/manifest.xml` into Excel for Mac |
| `npm run credentials` | Store the Pulse client credentials the server uses |
| `npm run docs` | Regenerate `docs/user-guide.md` from the guide content |
| `npm run icons` | Resize `assets/icon-source.png` into the three icons (macOS) |

Run `npm test && npm run check` before every commit.

## Layout

| Path | Purpose |
|---|---|
| `src/lib/` | All the logic, with no Office.js dependency, fully unit-tested: `time.js` (Excel dates, relative expressions, local resolution), `align.js` (carry-forward alignment, slicing, time-weighted average), `sample.js` (`TB.SAMPLE`), `columns.js` (quality columns), `quality.js`, `historian.js` (the REST client), `auth.js` (Pulse tokens), `settings.js`, `formula.js` (formula text the task pane inserts), `picker.js`, `format.js`, `args.js`, `errors.js`, `guide-content.js` (the guide, as data) |
| `src/functions/` | The `TB.*` functions: `functions.js` (entry point), `functions.json` (Excel's metadata), `functions.html` (the functions runtime page) |
| `src/taskpane/` | Connect and Insert UI, and `guide.js`, which writes the guide into a workbook |
| `src/runtime/` | `session.js` (one client per runtime, settings from `OfficeRuntime.storage`), `devlog.js` (localhost console relay) |
| `manifest.xml` | Manifest template; `tools/package.js` substitutes the host |
| `tools/` | `serve.js`, `package.js`, `bundle.js`, `check.js`, `make-docs.js`, `set-credentials.js`, `credentials.js`, `sideload-mac.js`, `make-icons.js`, `launcher-template/` |
| `guide/Timebase Historian Guide.xlsx` | The shipped guide workbook (see below) |
| `docs/` | This documentation |

## How it fits together

- **Two runtimes.** The task pane (`taskpane.html`) and the custom functions
  (`functions.html`) run separately. A shared runtime does not execute
  sideloaded custom functions on Excel for Mac 16.105; `tools/check.js` refuses
  a manifest that reintroduces one. Settings pass between the two through
  `OfficeRuntime.storage`; every function calls `loadSettings()` first.
- **The functions bundle.** Excel for Mac runs the manifest's functions
  `<Script>` as a classic script, so `tools/bundle.js` flattens `functions.js`
  and its imports into `functions.bundle.js`. `tools/serve.js` builds it on
  every request; `tools/package.js` writes it into `dist/`. The task pane loads
  plain ES modules.
- **Same origin.** The Historian and Pulse send no CORS headers, so everything
  -- add-in files, `/api`, `/auth` -- is served from one origin by
  `tools/serve.js`. It accepts loopback connections only, forwards only
  GET/HEAD to the Historian, and, with a credentials file, signs in to Pulse
  itself.
- **Returned cell values.** Timestamps, red quality values and `% Good` are
  `FormattedNumber` values (real numbers with their own number format); blanks
  are `""`, because Excel shows `null` in a spilled array as 0. This needs
  `allowCustomDataForDataTypeAny: true` in `functions.json`.
- **List parameters are matrices.** `tags` and `columns` are declared
  `"dimensionality": "matrix"`. Declared scalar, Excel would call the function
  once per array element and show `#CALC!`; `tools/check.js` enforces this.

See [decisions.md](decisions.md) for the reasoning and
[api-contract.md](api-contract.md) for the Historian behaviour relied on.

## Changing things

- **A function's arguments**: edit `functions.js` (JSDoc and code) and
  `functions.json` together; `tools/check.js` compares them. Update
  `formula.js` if the task pane builds that function, and the guide content.
- **The guide**: edit `src/lib/guide-content.js`, run `npm run docs`, and
  rebuild the workbook (below). `test/guide.test.js` checks every example
  formula against `functions.json` and the guide's named cells.
- **The manifest**: run `npm run build` and validate with
  `npx office-addin-manifest validate dist/manifest.xml`. Microsoft's validator
  catches schema errors that Excel otherwise ignores silently -- including `--`
  inside an XML comment, which makes the file invalid XML.

## Rebuilding the guide workbook

1. `npm run build`, `npm start`, `npm run sideload:mac`, restart Excel.
2. Open a copy of `dist/Timebase Historian.xlsx`.
3. Task pane > **Connect** > **Getting started** > **Add the getting-started guide**.
4. Check every example sheet calculates, then save as
   `guide/Timebase Historian Guide.xlsx` **with the task pane open** (that is
   what keeps the workbook linked to the add-in).

`tools/package.js` copies it to `dist/` and re-stamps its add-in link with the
current id and version on every build.

## Debugging in Excel

- **The console relay.** When the add-in is served from localhost,
  `src/runtime/devlog.js` posts `console.*` output and uncaught errors from
  both runtimes to `tools/serve.js`, which prints them as `addin [level] ...`.
  On a deployed host it does nothing.
- **Request log.** `tools/serve.js` logs every request's status and path (never
  bodies or query strings).
- **Excel's diagnostic logs** (Mac):
  `~/Library/Containers/com.microsoft.Excel/Data/Library/Logs/Diagnostics/EXCEL/`.
  `Office.Extensibility.Sandbox.CreateOsfControlV2` failing with
  `0x80043251` is the sideloading regression; `InstallFunctionsFromMetadata`
  failures point at `functions.json`.
- **Stale behaviour.** Excel caches add-in manifests and function
  registrations per add-in id and version. After changing `functions.json`, or
  between runtime models, quit Excel and move aside the contents of
  `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`.
  A version bump in `manifest.xml` also forces a refresh, but every workbook
  link names an exact version -- rebuild so the guide and launcher are
  re-stamped.
- **Spilled formulas from AppleScript.** Set `formula2`, not `formula`; the
  latter enters a legacy implicit-intersection formula (`=@TB...`).

## Releasing

1. Bump `<Version>` in `manifest.xml` and `version` in `package.json`.
2. `npm test && npm run check`.
3. `npm run build`; validate the manifest.
4. Rebuild the guide workbook if its content changed.
5. Commit, tag `vX.Y.Z`, push, and create a GitHub release from the tag. Users
   install from the repository (it includes the server and tools), so the
   source archive GitHub attaches to the release is the download; `dist/` is
   built locally by each user for their own host.
