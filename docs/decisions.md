# Decisions made building this add-in, and why

Kept separate from the README so the README can stay a user-facing "how to use
this" document. This file is for the next person (or the next me) who wonders
"why is it built this way."

## Office.js + custom functions, not VBA

Both were considered. VBA (`.xlam`) needs no HTTPS hosting, has no CORS
exposure, and runs on Windows without any web technology at all -- a real
advantage in a locked-down plant network. Office.js was chosen because it
works on Windows, Mac and Excel on the web from one codebase, and because
**custom functions under the shared runtime get full CORS support**
(Microsoft's own documentation states this explicitly), which directly answers
the biggest open question at design time: whether the Historian's REST API
sends CORS headers a browser-hosted add-in could actually use.

## Separate custom-functions runtime, not a shared runtime

This add-in was first built on the shared runtime (task pane and functions in
one page, `<Runtime lifetime="long">`), for CORS and for passing settings
through a module-level singleton. Tested on Excel for Mac 16.105, a sideloaded
shared-runtime add-in never executes its custom functions: every formula sits
at `#BUSY!` and the function body is never called. The same functions in their
own runtime page run normally. So:

- `src/functions/functions.html` is the custom-functions page; the task pane
  is a separate page. `tools/check.js` refuses a `<Runtime>` element.
- CORS is solved by being same-origin instead: the Historian and Pulse send no
  CORS headers at all (both answer a preflight with 405), so the add-in must be
  served from the origin that also serves `/api` -- `tools/serve.js` proxies it.
- Settings are shared through `OfficeRuntime.storage`, which both runtimes can
  read; each function calls `loadSettings()` first. With server-held
  credentials (`docs/security.md`) the functions need no credentials at all.
- After switching an add-in between shared and separate runtimes, Excel keeps
  routing calls by its cached registration until its add-in cache
  (`~/Library/Containers/com.microsoft.Excel/Data/Library/Application
  Support/Microsoft/Office/16.0/Wef/`) is cleared.

## One classic script for the functions

Excel for Mac evaluates the custom-functions `<Script>` as a classic script, so
an ES module there is a syntax error and the runtime silently never starts.
`tools/bundle.js` flattens `functions.js` and its imports into
`functions.bundle.js` (built per request by `tools/serve.js`, written into
`dist/` by `tools/package.js`). Everything else stays build-free.

## No ribbon button (for now)

Since the 24-25 Sept 2026 Office update, a sideloaded add-in that declares
add-in commands is not listed in Home > Add-ins and fails to activate from a
workbook (`0x80043251`, OfficeDev/office-js#6973). Builds omit the
`DesktopFormFactor` block by default; the add-in starts from any workbook saved
with its task pane open (`dist/Timebase Historian.xlsx` is a blank one).
`node tools/package.js <host> --ribbon` restores the button once Microsoft
fixes the regression.

## Build-free: plain ES modules, no webpack

Every `.js` file the browser loads is exactly the file on disk -- no
transpilation, no bundling step, no `node_modules` dependency tree to audit or
go stale. This trades away TypeScript and JSX for a smaller, more auditable
surface, which fit an internal engineering tool better than the scaffolding a
production consumer add-in would want. `tools/check.js` and the Node test
suite are the substitute for a build step's usual safety net.

If this add-in later needs React, TypeScript, or hot-reload during
development, the standard path is to adopt the Yeoman generator's webpack
config and move these files in wholesale -- nothing here is written in a way
that resists that later.

## Client-side alignment, not a server aggregation feature

The Historian's `/data` route has no `interval` or `aggregate` parameter and no
row cap. A tag whose value changes every second logs roughly 2.6 million points a month (the Historian records a point on every value or quality change), and Excel's
hard ceiling is 1,048,576 rows per sheet. Two choices existed: ask users to
always narrow their own time range, or do the bucketing client-side in
`align.js` and let a formula ask for "hourly average over 90 days" directly.
The latter was built because it is what makes the connector actually useful
for trending rather than only for short spot-checks, and because getting the
half-open bucket boundary and the boundary-point handling right once, tested,
beats every user re-deriving it in a helper column.

## Client-credentials auth, with the risk documented rather than hidden

Client credentials are the only mode that works unattended, which matters for
a workbook someone opens once a week rather than a live session. The trade-off
-- a secret stored in the workbook -- is real and is written up plainly in
`docs/security.md` rather than glossed over. Interactive Pulse sign-in
(authorization-code + PKCE) was the documented alternative and was not built in
this version; it is the first thing to add if wide distribution of a workbook
becomes the actual requirement.

## Read-only, no write-back

`POST`/`DELETE` routes exist on the Historian API and are not called anywhere
in this add-in. The stated goal was reading Historian data into Excel; adding
write-back changes the risk profile substantially (a bad formula could then
corrupt historized process data) and deserves its own explicit decision later
rather than arriving as a side effect of "while I'm in here."

## Not built, and not attempted: a WebSocket/streaming "live" mode

The Historian documents both a `/data/stream` HTTP route and a separate
WebSocket hub for live tag updates. Neither is reachable from a worksheet
custom function, which is a request/response call by construction -- there is
no route for a formula to keep a connection open indefinitely. "Live" in this
add-in means "the user or Excel recalculates `TB.VALUE` again," which is a
real but different thing from a push subscription. A true live-tile experience
would need a task-pane-hosted WebSocket client writing into named ranges via
Office.js, which is a distinct feature with its own UI, not a custom function.
