# The Historian API contract this add-in depends on

Distilled from the Historian's own OpenAPI v1 document (`GET /api/help`,
machine-readable at `/api/help/swagger.json` or wherever your deployment
exposes it) plus the platform documentation (Historian REST API, Registering
Timebase apps as OAuth clients in Pulse, How to Subscribe to Live Tag Updates,
Historian dataset/tag query filters). If the API changes, this is the list to
re-check, with the file each row would need to change.

---

## Transport

| Fact | Consequence |
|---|---|
| HTTP on 4511, HTTPS on 4512 by default | `normalizeBaseUrl` in `historian.js` defaults to `https` on **4512** |
| Swagger UI at `/api/help` | Not called by the add-in; documented for the reader who wants to inspect a live deployment |
| Security scheme is a single Pulse-issued **Bearer JWT**, described in the OpenAPI doc as *"Paste ONLY the raw Pulse JWT (no 'Bearer ' prefix)"* | The add-in sends `Authorization: Bearer <token>` itself; that Swagger-UI instruction is about pasting into the *Swagger* auth box, not the wire format |
| Responses are heavily abbreviated JSON: `n` name, `d` description, `f` format, `u` units, `fl` fields, `t` overloaded as tag / timestamp / data type depending on context | `decodeTag` / `decodeDataset` expand every field once, so nothing downstream touches the short keys |

## Authentication (Pulse)

| Fact | Consequence |
|---|---|
| Token from `POST {pulse}/auth/token`, form-encoded `grant_type=client_credentials` | `auth.js`, `createTokenProvider` |
| The token's `aud` claim carries the client's own id first, then its configured audiences. **To let Excel call the Historian, the Pulse client used here must list the Historian's own client id in its audiences** | Named explicitly in the 401 error message, because it is the single most common first-run failure |
| Pulse seeds well-known clients (`Historian`, `Collector`, `Atlas`, ...) with a random GUID secret on first start, while the component ships an empty secret | `ErrPulseTokenFailed`-equivalent message in `auth.js` explains the rotate-the-secret step |
| Tokens last ~3600s by default; `client_credentials` carries no refresh token | `createTokenProvider` caches until near expiry, then re-acquires; concurrent callers collapse onto one request (Excel recalculates many cells at once) |
| No documented clock-skew tolerance | 401 message names clock drift as a candidate cause |

## Data

| Route | Notes | Consequence |
|---|---|---|
| `GET /api/datasets/{dataset}/data` | **No aggregation, no server-side row cap.** Batches tags via repeated `tagname=` query parameters | `buildDataQuery` repeats `tagname`; `align.js` does all bucketing/decimation client-side, and `buildGrid`'s `maxRows` refuses rather than silently truncating |
| Response shape | `{s, e, tl: [{t: Tag, d: TVQ[]}]}`. **The first point in `d` is the boundary sample -- the value immediately before the requested window**, described as "including leading boundary data point" | `splitBoundary` identifies it by timestamp, not by array position, so a tag with no earlier history (and therefore no boundary point) is never miscounted |
| No time bound given | Returns just the latest point per tag | Used by `TB.VALUE`, which reads the last element of `d` |
| Time parameters | `start`/`end` (ISO, ambiguous local-vs-UTC by suffix), `unixStart`/`unixEnd` (seconds), `relativeStart`/`relativeEnd` (the `[clamp][offset]` grammar) | `time.js` classifies an Excel date, an ISO string, or a relative expression and picks the right query parameter; pure offsets (`-24h`) are passed through for the *server* to resolve; **clamped** expressions (`d`, `d+12h`, `w`, `M`, `y`) are resolved in the workstation's time zone and sent as an absolute `start`/`end` with the UTC offset, because the Historian resolves a clamp on its own clock -- UTC in the Docker image -- so `d+12h` would otherwise mean 12:00 UTC |
| Relative grammar | Clamp letters `y M w d h m`; offset unit letters `y M d h m s` (note: `w` clamps but is **not** a valid offset unit); case matters (`M` = months, `m` = minutes) | `parseRelative` implements exactly this grammar; every example in the OpenAPI description has a matching test in `test/time.test.js` |
| `GET /api/datasets/{dataset}/data/stream` | SignalR-style live stream over a persistent HTTP connection, MQTT-style `+`/`#` wildcards | **Not used.** A custom function is a request/response call, not a subscription; there is no route for a spreadsheet formula to stay open indefinitely. Live values in this add-in mean "re-run `TB.VALUE`", not a push subscription |
| `POST` routes on `/data`, `/data/{tagname}`, `/status`, and `DELETE` on datasets/tags | Write and delete operations | **Not called anywhere.** This add-in is read-only by design |

## Datasets

| Route | Notes |
|---|---|
| `GET /api/datasets` | Bare array, no pagination. `pa`/`ps` are purge age (days) / size (GB); `ldt`/`lda` are late-data tolerance (ms) / age (days) |
| `GET /api/datasets/{dataset}` | Single dataset, same shape |
| `GET /api/datasets/{dataset}/exists` | Bare boolean body |

## Tags

| Route | Notes |
|---|---|
| `GET /api/datasets/{dataset}/tags` | **Historian 2.0 returns tags grouped by category**, `{"User": [...], "System": [...]}`, not a bare array; `tagListFromResponse` accepts both. Two independent filter parameters: `query` (the structured filter language) and `contains` (plain substring). Both are exposed on `client.listTags`, not merged into one |
| Tag filter language | `name:`, `description:`, `datatype:`, `type:`, `format:`, free text, `AND`/`OR`/`NOT`, wildcards (`*`), quoted phrases. Same grammar as the Historian's own Tag Grid filter box | Passed through verbatim as the `query` parameter; the add-in does not attempt to parse or validate it, since the server is the authority on the grammar |
| `u` (units) dictionary encoding | Single key, negative → prefix unit (`{"-1":"US$"}`); single key, positive → suffix unit (`{"1":"kW"}`); exactly keys `0` and `1` → boolean labels; three or more keys → enumeration | `format.js`'s `classifyUnits` implements exactly this, keyed on entry *count* before entry value, because a suffix-unit map and a two-key enum can otherwise look alike |
| `l` (last value snapshot) | `{t, v, q, vs}` -- `vs` is the pre-formatted display string, already resolving format and units | `TB.VALUE(..., "display")` re-derives this via `format.js` rather than reusing `vs`, so it also works for a historical point read via `TB.HISTORY`, not only the current value |

## Quality

| Fact | Consequence |
|---|---|
| "Quality represented as per OPC (e.g. 192 is Good)" | `quality.js` masks the top two bits (`0xC0`) rather than checking `=== 192`, so Good-with-substatus codes are still recognised as Good |
| The dataset **status** endpoint documents specific codes: 64 = collector heartbeat lost, 28 = collector shut down, 24 = collector lost its data source | Surfaced via `quality.describe()` for diagnostic use; not otherwise interpreted |

## Deliberate omissions

- **Live subscriptions** (`/data/stream`, `/tags/stream`, and the separate WebSocket hub documented in "Subscribe to Live Tag Updates") are not used. A worksheet formula cannot hold a socket open; "live" in this add-in means recalculating `TB.VALUE`, which Excel does on every recalculation pass or on a timer the user sets up with Excel's own volatile-function or macro mechanisms.
- **Writes.** `POST .../data`, `POST .../data/{tagname}`, `POST .../status`, and every `DELETE` are absent from this add-in on purpose. Read-only was the stated goal; adding write-back is a deliberate, separate decision with its own risk (a bad formula could then corrupt historized data) and is not something this add-in should grow into silently.
- **The i3X interoperability surface**, if the Historian also exposes one, uses a different envelope and is out of scope here; the OpenAPI document this add-in was built from is the native Historian API only.

## Verified live (Historian 2.0-gc4, Pulse 2.0-gc4, Excel for Mac 16.105)

- Points are logged on change only: a new point whenever a tag's value or its
  quality changes. There is no periodic logging, so the gap between two points
  means "unchanged", and a point count is a count of changes.

- No CORS: `OPTIONS` to the Historian and to Pulse returns 405 with no
  `Access-Control-*` headers. Same-origin hosting (or the `tools/serve.js`
  proxy) is required, on every platform.
- The data route's leading boundary point is real, and can be a status point
  with quality 64 and **no `v`** ("collector heartbeat lost"). `TB.SAMPLE`
  skips it under the default `good` filter and seeds from the lookback read.
- A missing value must be returned to Excel as `""`: `null` in a spilled array
  displays as 0.
- **What the value-less quality-64 points are.** The Historian writes one on
  every tag a collector feeds when its `CollectorStatusMonitor` marks that
  collector Disconnected (logged as "col_... -> Disconnected" after a 30 s
  SignalR timeout). They are connection events, not the quality of any sample:
  buffered (store-and-forward) data arrives afterwards as Good (192/193).
  On a laptop-hosted Docker stack they line up to the second with the Mac
  waking from sleep, when the frozen containers' connections time out. The
  add-in deliberately keeps showing them (row kept, value carried, red, counted
  as quality 64) so an outage stays visible; `novalue "drop"` hides them.
- Quality **193** is Good with an OPC limit sub-status (low bits `01`); it is
  classified Good by the `0xC0` mask, never as an issue.
- Interpolation, aggregation and interval parameters do not exist on
  `/data` (checked in `/api/v1.json`); sampling is client-side.

## Things to verify on the first live run

1. Whether the deployment's actual Historian sends CORS headers for the origin this add-in is hosted from. Hosting the add-in *on the Historian itself* (the chosen approach) sidesteps this by making every call same-origin; confirm that remains true if the add-in's static files end up served from a different port or a reverse proxy in front of the Historian.
2. Whether `redirect: "manual"` behaves identically across the WebView2 host (Windows), Safari's WebView (Mac), and Excel on the web. The fetch spec defines `opaqueredirect`, but coverage of `manual` redirect mode has historically varied slightly by engine; `historian.js` checks both `res.type === "opaqueredirect"` and a literal 3xx status to cover either behavior.
3. Whether this Historian's `/auth/token` requires `scope` at all -- the field is sent only when configured, per the Pulse OAuth client docs' minimal example, but confirm against your Pulse's actual client registration.
4. ~~Collector logging rate~~ -- settled: the Historian logs a point only when a tag's value or quality changes (report by exception), never on a timer. Between two points the tag is unchanged, which is exactly what carry-forward and the time-weighted average assume. Timestamps carry sub-millisecond precision (truncated to milliseconds by the add-in).
