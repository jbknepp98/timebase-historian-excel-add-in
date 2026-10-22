# Security notes

Read this before distributing a workbook that uses this add-in beyond your own
machine.

## Where the Pulse client secret lives

There are two modes. Use the first.

### Recommended: the server holds it

`node tools/set-credentials.js` stores the client id and secret in
`~/.config/timebase-excel/credentials.json` (mode 0600, outside the repository;
override with `TIMEBASE_EXCEL_CREDENTIALS`). `tools/serve.js` then signs in to
Pulse itself, caches the token until shortly before it expires, and attaches it
to every Historian request it forwards. The task pane detects this
(`/__server-auth`), hides its credential fields, and wipes any secret it saved
earlier. **No workbook, task pane or Excel setting holds the secret.**

Because that server now hands out authenticated Historian reads, it:

- accepts connections from this machine only (loopback; anything else is
  refused with 403), and
- forwards only `GET`/`HEAD` to the Historian, matching the add-in's read-only
  design, so a page that tricks the browser into a request cannot write data.

Anyone logged in to the machine as you can still read the credentials file and
use the proxy, so give the Pulse client read-only Historian access (below).
For a shared deployment, run the same proxy role as a service next to the
Historian, behind whatever sign-in your site requires.

### Fallback: the add-in holds it

With no credentials file, the task pane asks for the client id and secret and
saves them in `OfficeRuntime.storage`: Office's per-add-in store on this
machine, shared by the task pane and the custom-functions runtime. It survives
Excel restarts and is **not** written into the `.xlsx` file. (Earlier versions
stored it in the workbook's document settings, which travel inside the file;
that is no longer done.) Each user on each machine enters it once.

## Practical mitigations, in order of effectiveness

1. **Scope the Pulse client narrowly.** Register a client whose only audience
   is the Historian, with no write-capable scope beyond what the Historian API
   itself needs (this API is read-only from the add-in's side regardless, but
   a narrowly scoped client also limits blast radius if the secret leaks via
   some other channel). See `docs/api-contract.md` for the audience
   requirement.
2. **Rotate the secret** if a workbook containing it is ever shared outside
   its intended audience, the same way you would rotate any other leaked
   credential. Pulse's UI supports secret rotation; the old secret stops
   working immediately.
3. **Prefer "not for wide distribution" workbooks for this connection method.**
   A personal analysis workbook that stays in one person's OneDrive is a very
   different risk than a template forwarded to fifty people.
4. **If broad distribution is the actual goal**, this add-in's client-
   credentials mode is the wrong fit. The natural alternative is an interactive
   Pulse sign-in (authorization-code + PKCE), which stores no secret in the
   workbook at all and attributes each read to the signed-in user. That flow
   was considered and intentionally not built in this version -- see
   `docs/decisions.md` -- because it needs a registered redirect URI and is
   fragile for unattended Power BI Service-style scheduled refresh, which does
   not apply to an interactive Excel session the way it does a service. If your
   distribution need changes, that is the first thing to build next.
5. **An on-premises data gateway or a thin backend proxy** that holds the
   secret server-side and issues short-lived, narrowly-scoped tokens to Excel
   clients is the strongest option, at the cost of another service to run and
   operate. Reasonable for an IT-managed rollout; overkill for one engineer's
   trending workbook.

## Other things worth knowing

- The add-in never writes to the Historian: no tag is created, no value is
  posted, no dataset is deleted. Confirm this stays true if you extend
  `functions.js` -- there is no code path today that calls any `POST` or
  `DELETE` route, and that absence is intentional (see `docs/api-contract.md`,
  "Deliberate omissions").
- `credentials: "omit"` is set on every request this add-in makes, on purpose:
  neither the Historian call nor the Pulse token call ever sends an ambient
  browser cookie. Only the bearer token this add-in explicitly attaches is
  used for authentication.
- Because the add-in is hosted **on the Historian itself** (the chosen
  deployment target), the task pane and every custom function call the
  Historian same-origin. That is good for CORS but means anyone who can reach
  the Historian's web port can also load this add-in's static files --
  ordinary for an internal engineering tool, worth knowing if the Historian is
  ever exposed more broadly than intended.
