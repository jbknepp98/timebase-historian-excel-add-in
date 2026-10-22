# Installation

This guide takes you from a working Timebase install (Historian and Pulse
running) to Excel formulas returning your data. It assumes you have never used
Pulse before. Allow about 20 minutes.

The add-in has two parts:

- **The add-in server** (`tools/serve.js`, Node.js) runs on the same computer
  as Excel. It serves the add-in to Excel and forwards Excel's requests to the
  Historian and to Pulse. This is required: the Historian and Pulse do not
  accept browser requests from other web addresses (they send no CORS
  headers), so Excel has to reach them through one address.
- **The add-in itself**, loaded into Excel from a manifest file.

```
Excel  --->  https://localhost:3443  (tools/serve.js)  --->  Historian  https://<host>:4512
                                                       --->  Pulse      https://<host>:4542
```

## What you need

| | Version | Check with |
|---|---|---|
| Timebase Historian and Pulse | 2.0 or later, running | Open `https://<host>:4542` in a browser |
| Node.js | 22 or later | `node --version` |
| Microsoft Excel | Microsoft 365: Excel for Mac 16.105+, Excel for Windows, or Excel on the web | Excel > About Excel |
| A Pulse administrator sign-in | | You will use it in step 2 |

Excel must be signed in with a licensed Microsoft 365 account. An unlicensed or
reduced-functionality Excel will not load add-ins.

## 1. Get the add-in

```bash
git clone https://github.com/jbknepp98/timebase-historian-excel-add-in.git
cd timebase-historian-excel-add-in
npm test
```

There are no dependencies to install. `npm test` should finish with every test
passing.

## 2. Create a Pulse client for Excel

Excel signs in to the Historian through Pulse, Timebase's identity service,
using an OAuth *client*: an id and a secret that Pulse issues. You will create
one named `Excel`.

### 2.1 Sign in to Pulse

1. In a browser, open your Pulse address -- by default **`https://<host>:4542`**
   (`https://localhost:4542` if Timebase runs on this computer).
2. If the browser warns that the certificate is not trusted, see
   [Trusting the Timebase certificate](#trusting-the-timebase-certificate) below,
   or continue past the warning for now.
3. Sign in. A new Pulse install uses **Administrator** / **timebase**. If
   nobody has changed that yet, change it now (Pulse prompts you, or use
   **Identity > Users**): the default password is public.

### 2.2 Add the client

1. In the left-hand menu, under **Identity**, click **Clients**. You will see
   clients Pulse created for the other Timebase components -- `Historian`,
   `Collector`, `Atlas` and so on. Leave those alone.
2. Click **Add client**. The **New client** form opens.
3. Fill it in:

   | Field | Value | Why |
   |---|---|---|
   | **Client Id** | `Excel` | Any name works; you cannot change it later |
   | **Grant Types** | **Client Credentials** only | Excel signs in as a service, with no user login |
   | **Scopes** | leave empty | Not needed |
   | **Redirect URIs** | leave empty | Only for interactive logins |
   | **Audiences** | `Historian` | **Essential.** See below |
   | **Token Expiry (s)** | leave at `3600` | One hour; the add-in renews automatically |

4. Click **Create**.

**About Audiences.** A Pulse token only works on the services listed in the
client's Audiences. The Historian accepts a token only if its own client id --
normally `Historian` -- is in that list. If you leave Audiences empty, Excel will
get a token but every Historian request will fail with *401 Unauthorized*. If
your Historian was set up with a different client id, use that: it is the
`Pulse > Auth > ClientId` value in the Historian's `settings.config`, and the
Historian's row in the Clients list.

### 2.3 Copy the client secret

1. In the Clients list, click the **Excel** row to open it.
2. Under **Security > Client secret**, click **Reveal** to show it, or **Copy**
   to copy it to the clipboard. Only administrators can see it.
3. Keep it somewhere safe for the next step. Treat it like a password: anyone
   with the id and secret can read your Historian data.

**Lost it, or think it leaked?** Open the client and click **Rotate secret**,
then confirm. Pulse generates a new secret and the old one stops working
immediately -- repeat step 4 with the new one.

### 2.4 Check it works (optional)

```bash
curl -sk -X POST https://localhost:4542/auth/token \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "grant_type=client_credentials&client_id=Excel&client_secret=PASTE-SECRET-HERE"
```

A JSON reply containing `"access_token"` means the client is right. `invalid_client`
or *Client 'Excel' could not be found* means the id or secret is wrong.

## 3. Trust the certificates

Two certificates are involved, and both must be trusted.

### 3.1 The add-in server's localhost certificate

Excel only loads add-ins over HTTPS, from a certificate the operating system
trusts. Microsoft's tool creates and trusts one for `localhost` (it asks for
your password once):

```bash
npx --yes office-addin-dev-certs install
```

### 3.2 The Timebase certificate

The Historian and Pulse use certificates signed by Pulse's own root
certificate authority. The add-in server needs that CA to check it is talking
to the real Historian. Copy it from the Historian into the project's `.certs`
folder:

- **Docker** (replace the container name with yours -- `docker ps` lists them):

  ```bash
  mkdir -p .certs
  docker cp <historian-container>:/historian/certificates/pulse-ca.cer .certs/pulse-ca.cer
  ```

- **Not Docker:** copy `pulse-ca.cer` from the Historian's `certificates`
  folder into `.certs/`.

Pulse also serves it at `https://<host>:4542/api/ca` (no sign-in needed).
`.certs/` is ignored by git.

#### Trusting the Timebase certificate

To stop browser warnings for Pulse and the Historian, add the same
`pulse-ca.cer` to your operating system's trusted roots (macOS: double-click
it, add to the System keychain, set it to *Always Trust*; Windows: double-click,
*Install Certificate* > *Local Machine* > *Trusted Root Certification
Authorities*). Timebase components may already have done this for you.

## 4. Give the add-in server the client credentials

```bash
node tools/set-credentials.js
```

Enter `Excel` as the client id and paste the secret from step 2.3 (it is not
shown as you type). They are saved to
`~/.config/timebase-excel/credentials.json`, readable only by you, outside the
project. Excel never sees or stores the secret; the server signs in to Pulse
and adds the token to every Historian request.

## 5. Build and start

```bash
npm run build      # = node tools/package.js https://localhost:3443
npm start          # = node tools/serve.js
```

`npm start` must keep running whenever you use the add-in. Its first lines
should say:

```
  /api/*  ->  https://localhost:4512
  /auth/*  ->  https://localhost:4542
  upstream CA: .../.certs/pulse-ca.cer
  Pulse auth:  server-held, client "Excel" from ~/.config/timebase-excel/credentials.json
```

**Historian or Pulse on another machine?** Point the server at them:

```bash
HISTORIAN_URL=https://historian.example.com:4512 PULSE_URL=https://historian.example.com:4542 npm start
```

(Windows PowerShell: `$env:HISTORIAN_URL="https://..."; $env:PULSE_URL="https://..."; npm start`.)

**Port 3443 in use?** Pick another, and build for it: `node tools/package.js
https://localhost:4443` then `node tools/serve.js 4443`.

## 6. Install the add-in in Excel

The build writes three files into `dist/` that you will use:

| File | Use |
|---|---|
| `dist/manifest.xml` | Installs the add-in into Excel |
| `dist/Timebase Historian Guide.xlsx` | The getting-started guide: examples, reference, troubleshooting |
| `dist/Timebase Historian.xlsx` | A blank workbook linked to the add-in |

### Excel for Mac

```bash
npm run sideload:mac
```

Then quit Excel completely (**Cmd-Q**) and open `dist/Timebase Historian Guide.xlsx`.
The **Timebase Historian** task pane opens on the right.

> **Why open a workbook rather than click a ribbon button?** Since Microsoft's
> 24-25 September 2026 Office update, add-ins installed this way that have a
> ribbon button fail to load (OfficeDev/office-js#6973). This build has no
> button and is started from a workbook linked to it instead. Any workbook you
> save with the task pane open stays linked. Excel lists the add-in under
> **Home > Add-ins**, but on Mac choosing it there does not open it -- open a
> linked workbook instead.

### Excel for Windows

1. **Home > Add-ins > More Add-ins > My Add-ins > Upload My Add-in** (in some
   versions **Insert > My Add-ins**), choose `dist\manifest.xml`.
2. Open `dist\Timebase Historian Guide.xlsx`.

For a team, publish `dist\manifest.xml` from a
[shared-folder catalog](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-add-in-deployment)
or the Microsoft 365 admin center. Windows has not been tested with this build
as thoroughly as Mac.

### Excel on the web

**Home > Add-ins > More Add-ins > Upload My Add-in**, choose `dist/manifest.xml`.
The server still has to run on the computer whose browser you are using.

## 7. Check it works

In the guide workbook:

1. The task pane's **Connect** tab says **Authentication is handled by the
   server** (client `Excel`). Click **Test connection** -- it should report your
   datasets.
2. On the **Setup** sheet, the `TB.DATASETS` table lists your datasets.
3. Type your own dataset and tag names into the shaded Setup cells. Every
   example sheet updates.

Then read the [User guide](user-guide.md), or use the task pane's **Insert**
tab to build formulas in your own workbooks.

## Upgrading

```bash
git pull
npm test
npm run build
npm run sideload:mac      # Mac; on Windows, upload dist\manifest.xml again
```

Restart `npm start` and Excel. Workbooks linked to the add-in keep working: the
link names the add-in by id, and the build stamps the guide and launcher with
the new version.

If formulas show `#BUSY!` or old behaviour after an upgrade, Excel is using a
cached copy of the add-in: quit Excel and clear its add-in cache (Mac:
`~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`
-- move its contents aside; Excel rebuilds it).

## Uninstalling

```bash
node tools/sideload-mac.js --remove     # Mac
node tools/set-credentials.js --clear   # removes the stored client secret
```

On Windows, remove it from **My Add-ins**. In Pulse, delete the **Excel** client
(**Identity > Clients > Excel > Edit > Delete**) or rotate its secret.

## If something goes wrong

| Symptom | Fix |
|---|---|
| `npm start` says *No development certificate found* | Step 3.1 |
| Test connection: *Pulse did not issue a token* | The client id or secret is wrong: step 2.3, then step 4, then restart `npm start` |
| Test connection or formulas: *401* | The client's **Audiences** does not include the Historian's client id (step 2.2) |
| *Could not reach Pulse* / *the Historian* | The server cannot reach them: check `HISTORIAN_URL` / `PULSE_URL` and that `.certs/pulse-ca.cer` is the right CA |
| The task pane never appears | `npm start` is not running, or the add-in is not installed (step 6). On Mac, open a linked workbook rather than using Home > Add-ins |
| *This add-in is no longer available* | Two installed copies, or a stale link: run `npm run sideload:mac` again, restart Excel |
| Bursts of red quality-64 values on a laptop | The computer slept and the Historian marked the collector disconnected. Expected; see [Concepts](concepts.md#quality-64-collector-disconnected) |

More in the [User guide's troubleshooting](user-guide.md#troubleshooting).
