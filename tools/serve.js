// Local HTTPS server for sideloading this add-in during development, and a
// same-origin proxy to the Historian and Pulse.
//
// Why a proxy: neither the Historian nor Pulse answers a CORS preflight (both
// return 405 to OPTIONS), and every Historian call carries an Authorization
// header, so it always preflights. A task pane served from anywhere other than
// the Historian's own origin is therefore blocked by the browser engine --
// WebView2 on Windows, WKWebView on Mac, the browser for Excel on the web.
// Serving the add-in AND forwarding /api/* and /auth/* from one origin removes
// CORS from the picture entirely. In the task pane, set both the Historian and
// the Pulse address to this server's address (https://localhost:3443).
//
// Office add-ins are only loaded over HTTPS, and Excel refuses a certificate
// the OS does not trust. Run once:
//   npx --yes office-addin-dev-certs install
// which writes a localhost certificate to ~/.office-addin-dev-certs and trusts
// it (one system prompt).
//
//   node tools/serve.js [port]
//
// Environment:
//   HISTORIAN_URL  upstream Historian   (default https://localhost:4512)
//   PULSE_URL      upstream Pulse       (default https://localhost:4542)
//   UPSTREAM_CA    PEM file of the CA that signed the upstream certificates
//                  (default .certs/pulse-root-ca.pem, if present). On a Mac,
//                  export it from the keychain with:
//                    security find-certificate -p -c "Timebase Pulse Root CA" > .certs/pulse-root-ca.pem

import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import https from "node:https";
import tls from "node:tls";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import { bundle } from "./bundle.js";
import { CREDENTIALS_PATH, loadCredentials } from "./credentials.js";

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const PORT = Number(process.argv[2]) || 3443;

const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const CERT_FILE = path.join(CERT_DIR, "localhost.crt");
const KEY_FILE = path.join(CERT_DIR, "localhost.key");

const UPSTREAMS = [
  { prefix: "/api/", target: new URL(process.env.HISTORIAN_URL || "https://localhost:4512") },
  { prefix: "/auth/", target: new URL(process.env.PULSE_URL || "https://localhost:4542") },
];

const HISTORIAN = UPSTREAMS[0].target;
const PULSE = UPSTREAMS[1].target;

// --- server-held Pulse credentials ---------------------------------------------
//
// When tools/set-credentials.js has written a credentials file, this server
// signs in to Pulse itself and attaches the token to every Historian request.
// Excel then never needs, stores or sends the client secret: the task pane sees
// /__server-auth report "enabled" and turns its own authentication off.

let credentials = null;
try {
  credentials = loadCredentials();
} catch (err) {
  console.error(`Ignoring credentials: ${err.message}`);
}

let cachedToken = null; // { token, expiresAtMs }
let tokenInFlight = null;

function requestToken() {
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: credentials.clientId,
    client_secret: credentials.clientSecret,
  });
  if (credentials.scope) body.append("scope", credentials.scope);
  const payload = body.toString();

  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        hostname: PULSE.hostname,
        port: PULSE.port || 443,
        method: "POST",
        path: "/auth/token",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Content-Length": Buffer.byteLength(payload),
          Accept: "application/json",
        },
        ca: upstreamCa,
      },
      (res) => {
        let text = "";
        res.setEncoding("utf8");
        res.on("data", (c) => (text += c));
        res.on("end", () => {
          let json = null;
          try {
            json = JSON.parse(text);
          } catch {
            /* reported below */
          }
          if (res.statusCode !== 200 || !json || !json.access_token) {
            const said = json ? [json.error, json.error_description].filter(Boolean).join(": ") : text.slice(0, 200);
            reject(new Error(`Pulse refused client "${credentials.clientId}" (HTTP ${res.statusCode}): ${said}`));
            return;
          }
          const lifetimeMs = Number(json.expires_in ?? 3600) * 1000;
          // Renew a minute early: the Historian validates with no clock tolerance.
          resolve({ token: json.access_token, expiresAtMs: Date.now() + Math.max(0, lifetimeMs - 60000) });
        });
      },
    );
    req.on("error", (err) => reject(new Error(`Could not reach Pulse at ${PULSE.origin}: ${err.message}`)));
    req.end(payload);
  });
}

/** A valid token, cached until shortly before expiry; concurrent callers share one request. */
async function serverToken() {
  if (cachedToken && cachedToken.expiresAtMs > Date.now()) return cachedToken.token;
  if (!tokenInFlight) {
    tokenInFlight = requestToken()
      .then((t) => {
        cachedToken = t;
        return t.token;
      })
      .finally(() => {
        tokenInFlight = null;
      });
  }
  return tokenInFlight;
}

// The CA that signed the Historian's and Pulse's certificates. The Historian
// keeps a copy as certificates/pulse-ca.cer (DER); a PEM export works too.
const caCandidates = process.env.UPSTREAM_CA
  ? [process.env.UPSTREAM_CA]
  : [path.join(ROOT, ".certs", "pulse-root-ca.pem"), path.join(ROOT, ".certs", "pulse-ca.cer")];
const caFile = caCandidates.find((f) => fs.existsSync(f)) ?? caCandidates[0];

function readCaPem(file) {
  const raw = fs.readFileSync(file);
  if (raw.subarray(0, 64).toString("latin1").includes("-----BEGIN")) return raw.toString("utf8");
  return new crypto.X509Certificate(raw).toString(); // DER in, PEM out
}

let upstreamCa;
if (fs.existsSync(caFile)) {
  try {
    upstreamCa = [...tls.rootCertificates, readCaPem(caFile)];
  } catch (err) {
    console.error(`Ignoring ${caFile}: not a certificate (${err.message})`);
  }
}

if (!fs.existsSync(CERT_FILE) || !fs.existsSync(KEY_FILE)) {
  console.error(
    [
      "No development certificate found at:",
      `  ${CERT_FILE}`,
      `  ${KEY_FILE}`,
      "",
      "Office add-ins require HTTPS even on localhost, and Excel will not accept an",
      "untrusted self-signed certificate. Generate and install one with:",
      "",
      "  npx --yes office-addin-dev-certs install",
      "",
      "That command installs the certificate into your OS trust store (one prompt),",
      "then re-run this script.",
    ].join("\n"),
  );
  process.exit(1);
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8",
  ".map": "application/json; charset=utf-8",
};

// Hop-by-hop headers must not be forwarded by a proxy (RFC 9110 section 7.6.1).
const HOP_BY_HOP = new Set([
  "connection",
  "keep-alive",
  "proxy-connection",
  "transfer-encoding",
  "upgrade",
  "te",
  "trailer",
]);

function proxy(req, res, target, { bearer = null } = {}) {
  const headers = {};
  for (const [k, v] of Object.entries(req.headers)) {
    if (!HOP_BY_HOP.has(k.toLowerCase())) headers[k] = v;
  }
  headers.host = target.host;
  if (bearer) headers.authorization = `Bearer ${bearer}`;

  const upstream = https.request(
    {
      protocol: target.protocol,
      hostname: target.hostname,
      port: target.port || 443,
      method: req.method,
      path: req.url,
      headers,
      ca: upstreamCa,
    },
    (up) => {
      const out = {};
      for (const [k, v] of Object.entries(up.headers)) {
        if (!HOP_BY_HOP.has(k.toLowerCase())) out[k] = v;
      }
      // A rejected server-held token is dropped, so the next request signs in again.
      if (bearer && up.statusCode === 401) cachedToken = null;
      res.writeHead(up.statusCode || 502, out);
      up.pipe(res);
    },
  );
  upstream.on("error", (err) => {
    if (!res.headersSent) res.writeHead(502, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "upstream_unreachable", error_description: `${target.origin}: ${err.message}` }));
  });
  req.pipe(upstream);
}

function isLoopback(address) {
  return address === "127.0.0.1" || address === "::1" || address === "::ffff:127.0.0.1";
}

function serveStatic(req, res) {
  // Strip query string, decode percent-escapes, and resolve without ever
  // leaving ROOT -- a path like /../../etc/passwd must not escape the
  // project directory just because this is "only" a dev server.
  const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  const safePath = path.normalize(urlPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(ROOT, safePath === "/" ? "index.html" : safePath);
  if (!filePath.startsWith(ROOT + path.sep) || /(^|[/\\])\./.test(path.relative(ROOT, filePath))) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      res.end(`Not found: ${safePath}`);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      // Never cache during development: Excel's WebView otherwise holds a
      // stale functions.js or manifest resource across edits.
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

const server = https.createServer(
  {
    cert: fs.readFileSync(CERT_FILE),
    key: fs.readFileSync(KEY_FILE),
  },
  (req, res) => {
    // Loopback only. With server-held credentials this proxy hands out
    // authenticated Historian reads, so it must not be reachable from another
    // machine on the network.
    if (!isLoopback(req.socket.remoteAddress)) {
      res.writeHead(403);
      res.end("Forbidden: this development server only accepts connections from this machine.");
      return;
    }
    // Path only: query strings can carry tag names, and request bodies (which
    // carry the Pulse client secret) are never logged.
    res.on("finish", () => console.log(`${res.statusCode} ${req.method} ${(req.url || "").split("?")[0]}`));
    if (req.method === "POST" && (req.url || "") === "/__devlog") {
      let body = "";
      req.on("data", (c) => {
        if (body.length < 8000) body += c;
      });
      req.on("end", () => {
        console.log(`  addin ${body}`);
        res.writeHead(204);
        res.end();
      });
      return;
    }
    // Built on every request, so an edit to any src/lib module is live on the
    // next Excel reload without a separate build step.
    if ((req.url || "").split("?")[0] === "/src/functions/functions.bundle.js") {
      try {
        const code = bundle(path.join(ROOT, "src", "functions", "functions.js"));
        res.writeHead(200, { "Content-Type": "text/javascript; charset=utf-8", "Cache-Control": "no-store" });
        res.end(code);
      } catch (err) {
        console.error(`bundle failed: ${err.message}`);
        res.writeHead(500, { "Content-Type": "text/plain" });
        res.end(`bundle failed: ${err.message}`);
      }
      return;
    }
    if ((req.url || "").split("?")[0] === "/__server-auth") {
      res.writeHead(200, { "Content-Type": "application/json", "Cache-Control": "no-store" });
      res.end(JSON.stringify(credentials ? { enabled: true, clientId: credentials.clientId } : { enabled: false }));
      return;
    }

    const route = UPSTREAMS.find((u) => (req.url || "").startsWith(u.prefix));
    if (!route) {
      serveStatic(req, res);
      return;
    }

    if (credentials && route.target === HISTORIAN) {
      // Read-only, like the add-in itself: with this server's credentials
      // attached, a write would be a write the browser never authenticated.
      if (req.method !== "GET" && req.method !== "HEAD") {
        res.writeHead(405, { "Content-Type": "application/json", Allow: "GET, HEAD" });
        res.end(JSON.stringify({ error: "read_only", error_description: "This proxy only forwards Historian reads." }));
        return;
      }
      serverToken().then(
        (token) => proxy(req, res, route.target, { bearer: token }),
        (err) => {
          console.error(err.message);
          res.writeHead(502, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: "pulse_token_failed", error_description: err.message }));
        },
      );
      return;
    }

    proxy(req, res, route.target);
  },
);

// A client that rejects this server's certificate never gets as far as a
// request, so without this it would fail with nothing in the log.
server.on("tlsClientError", (err) => console.log(`TLS handshake failed: ${err.code || ""} ${err.message}`));

server.listen(PORT, () => {
  console.log(`Serving ${ROOT}`);
  console.log(`  https://localhost:${PORT}/src/taskpane/taskpane.html`);
  for (const u of UPSTREAMS) console.log(`  ${u.prefix}*  ->  ${u.target.origin}`);
  console.log(upstreamCa ? `  upstream CA: ${caFile}` : "  upstream CA: system default (no UPSTREAM_CA file found)");
  console.log(
    credentials
      ? `  Pulse auth:  server-held, client "${credentials.clientId}" from ${CREDENTIALS_PATH}`
      : `  Pulse auth:  none held here (run node tools/set-credentials.js to keep credentials out of Excel)`,
  );
  console.log("");
  console.log(`In the task pane, set the Historian AND Pulse address to https://localhost:${PORT}.`);
  console.log(`Build the manifest for this server with: node tools/package.js https://localhost:${PORT}`);
});
