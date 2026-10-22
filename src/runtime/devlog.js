// Development-only console relay.
//
// Excel's add-in web view has no console you can see without extra setup, and
// on Mac a failing custom function shows nothing more than #BUSY! or #VALUE!.
// When this page is served from localhost (tools/serve.js), console output and
// uncaught errors are also POSTed to /__devlog, which the dev server prints.
// On any other host this file does nothing, so a deployed add-in never sends
// its console anywhere.
//
// A classic script, not a module, loaded before everything else so it catches
// errors thrown while the modules themselves are evaluated.

(function () {
  if (location.hostname !== "localhost" && location.hostname !== "127.0.0.1") return;

  function send(level, parts) {
    try {
      const text = parts
        .map((p) => {
          if (p instanceof Error) return `${p.name}: ${p.message}\n${p.stack || ""}`;
          if (typeof p === "string") return p;
          try {
            return JSON.stringify(p);
          } catch {
            return String(p);
          }
        })
        .join(" ");
      // Never relay anything that looks like a credential.
      const safe = text.replace(/(client_secret|access_token|authorization|bearer)\S*/gi, "$1=<redacted>");
      fetch("/__devlog", {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: `[${level}] ${safe}`.slice(0, 4000),
        keepalive: true,
      }).catch(() => {});
    } catch {
      /* the relay must never break the page */
    }
  }

  for (const level of ["log", "info", "warn", "error"]) {
    const original = console[level].bind(console);
    console[level] = (...args) => {
      original(...args);
      send(level, args);
    };
  }
  window.addEventListener("error", (e) => send("uncaught", [e.error || e.message, `${e.filename}:${e.lineno}`]));
  window.addEventListener("unhandledrejection", (e) => send("unhandled-rejection", [e.reason]));
  send("info", [`devlog active: ${location.href} ua=${navigator.userAgent}`]);
})();
