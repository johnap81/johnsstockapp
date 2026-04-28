/* eslint-disable no-restricted-globals */
/** Bump when shipping UI/logic changes — old caches are deleted on activate. */
const CACHE = "jsa-v53";

/** These files must not be served stale from precache (otherwise tabs/charts vanish until cache clears). */
const NETWORK_FIRST_PATHS = new Set(["/app.js", "/styles.css"]);

const PRECACHE = [
  "/",
  "/index.html",
  "/app.js",
  "/styles.css",
  "/manifest.webmanifest",
  "/README.html",
  "/icons/icon.svg",
  "/icons/maskable.svg",
  "/icons/apple-touch-icon.svg",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(PRECACHE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    (async () => {
      for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
      await self.clients.claim();
    })()
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.startsWith("/api/")) return;

  const accept = req.headers.get("accept") || "";

  e.respondWith(
    (async () => {
      const path = url.pathname;

      // Always try network first for pages (HTML) so index.html updates are not stuck behind an old cache.
      if (req.mode === "navigate" || accept.includes("text/html")) {
        try {
          return await fetch(req);
        } catch {
          const hit = await caches.match(req);
          if (hit) return hit;
          throw new TypeError("offline");
        }
      }

      // App shell scripts/styles: network-first so chart tabs, portfolio logic, and overlays always match server.
      if (NETWORK_FIRST_PATHS.has(path)) {
        try {
          const res = await fetch(req, { cache: "no-store" });
          if (res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
          return res;
        } catch {
          const hit = await caches.match(req);
          if (hit) return hit;
          throw new TypeError("offline");
        }
      }

      const hit = await caches.match(req);
      if (hit) return hit;
      const res = await fetch(req);
      const copy = res.clone();
      if (res.ok) caches.open(CACHE).then((c) => c.put(req, copy));
      return res;
    })()
  );
});
