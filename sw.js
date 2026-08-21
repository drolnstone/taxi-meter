/*  Taxi Meter service worker.

    Purpose is narrow and deliberate: make the app open without a signal. The
    meter itself has always worked offline once loaded, but the page and Leaflet
    were fetched from the network on every open, so a driver's very first visit
    of the day — or after the browser evicted the page — needed data.

    Strategy:
      the app shell   network first, fall back to cache   (a redeploy is picked
                                                           up as soon as there is
                                                           a signal, and the old
                                                           copy keeps working
                                                           when there is not)
      Leaflet         cache first                          (pinned to 1.9.4, so
                                                           it never changes)
      map tiles       never cached                         (thousands of them, and
                                                           OSM's tile policy asks
                                                           you not to)
      the backend     never cached                         (fares and driver
                                                           records must not be
                                                           served stale, ever)

    Bump CACHE_VERSION on every deploy of index.html.  */

const CACHE_VERSION = "taxi-meter-v4";

const SHELL = [
  "./",
  "./index.html",
  "./manifest.json"
];

const VENDOR = [
  "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css",
  "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"
];

self.addEventListener("install", event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_VERSION);
    /* Individually, so one unreachable vendor URL cannot fail the whole install
       and leave the driver with no offline copy at all. */
    await Promise.all(SHELL.concat(VENDOR).map(async url => {
      try { await cache.add(new Request(url, { cache: "reload" })); }
      catch (e) { console.warn("[sw] could not precache", url, e); }
    }));
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

function isTile(url) {
  return /tile\.openstreetmap\.org/.test(url);
}
function isBackend(url) {
  return /script\.google\.com|api\.postcodes\.io|photon\.komoot\.io/.test(url);
}

self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = req.url;

  /* Fares, driver records and address lookups are never served from cache. A
     stale driver record could let a suspended ID start a job. */
  if (isBackend(url) || isTile(url)) return;

  if (VENDOR.indexOf(url) !== -1) {
    event.respondWith((async () => {
      const hit = await caches.match(req);
      if (hit) return hit;
      const res = await fetch(req);
      if (res && res.ok) (await caches.open(CACHE_VERSION)).put(req, res.clone());
      return res;
    })());
    return;
  }

  if (req.mode === "navigate" || /\.(html|json)$/.test(new URL(url).pathname) ||
      new URL(url).pathname.endsWith("/")) {
    event.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res && res.ok) (await caches.open(CACHE_VERSION)).put(req, res.clone());
        return res;
      } catch (e) {
        const hit = await caches.match(req) || await caches.match("./index.html");
        if (hit) return hit;
        return new Response(
          "<!doctype html><meta charset=utf-8><title>Offline</title>" +
          "<body style='font:16px system-ui;background:#0b0b0b;color:#fff;padding:32px'>" +
          "<h1>No connection</h1><p>The meter could not be loaded from the network and " +
          "there is no saved copy on this phone yet. Open the app once with a signal.</p>",
          { headers: { "Content-Type": "text/html; charset=utf-8" }, status: 503 });
      }
    })());
  }
});
