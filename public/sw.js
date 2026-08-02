// Pokémon GO Shop — offline service worker
// Strategy: network-first for navigation & same-origin requests, falling back
// to cache when offline. Everything successfully fetched gets cached, so the
// app (and its data, which already lives in localStorage) keeps working with
// no network at all after the first successful visit.

const CACHE_NAME = "pgs-cache-v1";
// Add any extra static assets you want guaranteed to be pre-cached on install.
// Hashed build files (e.g. Vite's /assets/*.js) don't need to be listed here —
// they get cached automatically the first time they're fetched.
const PRECACHE_URLS = ["./", "./index.html", "./manifest.json"];

self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      Promise.all(
        PRECACHE_URLS.map((url) => cache.add(url).catch(() => {
          // ignore individual precache failures (e.g. file doesn't exist yet)
        }))
      )
    )
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return; // don't touch POST/etc.
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return; // don't cache third-party/CDN requests

  event.respondWith(
    caches.open(CACHE_NAME).then(async (cache) => {
      try {
        const fresh = await fetch(request);
        if (fresh && fresh.status === 200) {
          cache.put(request, fresh.clone());
        }
        return fresh;
      } catch (err) {
        const cached = await cache.match(request, { ignoreSearch: true });
        if (cached) return cached;
        if (request.mode === "navigate") {
          const shell = await cache.match("./index.html");
          if (shell) return shell;
        }
        throw err;
      }
    })
  );
});
