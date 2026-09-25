// Bump this on every deploy that changes any cached file — it forces old
// caches to be dropped so users actually get the new version.
const CACHE_VERSION = "v11";
const CACHE_NAME = "stock-radial-" + CACHE_VERSION;

// Same-origin app shell + the two CDN libraries the app depends on.
// Precached at install time so the app opens fully offline afterwards.
const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-192-maskable.png",
  "./icons/icon-512.png",
  "./icons/icon-512-maskable.png",
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
  "https://cdn.jsdelivr.net/npm/animejs@3.2.2/lib/anime.min.js",
  "https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Rajdhani:wght@500;600;700&display=swap"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      await Promise.all(
        PRECACHE_URLS.map((url) =>
          fetch(url, { mode: url.startsWith("http") ? "no-cors" : "same-origin" })
            .then((res) => cache.put(url, res))
            .catch(() => {
              // Ignore individual failures (e.g. offline on first install, or a
              // blocked font request) so one bad resource doesn't sink the rest.
            })
        )
      );
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((names) => Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

// Stale-while-revalidate: serve the cached copy instantly (works offline),
// and refresh the cache in the background whenever the network is available.
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const url = new URL(event.request.url);
  if (url.origin === "https://adizjust.pythonanywhere.com" && url.pathname.startsWith("/api/")) return;
  if (url.origin === self.location.origin && url.pathname.endsWith("/stock.xlsx")) {
    event.respondWith(fetch(event.request, { cache: "no-store" }));
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const network = fetch(event.request)
        .then((res) => {
          if (res && (res.ok || res.type === "opaque")) {
            const copy = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return res;
        })
        .catch(() => cached); // offline and not cached yet -> nothing we can do

      return cached || network;
    })
  );
});
