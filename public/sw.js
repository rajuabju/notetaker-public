// App-shell cache so the installed app opens instantly. API calls are never cached.
const CACHE = "deal-notes-v8";
const SHELL = ["/", "/styles.css", "/app.js", "/manifest.webmanifest", "/icons/icon.svg"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin || url.pathname.startsWith("/api/")) return;
  // Network first (so updates and Access sign-in work), cached copy when offline.
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        if (res.ok && res.type === "basic") {
          const copy = res.clone();
          // Every page (/, /review/…, /act/…) is the same index.html: keep one copy under "/".
          caches.open(CACHE).then((c) => c.put(e.request.mode === "navigate" ? "/" : e.request, copy));
        }
        return res;
      })
      .catch(() => caches.match(e.request.mode === "navigate" ? "/" : e.request).then((r) => r || caches.match("/"))),
  );
});
