/* Muir Receipts service worker — caches the app shell so the icon opens instantly and offline.
   Network calls to Microsoft (login, Graph) are never cached. Bump CACHE on each release. */
const CACHE = "muir-receipts-v0.7.0";
const SHELL = ["./", "./index.html", "./app.js?v=0.7.0", "./scan.js?v=0.7.0", "./config.js", "./msal-browser.min.js", "./manifest.webmanifest",
  "./icons/icon-192.png", "./icons/icon-512.png", "./icons/apple-touch-icon.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin || e.request.method !== "GET") return; // Graph / login: straight to network
  if (url.searchParams.has("nosw")) return;                                   // the app's own "is there a newer version?" check
  // Images (the team photo, icons): cache-first — downloaded once per version, not on every open.
  if (e.request.destination === "image" || /\.(jpg|jpeg|png|webp)$/i.test(url.pathname)) {
    e.respondWith(caches.match(e.request).then((m) => m || fetch(e.request).then((r) => {
      if (r.ok) { const copy = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
      return r;
    })));
    return;
  }
  // Versioned scripts (app.js?v=…, scan.js?v=…) never change: cache-first.
  if (/\.js$/.test(url.pathname) && url.searchParams.has("v")) {
    e.respondWith(caches.match(e.request).then((m) => m || fetch(e.request).then((r) => { if (r.ok) { const c2 = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, c2)); } return r; })));
    return;
  }
  // Everything else in the app (the page, config.js, msal): open instantly from the phone's copy, and fetch a fresh
  // copy in the background for next time (the app checks for a newer version itself and offers "Tap to update").
  const isPage = e.request.mode === "navigate" || url.pathname.endsWith("/") || url.pathname.endsWith("/index.html");
  const key = isPage ? "./index.html" : e.request;
  e.respondWith(caches.open(CACHE).then(async (c) => {
    const cached = await c.match(key, { ignoreSearch: isPage });
    const fresh = fetch(isPage ? "./index.html" : e.request, { cache: "no-cache" }).then(async (r) => {
      if (r.ok) {
        await c.put(key, r.clone());
      }
      return r;
    }).catch(() => null);
    if (cached) { e.waitUntil(fresh); return cached; }
    return (await fresh) || (await c.match("./index.html")) || Response.error();
  }));
});

// Tapping a reminder opens (or focuses) the app
self.addEventListener("notificationclick", (e) => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((cs) => {
    for (const c of cs) { if ("focus" in c) return c.focus(); }
    return self.clients.openWindow("./");
  }));
});
