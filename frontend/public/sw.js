const CACHE_NAME = "fichas-cnm-v3";
const BASE_PATH = new URL(self.registration.scope).pathname.replace(/\/$/, "");
const APP_SHELL = `${BASE_PATH}/`;

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      await cache.add(APP_SHELL);
      const shell = await cache.match(APP_SHELL);
      const html = await shell.text();
      const assets = [...html.matchAll(/(?:src|href)="([^"]+)"/g)]
        .map((match) => new URL(match[1], self.location.origin).pathname)
        .filter((path) => path.startsWith(`${BASE_PATH}/_next/static/`) || path.startsWith(`${BASE_PATH}/icon-`));
      await cache.addAll([...new Set(assets)]);
    }).then(() => self.skipWaiting()),
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || url.origin !== self.location.origin) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(APP_SHELL, copy));
        return response;
      }).catch(() => caches.match(APP_SHELL)),
    );
    return;
  }

  if (url.pathname.startsWith(`${BASE_PATH}/_next/static/`) || url.pathname.startsWith(`${BASE_PATH}/icon-`) || url.pathname.endsWith("/icon.svg")) {
    event.respondWith(
      caches.match(request).then((cached) => cached || fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      })),
    );
  }
});
