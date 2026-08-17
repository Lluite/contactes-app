const CACHE_NAME = "contactes-app-v26";
const APP_FILES = [
  "./",
  "./index.html?v=26",
  "./styles.css?v=26",
  "./app.js?v=26",
  "./manifest.webmanifest?v=26",
  "./robots.txt",
  "./jszip.min.js",
  "./office-desk-premium.png",
  "./office-meeting-premium.png",
  "./office-lounge-premium.png",
];

self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_FILES)));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    Promise.all([
      caches.keys().then((keys) =>
        Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
      ),
      self.clients.claim(),
    ])
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone)).catch(() => {});
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
