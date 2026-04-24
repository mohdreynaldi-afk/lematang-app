const CACHE_NAME = "lematang-v1";

const urlsToCache = [
  "/",
  "/index.html"
];

// install
self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(urlsToCache))
  );
});

// fetch
self.addEventListener("fetch", event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => response || fetch(event.request))
  );
});

<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <script>
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('service-worker.js')
    .then(() => console.log('SW aktif'))
    .catch(err => console.log(err));
}
</script>
  </body>
</html>
