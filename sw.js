// Ganpati App — Service Worker v5
// Caches the app shell for offline use with correct GitHub Pages path.

const CACHE_NAME = 'ganpati-v5';
const APP_PATH   = '/usha-complex-ganesh-utsav-mandal/';

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      cache.addAll([APP_PATH, APP_PATH + 'index.html'])
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  // Only intercept page navigation — API/CDN/font requests go straight to network
  if (event.request.mode !== 'navigate') return;

  event.respondWith(
    fetch(event.request)
      .then(res => {
        // Cache a fresh copy on every successful online load
        const clone = res.clone();
        caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
        return res;
      })
      .catch(() =>
        // Offline: serve the cached page
        caches.match(event.request)
          .then(cached => cached || caches.match(APP_PATH))
      )
  );
});
