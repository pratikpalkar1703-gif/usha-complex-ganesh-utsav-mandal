// Ganpati App — Service Worker v3
// ONLY caches the app page for offline loading.
// ALL other requests (API, fonts, CDN) pass through to network untouched.

const CACHE_NAME = 'ganpati-v3';

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      cache.addAll(['/ganpati/', '/ganpati/index.html'])
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  // Delete ALL old caches (removes any bad cached API responses from v1)
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  // ONLY intercept page navigation requests — serve from cache if offline
  // Everything else (API calls, CDN, fonts) goes straight to network
  if (event.request.mode !== 'navigate') return;

  event.respondWith(
    fetch(event.request)
      .then(res => {
        // Update cache with fresh copy on every successful load
        const clone = res.clone();
        caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
        return res;
      })
      .catch(() => caches.match('/ganpati/index.html'))
  );
});
