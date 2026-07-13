// Ganpati App — Service Worker v2
// Caches the app shell so it loads offline instantly

const CACHE_NAME = 'ganpati-v2';
const APP_SHELL  = [
  '/ganpati/',
  '/ganpati/index.html'
];

// Hostnames that must NEVER be cached — always go straight to network
const PASSTHROUGH_HOSTS = [
  'script.google.com',
  'script.googleusercontent.com',
  'googleapis.com',
  'accounts.google.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com'
];

function isPassthrough(url) {
  return PASSTHROUGH_HOSTS.some(h => url.hostname === h || url.hostname.endsWith('.' + h));
}

// Install: cache the app shell
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

// Activate: clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch handler
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // NEVER cache or intercept Google API calls — let them go straight to network
  // This is critical: GAS redirects to googleusercontent.com which must not be cached
  if (isPassthrough(url)) return;

  // For navigation (loading the app page) — network first, fall back to cache
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then(res => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
          return res;
        })
        .catch(() => caches.match('/ganpati/index.html'))
    );
    return;
  }

  // For CDN assets (html2canvas, jsPDF etc) — network first, cache fallback
  // Only cache responses with ok status (never cache opaque responses)
  event.respondWith(
    fetch(event.request)
      .then(res => {
        if (res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
        }
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
