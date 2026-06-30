// FCCI Service Worker - enables offline basic shell and fast loading
const CACHE_NAME = 'fcci-cache-v2';  // ← binago mula v1 papuntang v2
const STATIC_ASSETS = [
  '/static/fcci_logo.jpeg',
  'https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.map(k => caches.delete(k)))  // burahin LAHAT ng lumang caches
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  // I-exclude ang CSS files sa caching — palaging kunin mula
  // sa network para hindi ma-stuck sa lumang version
  if (event.request.url.endsWith('.css')) {
    event.respondWith(fetch(event.request));
    return;
  }

  // Cache-first lang para sa images at fonts (hindi madalas magbago)
  if (event.request.url.includes('/static/') &&
      !event.request.url.endsWith('.css') &&
      !event.request.url.endsWith('.js')) {
    event.respondWith(
      caches.match(event.request).then(cached =>
        cached || fetch(event.request).then(response => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        })
      )
    );
  }
});
