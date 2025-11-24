const CACHE_NAME = 'woodapp-v1';
const PRECACHE_URLS = [
  '/',
  '/index.html',
  '/style.css',
  '/script.js',
  '/manifest.json',
  '/stovemanufacturers.json',
  '/woodicon.png'
  // add icon filenames below if you place them at root
  // '/stove-icon-192.png',
  // '/stove-icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(PRECACHE_URLS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // for navigation requests, try network first then cache fallback
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).then(resp => {
        const copy = resp.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return resp;
      }).catch(() => caches.match('/index.html'))
    );
    return;
  }

  // for other requests, respond with cache-first then network
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request).then(resp => {
      // optionally cache the response
      return resp;
    })).catch(() => {
      // final fallback: try index.html for routes
      return caches.match('/index.html');
    })
  );
});
