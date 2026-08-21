/* Service Worker — Finca La Esperanza
   Cachea el "esqueleto" de la app para que abra sin internet.
   Los datos NO se guardan aquí (eso lo hace IndexedDB en la app);
   esto solo hace que la pantalla cargue estando offline. */

const CACHE = 'finca-esperanza-v5';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './logo.jpg',
  './logo-badge.jpg',
  './icon-192.png',
  './icon-512.png',
  './icon.svg'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  // No interceptar las llamadas al Apps Script (deben ir a la red)
  if (req.url.includes('script.google.com') || req.method !== 'GET') return;

  e.respondWith(
    caches.match(req).then(cached => {
      if (cached) return cached;
      return fetch(req).then(resp => {
        if (resp && resp.status === 200 && req.url.startsWith(self.location.origin)) {
          const copy = resp.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return resp;
      }).catch(() => cached);
    })
  );
});
