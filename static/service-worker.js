const CACHE_NAME = 'akuzfiro-v1';
const ASSETS = [
  '/',
  '/manifest.json'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // Solo cachea assets estáticos, las llamadas al API van siempre a la red
  if (e.request.url.includes('/chat') ||
      e.request.url.includes('/tts') ||
      e.request.url.includes('/memoria') ||
      e.request.url.includes('/generar-')) {
    e.respondWith(fetch(e.request));
    return;
  }
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
