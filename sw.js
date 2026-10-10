/*
  Archivo: sw.js
  OpenFondos · App envoltorio (PWA) · service worker · v1.0
  Guarda en caché solo la carcasa de la app (esta página, manifiesto e iconos).
  El panel y los datos se cargan siempre en directo desde Apps Script; nunca se guardan aquí.
*/
const CACHE = 'openfondos-app-v1';
const CARCASA = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];

self.addEventListener('install', ev => {
  ev.waitUntil(caches.open(CACHE).then(c => c.addAll(CARCASA)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', ev => {
  ev.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Red primero para la carcasa (así las actualizaciones se ven al momento); caché si no hay conexión.
self.addEventListener('fetch', ev => {
  const url = new URL(ev.request.url);
  if (ev.request.method !== 'GET' || url.origin !== self.location.origin) return;
  ev.respondWith(
    fetch(ev.request)
      .then(r => { const copia = r.clone(); caches.open(CACHE).then(c => c.put(ev.request, copia)); return r; })
      .catch(() => caches.match(ev.request).then(r => r || caches.match('index.html')))
  );
});
