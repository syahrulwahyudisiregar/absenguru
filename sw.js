// Service worker Ver. 005
// - File aplikasi: network-first (update index.html langsung terbaca), cadangan dari cache.
// - Script & model face-api (cdn.jsdelivr.net): cache-first, tersimpan setelah tablet dibuka sekali dengan internet,
//   sehingga tablet tetap bisa menyala saat internet mati.
// - Request ke Google Apps Script TIDAK di-cache (selalu online).
const CACHE = 'absensi-v5';
const CDN_CACHE = 'absensi-cdn-v1';
const FILES = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];
const FACE_JS = 'https://cdn.jsdelivr.net/npm/@vladmandic/face-api/dist/face-api.js';

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(FILES))
      .then(() => caches.open(CDN_CACHE))
      .then(c => c.add(new Request(FACE_JS, { mode: 'cors' })).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(
    k.filter(x => x !== CACHE && x !== CDN_CACHE).map(x => caches.delete(x))
  )));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const u = new URL(e.request.url);

  if (u.hostname === 'cdn.jsdelivr.net') {
    e.respondWith(caches.open(CDN_CACHE).then(c =>
      c.match(e.request).then(hit => hit || fetch(e.request).then(r => {
        if (r && (r.ok || r.type === 'opaque')) c.put(e.request, r.clone());
        return r;
      }))
    ));
    return;
  }

  if (u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => {
    const c = r.clone(); caches.open(CACHE).then(ch => ch.put(e.request, c)); return r;
  }).catch(() => caches.match(e.request)));
});
