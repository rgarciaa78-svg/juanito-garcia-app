const VERSION = 'v14';
self.addEventListener('install', e => { self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(keys => Promise.all(keys.map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.url.includes('index.html') || e.request.url.endsWith('/') || e.request.url.includes('juanito-garcia-app/?') || !e.request.url.includes('.')) {
    e.respondWith(fetch(e.request, { cache: 'no-store' }).catch(() => caches.match(e.request)));
  }
});
