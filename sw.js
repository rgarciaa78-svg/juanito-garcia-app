const VERSION = 'v19-nocache';
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.map(k => caches.delete(k))))
    .then(() => self.clients.claim())
  );
});

// OJO: este handler interceptaba TODAS las peticiones, incluidos los POST del
// chat al Worker de Cloudflare. Reconstruir un POST con `fetch(e.request, ...)`
// falla —el cuerpo ya se consumió—, el catch devolvía "Sin conexión" en 78 ms,
// y la app lo leía como "el modelo de IA está ocupado". La IA nunca llegó a
// salir del navegador: con curl funcionaba porque curl no pasa por aquí.
//
// Solo se tocan los GET del mismo origen, que es lo único que este worker
// necesita controlar. Todo lo demás va directo a la red.
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // POST, PUT: sin tocar
  if (new URL(req.url).origin !== self.location.origin) return;  // otros dominios: sin tocar
  e.respondWith(
    fetch(req, { cache: 'no-store' })
      .catch(() => new Response('Sin conexión', { status: 503 }))
  );
});
