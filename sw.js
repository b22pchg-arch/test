const CACHE_NAME = 'trac-nghiem-excel-offline-v22-guide-export-study';
const ASSETS = ['./','./index.html','./manifest.webmanifest','./icon.svg','./icon-192.png','./icon-512.png'];
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS.map(u => new Request(u, {cache:'reload'})))).catch(()=>{}));
});
self.addEventListener('activate', event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('message', event => { if(event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting(); });
self.addEventListener('fetch', event => {
  const req = event.request;
  if(req.method !== 'GET') return;
  event.respondWith(caches.match(req).then(cached => cached || fetch(req).then(resp => {
    const copy = resp.clone();
    caches.open(CACHE_NAME).then(cache => cache.put(req, copy)).catch(()=>{});
    return resp;
  }).catch(() => caches.match('./index.html'))));
});
