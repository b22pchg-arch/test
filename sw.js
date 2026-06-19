const CACHE_NAME = 'excel-quiz-v45-30-html-open-fix';
const ASSETS = ['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png','./xlsx.full.min.js','./default-bank.json'];
self.addEventListener('install', event => { self.skipWaiting(); event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))); });
self.addEventListener('activate', event => { event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', event => {
  if(event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if(url.pathname.endsWith('/default-bank.json')){
    event.respondWith(fetch(event.request).then(resp => { const copy = resp.clone(); caches.open(CACHE_NAME).then(cache => cache.put('./default-bank.json', copy)); return resp; }).catch(() => caches.match('./default-bank.json')));
    return;
  }
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request).then(resp => { const copy = resp.clone(); caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy)); return resp; }).catch(() => caches.match('./index.html'))));
});
