const CACHE='quiz-excel-offline-v10-20260613-compact-force-update';
const ASSETS=['./index.html','./app.js','./xlsx.full.min.js','./manifest.webmanifest','./icon.svg'];
self.addEventListener('install', event => { self.skipWaiting(); event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS))); });
self.addEventListener('activate', event => { event.waitUntil((async()=>{ const names=await caches.keys(); await Promise.all(names.filter(n=>n!==CACHE && n.includes('quiz-excel')).map(n=>caches.delete(n))); await self.clients.claim(); })()); });
self.addEventListener('message', event => { if(event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting(); });
self.addEventListener('fetch', event => {
  const req=event.request;
  if(req.method !== 'GET') return;
  event.respondWith((async()=>{
    if(req.mode === 'navigate'){
      try { const fresh = await fetch(req); const cache=await caches.open(CACHE); cache.put('./index.html', fresh.clone()); return fresh; }
      catch(e){ return (await caches.match('./index.html')) || Response.error(); }
    }
    const cached = await caches.match(req);
    if(cached) return cached;
    try { const fresh = await fetch(req); const cache=await caches.open(CACHE); cache.put(req, fresh.clone()); return fresh; }
    catch(e){ return Response.error(); }
  })());
});