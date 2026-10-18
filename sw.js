// Garde l'appli disponible quand le réseau coupe. Les données (Supabase) ne passent jamais par ce cache.
const C="suivi-chantier-v2";
self.addEventListener("install",()=>self.skipWaiting());
self.addEventListener("activate",e=>e.waitUntil(self.clients.claim()));
self.addEventListener("fetch",e=>{const r=e.request;if(r.method!=="GET")return;const u=new URL(r.url);
  const app=u.origin===location.origin,lib=/(^|\.)(cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com)$/.test(u.hostname);
  if(!app&&!lib)return;
  e.respondWith(fetch(r).then(res=>{if(res.ok||res.type==="opaque"){const cp=res.clone();caches.open(C).then(c=>c.put(app?u.origin+u.pathname:r,cp))}return res})
    .catch(()=>caches.match(app?u.origin+u.pathname:r)))});
