// Garde l'appli disponible quand le réseau coupe. Les données (Supabase) ne passent jamais par ce cache.
const C="suivi-chantier-v4";
self.addEventListener("install",()=>self.skipWaiting());
self.addEventListener("activate",e=>e.waitUntil(self.clients.claim()));
self.addEventListener("fetch",e=>{const r=e.request;if(r.method!=="GET")return;const u=new URL(r.url);
  const app=u.origin===location.origin,lib=/(^|\.)(cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com)$/.test(u.hostname);
  if(!app&&!lib)return;
  e.respondWith(fetch(r).then(res=>{if(res.ok||res.type==="opaque"){const cp=res.clone();caches.open(C).then(c=>c.put(app?u.origin+u.pathname:r,cp))}return res})
    .catch(()=>caches.match(app?u.origin+u.pathname:r)))});
// Notifications sur le téléphone (espace DOMetVIE) : envoyées par la fonction Supabase « notif-push »
self.addEventListener("push",e=>{let d={};try{d=e.data.json()}catch(x){d={body:e.data?e.data.text():""}}
  e.waitUntil(self.registration.showNotification(d.title||"Suivi chantier AFER",{body:d.body||"",tag:d.tag||undefined,icon:"icone-192.png",data:{url:d.url||""}})
    .then(()=>self.registration.getNotifications()).then(L=>{if(navigator.setAppBadge)return navigator.setAppBadge(L.length)}).catch(()=>{}))});
self.addEventListener("notificationclick",e=>{e.notification.close();const u=(e.notification.data||{}).url||self.registration.scope;
  e.waitUntil(clients.matchAll({type:"window",includeUncontrolled:true}).then(L=>{const w=L.find(c=>c.url.startsWith(self.registration.scope));
    if(w&&w.navigate)return w.navigate(u).then(c=>(c||w).focus()).catch(()=>w.focus());return clients.openWindow(u)}))});
