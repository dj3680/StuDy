/* StuDy service worker — offline shell + installability + OneSignal push.
   The OneSignal worker code is merged into this worker (importScripts at
   the top) so ONE service worker serves both offline caching and push. */
try{ importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js"); }catch(e){}
const CACHE = "study-v2-90";
const CORE = ["./", "./index.html", "./questions.js", "./questions2.js", "./questions3.js", "./manifest.webmanifest",
  ];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(()=>self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())); });
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c=>c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request)));
});
/* Push display is handled by the OneSignal worker code imported above. */
