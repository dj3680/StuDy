/* StuDy service worker — offline shell + installability.
   NOTE: true Web Push needs a push server (phase 2). This SW keeps the
   app installable and usable offline for already-loaded content. */
const CACHE = "study-v2-12";
const CORE = ["./", "./index.html", "./questions.js", "./manifest.webmanifest",
  ];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(()=>self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())); });
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c=>c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request)));
});
self.addEventListener("push", e => {
  const d = e.data ? e.data.json() : { title: "StuDy", body: "New update for you!" };
  e.waitUntil(self.registration.showNotification(d.title, { body: d.body, icon: "assets/icon-192.png" }));
});
