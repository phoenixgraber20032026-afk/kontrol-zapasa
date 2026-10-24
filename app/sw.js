/* Контроль неснижаемого запаса — service worker (режим «сайт»): страница и Chart.js хранятся в кэше,
   чтобы приложение открывалось мгновенно и без связи (данные при этом берутся из Google, когда связь есть). */
var CACHE = 'kz-app-v2.0';
var FILES = ['./', './index.html', './chart.umd.min.js', './manifest.json', './icon-192.png', './icon-512.png'];
self.addEventListener('install', function (e) { e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(FILES); }).then(function () { return self.skipWaiting(); })); });
self.addEventListener('activate', function (e) { e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); })); });
self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;           // данные Google (POST /exec) — всегда в сеть
  if (/^\/(api|edge|exec|ping)(\/|$)/.test(url.pathname)) return;                        // ускоритель: данные не кэшируются
  e.respondWith(fetch(e.request).then(function (r) { var cp = r.clone(); caches.open(CACHE).then(function (c) { c.put(e.request, cp); }); return r; })
    .catch(function () { return caches.match(e.request).then(function (m) { return m || caches.match('./index.html'); }); }));
});
