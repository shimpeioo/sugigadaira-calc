// 杉ヶ平料金計算機 Service Worker
// バージョン更新時はこの番号を変えると、次回起動時に最新ファイルが取得される
var CACHE_NAME = "sugigadaira-calc-v1";

var ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Stale-while-revalidate：キャッシュ即返し→裏でネットワーク取得して更新
self.addEventListener("fetch", function (event) {
  if (event.request.method !== "GET") return;
  event.respondWith(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.match(event.request).then(function (cachedResponse) {
        var networkFetch = fetch(event.request).then(function (networkResponse) {
          if (networkResponse && networkResponse.status === 200) {
            cache.put(event.request, networkResponse.clone());
          }
          return networkResponse;
        }).catch(function () {
          return cachedResponse;
        });
        return cachedResponse || networkFetch;
      });
    })
  );
});
