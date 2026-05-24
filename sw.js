// 杉ヶ平料金計算機 Service Worker
// バージョン更新時はこの番号を変えると、次回起動時に最新ファイルが取得される
var CACHE_NAME = "sugigadaira-calc-v4";

// コアアセット（必須・失敗したらSWインストール失敗）
var CORE_ASSETS = [
  "./",
  "./index.html",
  "./records.html",
  "./records-shared.js",
  "./tsukikei-template.xlsx",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

// 外部アセット（CDN・失敗してもSW自体は成立。fetchハンドラ側で再試行可能）
var EXTERNAL_ASSETS = [
  "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      var core = cache.addAll(CORE_ASSETS);
      var external = Promise.all(EXTERNAL_ASSETS.map(function (url) {
        return fetch(url, { mode: "cors" })
          .then(function (res) { if (res && res.ok) return cache.put(url, res); })
          .catch(function () { /* 取得失敗は許容（次回fetchで取得を試みる） */ });
      }));
      return Promise.all([core, external]);
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
