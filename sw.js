/* 健康記錄機器人：離線快取。主程式先預載；文字辨識的程式與字庫第一次用到時才快取。 */
var V="health-robot-v2.4";
var CORE=["./","index.html","manifest.webmanifest","lib/xlsx.full.min.js","icons/icon-192.png","icons/icon-512.png"];
self.addEventListener("install",function(e){e.waitUntil(caches.open(V).then(function(c){return c.addAll(CORE)}).then(function(){return self.skipWaiting()}))});
self.addEventListener("activate",function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==V}).map(function(k){return caches.delete(k)}))}).then(function(){return self.clients.claim()}))});
self.addEventListener("fetch",function(e){
  var u=new URL(e.request.url);if(e.request.method!=="GET"||u.origin!==location.origin)return;
  var page=e.request.mode==="navigate"||/\/(index\.html)?$/.test(u.pathname);
  if(page){ /* 主頁：有網路就拿最新版，沒網路用快取 */
    e.respondWith(fetch(e.request).then(function(r){var c=r.clone();caches.open(V).then(function(x){x.put(e.request,c)});return r}).catch(function(){return caches.match(e.request).then(function(r){return r||caches.match("index.html")})}));return}
  e.respondWith(caches.match(e.request).then(function(r){return r||fetch(e.request).then(function(n){if(n.ok){var c=n.clone();caches.open(V).then(function(x){x.put(e.request,c)})}return n})}));
});
