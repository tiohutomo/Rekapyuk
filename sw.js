var C="rekapyuk-v1";
var LIBS=["https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js","https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js","https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js","https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js"];
self.addEventListener("install",function(e){
  e.waitUntil(caches.open(C).then(function(c){
    return c.addAll(["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png"]).then(function(){
      return Promise.all(LIBS.map(function(u){return c.add(u).catch(function(){})}))
    })
  }).then(function(){return self.skipWaiting()}))
});
self.addEventListener("activate",function(e){
  e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==C}).map(function(k){return caches.delete(k)}))}).then(function(){return self.clients.claim()}))
});
self.addEventListener("fetch",function(e){
  var r=e.request;if(r.method!=="GET")return;
  var u=new URL(r.url),same=u.origin===location.origin,cdn=/(^|\.)(cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|tessdata\.projectnaptha\.com)$/.test(u.hostname);
  if(!same&&!cdn)return;
  function store(res){if(res&&(res.ok||res.type==="opaque")){var cp=res.clone();caches.open(C).then(function(c){c.put(r,cp)})}return res}
  if(same){e.respondWith(fetch(r).then(store).catch(function(){return caches.match(r).then(function(h){return h||caches.match("index.html")})}))}
  else{e.respondWith(caches.match(r).then(function(h){return h||fetch(r).then(store)}))}
});
