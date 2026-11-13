// 발자취 서비스 워커 — 앱 셸은 캐시, 구글 시트 API는 항상 네트워크
const CACHE = 'footprint-v17';
const SHELL = ['./', 'index.html', 'manifest.json', 'icons/icon-192.png', 'icons/apple-touch-icon.png',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.hostname.endsWith('google.com') || url.hostname.endsWith('googleusercontent.com') || url.hostname.endsWith('kakao.com') || url.hostname.endsWith('daumcdn.net')) return;
  // 네트워크 우선, 실패 시 캐시 (오프라인에서도 앱이 열리도록)
  e.respondWith(
    fetch(e.request).then(res => {
      if (res.ok && (url.origin === location.origin || SHELL.includes(e.request.url))) {
        const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy));
      }
      return res;
    }).catch(() => caches.match(e.request).then(r => r || caches.match('index.html')))
  );
});
