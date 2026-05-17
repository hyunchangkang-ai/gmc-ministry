const CACHE = 'gmc-v6';
const STATIC_ASSETS = [
  '/gmc-ministry/icons/icon-192.png',
  '/gmc-ministry/icons/icon-512.png',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC_ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Apps Script / 외부 API는 항상 네트워크
  if (url.includes('script.google.com')) return;

  // HTML 파일은 항상 네트워크 우선, 실패 시에만 캐시
  if (url.endsWith('.html') || url.includes('/gmc-ministry/') && !url.match(/\.(png|jpg|ico|json|js|css)$/)) {
    e.respondWith(
      fetch(e.request, { cache: 'no-store' })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // 아이콘 등 정적 자산은 캐시 우선
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
