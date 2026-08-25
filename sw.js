// 여행이력 앱 서비스워커
// - 앱 화면(껍데기)만 캐시해서 오프라인에서도 앱이 "열리기는" 하도록 함
// - 구글시트 저장/조회 API 요청은 항상 네트워크로 그대로 통과시킴 (캐시하지 않음)
const CACHE_VERSION = 'v2'; // index.html의 APP_VERSION이 바뀔 때 함께 올려주세요
const CACHE_NAME = 'travellog-shell-' + CACHE_VERSION;

const SHELL_FILES = [
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // 같은 출처(우리 앱 파일)만 캐시 대상으로 처리. 구글 스프레드시트 API 등
  // 다른 출처(cross-origin) 요청은 서비스워커가 손대지 않고 그대로 네트워크로 보냄.
  if (url.origin !== self.location.origin) return;
  if (req.method !== 'GET') return;

  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req).then((res) => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        }
        return res;
      }).catch(() => cached);
      return cached || network;
    })
  );
});
