/* 오프라인 열람 지원
   ------------------------------------------------------------------
   데이터(data/)는 '網 우선' — 새 자료를 올리면 곧바로 반영되고,
   연결이 없을 때만 캐시본을 쓴다.
   셸(html/css/js)은 '캐시 우선' — 빠르게 뜨되 뒤에서 갱신한다. */
const V = 'sinhangeul-v6';
const SHELL = ['./', './index.html', './assets/css/app.css',
               './assets/js/app.js', './assets/js/search.worker.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (u.origin !== location.origin || e.request.method !== 'GET') return;
  // 번역 DOCX 내려받기는 캐시에 담지 않고 그대로 보낸다
  if (u.pathname.includes('/sources/')) return;

  // 대조 데이터와 도구(apps/): 網 우선 (갱신하면 즉시 보인다)
  if (u.pathname.includes('/data/') || u.pathname.includes('/apps/')) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          if (res.ok) {
            const cl = res.clone();
            caches.open(V).then(c => c.put(e.request, cl));
          }
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // 셸: 캐시 우선 + 뒤에서 갱신
  e.respondWith(
    caches.match(e.request).then(hit => {
      const net = fetch(e.request).then(res => {
        if (res.ok) {
          const cl = res.clone();
          caches.open(V).then(c => c.put(e.request, cl));
        }
        return res;
      }).catch(() => hit);
      return hit || net;
    })
  );
});
