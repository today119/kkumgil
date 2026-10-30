/* ══════════════════════════════════════════════════════════════
   sw.js — 오프라인 저장
   ★ 산에서 신호가 약할 때 «열자마자 바로» 떠야 한다. 그게 이 앱의 존재 이유다.
     코스·사진을 폰에 받아 두고, 그 뒤엔 네트워크 없이도 연다.

   규칙
     · 내 파일(코스·사진·화면)  → 저장해 둔 것을 «먼저» 주고, 뒤에서 조용히 갱신
     · 카카오 지도 타일         → 저장하지 않는다(약관·용량). 신호 없으면 지도만 안 보인다
   ══════════════════════════════════════════════════════════════ */
const V = 'kkumgil-v142';
const CORE = ['./', './index.html', './courses.json', './course-day1.json', './course-day2.json',
              './manifest.json', './icon-192.png', './icon-512.png', './stamps.json',
              './hero.jpg', './t-poem.jpg', './t-catch.jpg', './t-shirt-dpw2.jpg',
              './t-photo.jpg', './t-reels.jpg', './t-report.jpg', './jsqr.min.js', './prep.json', './t-prep.jpg', './route-day1.jpg', './route-day2.jpg', './shirt-w1.jpg', './shirt-w2.jpg', './shirt-w3.jpg', './shirt-w4.jpg', './shirt-w5.jpg', './t-plan.jpg', './dash.html'];

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const c = await caches.open(V);
    await c.addAll(CORE).catch(() => {});
    // ★ 사진(6.7MB)은 여기서 받지 않는다(10/15). 설치 중에 한꺼번에 받으면 그동안 코스 목록 같은
    //   작은 자료가 줄을 서서 첫 화면 카드가 한참 늦게 떴다. 사진은 화면이 다 뜬 뒤 앱이 PRECACHE 로
    //   한 장씩 천천히 받게 한다(index.html 의 autoPrecache).
    self.skipWaiting();
  })());
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== V) await caches.delete(k);
    await self.clients.claim();
  })());
});

/* ★ 화면(html)과 자료(json)는 «인터넷 먼저», 안 되면 저장본 — 고친 게 바로 보이게.
     예전엔 전부 «저장본 먼저»여서 홈 화면 앱에서 옛 화면이 계속 떴다(10/3 사용자 화면에서 확인).
     사진·그림은 «저장본 먼저»(용량이 커서 산에서 다시 받으면 느리다). */
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (u.origin !== location.origin) return;          // 카카오 지도 등 남의 것은 건드리지 않는다
  const fresh = e.request.mode === 'navigate' || /\.(html|json|js)$/.test(u.pathname) || u.pathname.endsWith('/') || u.pathname.includes('/car/') || u.pathname.includes('/lead/');  // 진행·차량 카드는 자주 고친다
  e.respondWith((async () => {
    const c = await caches.open(V);
    const hit = await c.match(e.request, { ignoreSearch: true });
    if (fresh) {
      try {
        const r = await fetch(e.request, { cache: 'no-cache' });
        if (r && r.ok) { c.put(e.request, r.clone()); return r; }
      } catch (err) { /* 신호 없음 → 저장본 */ }
      return hit || new Response('', { status: 504 });
    }
    const net = fetch(e.request).then((r) => {
      if (r && r.ok) c.put(e.request, r.clone());
      return r;
    }).catch(() => null);
    return hit || (await net) || new Response('', { status: 504 });
  })());
});

/* 화면에서 「미리 받기」를 누르면 사진까지 한 번에 받아 둔다 */
self.addEventListener('message', (e) => {
  if (e.data && e.data.type === 'PRECACHE') {
    e.waitUntil((async () => {
      const c = await caches.open(V);
      const urls = e.data.urls || [];
      let done = 0;
      for (const u of urls) {                       // 하나씩, 이미 있는 건 건너뛴다
        if (!(await c.match(u))) await c.add(u).catch(() => {});
        done++;
      }
      (await self.clients.matchAll()).forEach((cl) => cl.postMessage({ type: 'PRECACHE_DONE', done, total: urls.length }));
    })());
  }
});
