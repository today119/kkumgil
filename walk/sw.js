/* ══════════════════════════════════════════════════════════════
   sw.js — 오프라인 저장
   ★ 산에서 신호가 약할 때 «열자마자 바로» 떠야 한다. 그게 이 앱의 존재 이유다.
     코스·사진을 폰에 받아 두고, 그 뒤엔 네트워크 없이도 연다.

   규칙
     · 내 파일(코스·사진·화면)  → 저장해 둔 것을 «먼저» 주고, 뒤에서 조용히 갱신
     · 카카오 지도 타일         → 저장하지 않는다(약관·용량). 신호 없으면 지도만 안 보인다
   ══════════════════════════════════════════════════════════════ */
const V = 'kkumgil-v35';
const CORE = ['./', './index.html', './courses.json', './course-day1.json', './course-day2.json',
              './manifest.json', './icon-192.png', './icon-512.png', './stamps.json',
              './hero.jpg', './t-poem.jpg', './t-catch.jpg', './t-shirt-dpw2.jpg',
              './t-photo.jpg', './t-reels.jpg', './t-report.jpg', './jsqr.min.js', './prep.json', './t-prep.jpg', './route-day1.jpg', './route-day2.jpg', './shirt-w1.jpg', './shirt-w2.jpg', './shirt-w3.jpg', './shirt-w4.jpg', './shirt-w5.jpg', './t-plan.jpg'];

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const c = await caches.open(V);
    await c.addAll(CORE).catch(() => {});
    // 사진은 용량이 커서 실패해도 설치를 막지 않는다(한 장씩 따로 담는다)
    try {
      const list = await (await fetch('./course-day1.json', { cache: 'no-cache' })).json();
      const list2 = await (await fetch('./course-day2.json', { cache: 'no-cache' })).json();
      const photos = [];
      [list, list2].forEach((d) => (d.points || []).forEach((p) => (p.photos || []).forEach((f) => photos.push('./photos/' + f))));
      await Promise.all([...new Set(photos)].map((u) => c.add(u).catch(() => {})));
    } catch (err) { /* 사진은 나중에 받아도 된다 */ }
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
  const fresh = e.request.mode === 'navigate' || /\.(html|json|js)$/.test(u.pathname) || u.pathname.endsWith('/');
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
      for (const u of urls) { await c.add(u).catch(() => {}); done++; }
      (await self.clients.matchAll()).forEach((cl) => cl.postMessage({ type: 'PRECACHE_DONE', done, total: urls.length }));
    })());
  }
});
