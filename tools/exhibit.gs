/* ══════════════════════════════════════════════════════════════
   꿈길 걷기 — 사진전·릴스전 출품 시트 (Google Apps Script)

   ★ 흐름: 학생이 앱 「📤 출품하기」 → 구글 폼(파일 업로드) → 이 응답 시트
           → 선생님이 「공개」 체크한 것만 앱 갤러리에 뜬다(친구 얼굴이 동의 없이 퍼지지 않게)
           → 「우수작」 체크 → 메뉴 한 번으로 학생회 전달 폴더 → 학생회 인스타 좋아요 투표
   ★ 사진은 공유 설정을 건드리지 않는다. 서버(이 스크립트)가 썸네일을 받아 그림 글자로 내려 준다.
     영상 재생만 「학교 계정 + 링크」 보기 권한을 연다(공개 체크한 영상에 한해, 누를 때).

   쓰는 법 (학교 계정, 한 번만)
     1. 사진전 폼·릴스전 폼을 «같은 스프레드시트»에 응답 연결 → 탭 이름을 「사진전」「릴스전」으로 바꾸기
     2. 그 시트에서 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 저장
     3. 시트 새로고침 → 메뉴 「📸 꿈길 출품」 → ① 체크 칸 만들기
     4. Apps Script 화면 → 배포 → 새 배포 → 유형 «웹 앱» → 실행: 나 / 액세스: 영종고 사용자 → 배포
        → 나온 «웹 앱 URL» 을 앱(stamps.json 의 gallery)에 넣으면 앱에 「🖼 작품 보기」 버튼이 켜진다
   ══════════════════════════════════════════════════════════════ */

var E_APP = 'https://today119.github.io/kkumgil/walk/?tab=mis';
var E_COLS = ['공개', '우수작'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📸 꿈길 출품')
    .addItem('① 체크 칸 만들기(공개·우수작)', 'eSetup')
    .addItem('② 학생회 전달 폴더 만들기(우수작)', 'eCouncil')
    .addSeparator()
    .addItem('❓ 사용법', 'eHelp')
    .addToUi();
}

function eHelp() {
  SpreadsheetApp.getUi().alert('📸 꿈길 출품',
    '「공개」 체크 → 앱 갤러리(학교 계정만 볼 수 있음)에 떠요. 체크를 풀면 바로 사라져요.\n' +
    '「우수작」 체크 → ② 학생회 전달 폴더 만들기를 누르면 원본 사본이 폴더 하나에 모여요. 그 폴더를 학생회에 공유하세요.\n\n' +
    '갤러리 주소: Apps Script → 배포 → 배포 관리 → 웹 앱 URL\n' +
    '탭 이름에 「사진」「릴스」가 들어가야 종류를 알아봐요.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/* 탭 종류 — 이름에 «릴스»/«사진», 아니면 연결된 폼 제목으로 */
function eKind(sh) {
  var n = sh.getName();
  if (/릴스/.test(n)) return 'reels';
  if (/사진/.test(n)) return 'photo';
  try {
    var u = sh.getFormUrl();
    if (u) { var t = FormApp.openByUrl(u).getTitle(); if (/릴스/.test(t)) return 'reels'; if (/사진/.test(t)) return 'photo'; }
  } catch (e) {}
  return '';
}
function eTabs() {
  return SpreadsheetApp.getActive().getSheets().map(function (sh) { return { sh: sh, kind: eKind(sh) }; })
    .filter(function (x) { return x.kind; });
}

function eSetup() {
  var tabs = eTabs(), msg = [];
  if (!tabs.length) { SpreadsheetApp.getUi().alert('「사진전」「릴스전」 탭을 못 찾았어요. 응답 탭 이름을 바꿔 주세요.'); return; }
  tabs.forEach(function (t) {
    var sh = t.sh, last = Math.max(sh.getLastColumn(), 1);
    var H = sh.getRange(1, 1, 1, last).getDisplayValues()[0];
    E_COLS.forEach(function (c) {
      var i = H.indexOf(c);
      if (i < 0) { last++; sh.getRange(1, last).setValue(c).setFontWeight('bold').setBackground('#FEF3C7'); i = last - 1; H.push(c); }
      sh.getRange(2, i + 1, 800, 1).insertCheckboxes();
    });
    msg.push(sh.getName() + (t.kind === 'reels' ? ' (릴스전)' : ' (사진전)'));
  });
  SpreadsheetApp.getUi().alert('체크 칸을 만들었어요: ' + msg.join(', ') + '\n새 응답이 들어와도 800줄까지 체크 칸이 있어요.');
}

/* ── 응답 읽기 ── */
function eMask(nm) {
  nm = String(nm || '').replace(/\d+\s*번/g, '').replace(/\s+/g, ' ').trim();
  return nm.replace(/[가-힣]{2,4}$/, function (s) {
    return s.length === 2 ? s[0] + '○' : s[0] + new Array(s.length - 1).join('○') + s[s.length - 1];
  });
}
function eRows(kind) {
  var out = [];
  eTabs().forEach(function (t) {
    if (kind && t.kind !== kind) return;
    var v = t.sh.getDataRange().getValues(); if (v.length < 2) return;
    var H = v[0].map(function (x) { return String(x).replace(/\s+/g, ''); });
    var col = function (re) { for (var i = 0; i < H.length; i++) if (re.test(H[i])) return i; return -1; };
    var cOpen = H.indexOf('공개'), cBest = H.indexOf('우수작');
    var cMail = col(/이메일/), cName = col(/이름|성명/), cMis = col(/미션|장소|어디서/), cTitle = col(/제목/), cDesc = col(/설명|소개/), cTeam = col(/함께|모둠/);
    var cFile = -1;
    for (var i = 0; i < H.length && cFile < 0; i++)
      for (var r = 1; r < v.length; r++) if (/drive\.google\.com\/open\?id=/.test(String(v[r][i]))) { cFile = i; break; }
    if (cFile < 0) return;
    v.slice(1).forEach(function (r, k) {
      var ids = (String(r[cFile]).match(/id=[\w-]+/g) || []).map(function (s) { return s.slice(3); });
      if (!ids.length) return;
      var mail = cMail >= 0 ? String(r[cMail]) : '', m = mail.match(/s\d\d(\d)(\d\d)\d\d@/), who = '';
      if (m) who = m[1] + '학년 ' + Number(m[2]) + '반 ';
      who += eMask(cName >= 0 ? r[cName] : '');
      out.push({
        kind: t.kind, row: k + 2, ids: ids, who: who.trim(),
        open: cOpen >= 0 && r[cOpen] === true, best: cBest >= 0 && r[cBest] === true,
        mission: cMis >= 0 ? String(r[cMis]) : '', title: cTitle >= 0 ? String(r[cTitle]) : '',
        desc: cDesc >= 0 ? String(r[cDesc]) : '', team: cTeam >= 0 ? String(r[cTeam]) : ''
      });
    });
  });
  return out;
}
/* 공개 파일 목록 — 썸네일마다 시트를 다시 읽지 않게 1분 캐시 */
function eOpenIds() {
  var c = CacheService.getScriptCache(), hit = c.get('openIds');
  if (hit) return JSON.parse(hit);
  var o = {};
  eRows('').forEach(function (x) { if (x.open) x.ids.forEach(function (id) { o[id] = x.kind; }); });
  try { c.put('openIds', JSON.stringify(o), 60); } catch (e) {}
  return o;
}

/* ── 학생회 전달 ── */
function eCouncil() {
  var best = eRows('').filter(function (x) { return x.best; });
  if (!best.length) { SpreadsheetApp.getUi().alert('「우수작」에 체크된 작품이 없어요.'); return; }
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm');
  var root = DriveApp.createFolder('꿈길 우수작 — 학생회 전달 ' + stamp);
  var sub = { photo: root.createFolder('사진전'), reels: root.createFolder('릴스전') }, n = 0, fail = [];
  best.forEach(function (x, i) {
    x.ids.forEach(function (id, j) {
      try {
        var f = DriveApp.getFileById(id), ext = (f.getName().match(/\.\w+$/) || [''])[0];
        var nm = ('0' + (i + 1)).slice(-2) + (x.ids.length > 1 ? '-' + (j + 1) : '') + '_' + x.who + (x.title ? '_' + x.title : '') + ext;
        f.makeCopy(nm.replace(/[\\/:*?"<>|]/g, ''), sub[x.kind]); n++;
      } catch (e) { fail.push(x.who); }
    });
  });
  var list = best.map(function (x, i) { return [('0' + (i + 1)).slice(-2), x.kind === 'reels' ? '릴스전' : '사진전', x.who, x.title, x.team, x.desc]; });
  var ss = SpreadsheetApp.create('우수작 목록 ' + stamp), sh = ss.getSheets()[0];
  sh.appendRow(['번호', '종류', '학생', '제목', '함께 만든 친구', '설명']);
  sh.getRange(2, 1, list.length, 6).setValues(list);
  DriveApp.getFileById(ss.getId()).moveTo(root);
  SpreadsheetApp.getUi().alert('파일 ' + n + '개를 모았어요' + (fail.length ? ' (못 옮긴 것: ' + fail.join(', ') + ')' : '') +
    '.\n폴더: ' + root.getUrl() + '\n\n이 폴더를 학생회 담당 학생 계정에 «보기» 로 공유하세요. 이름은 가운데 글자를 가렸어요.');
}

/* ── 앱 갤러리 (웹 앱) ── */
function doGet(e) {
  var k = (e && e.parameter && e.parameter.k) === 'reels' ? 'reels' : 'photo';
  var items = eRows('').filter(function (x) { return x.open; }).map(function (x) {
    return { kind: x.kind, ids: x.ids, who: x.who, mission: x.mission, title: x.title, desc: x.desc, team: x.team };
  });
  var data = JSON.stringify(items).replace(/</g, '\\u003c');
  var html = E_HTML.replace('__DATA__', function () { return data; })
                   .replace('__K__', k).replace(/__APP__/g, E_APP);
  return HtmlService.createHtmlOutput(html).setTitle('꿈길 걷기 작품 보기')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

/* 썸네일 — 공개 체크된 파일만. 드라이브 thumbnailLink 를 서버가 받아 data: 로 */
function eThumb(id, big) {
  if (!eOpenIds()[id]) return '';
  var key = 'th' + (big ? 'B' : '') + id, c = CacheService.getScriptCache(), hit = c.get(key);
  if (hit) return hit;
  var tok = ScriptApp.getOAuthToken(), H = { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
  var meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink', H).getContentText() || '{}');
  var blob;
  if (meta.thumbnailLink) {
    var r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, big ? '=w1400' : '=w600'), H);
    if (r.getResponseCode() === 200) blob = r.getBlob();
  }
  if (!blob) { try { blob = DriveApp.getFileById(id).getThumbnail(); } catch (e) {} }
  if (!blob) return '';
  var out = 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
  if (out.length < 95000) c.put(key, out, 21600);
  return out;
}

/* 영상 재생 — 누를 때 그 파일만 «학교 계정 + 링크» 보기로 열고 미리보기 주소를 준다 */
function ePlay(id) {
  if (eOpenIds()[id] !== 'reels') return '';
  var f = DriveApp.getFileById(id);
  try { f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return 'https://drive.google.com/file/d/' + id + '/preview';
}

var E_HTML = '<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>' +
  '*{box-sizing:border-box}body{margin:0;font-family:-apple-system,"Apple SD Gothic Neo","Noto Sans KR",sans-serif;background:#F1F5FA;color:#0F172A}' +
  'header{position:sticky;top:0;background:#fff;border-bottom:1px solid #E2E8F0;padding:12px 16px;display:flex;align-items:center;gap:10px;z-index:5}' +
  'header a{color:#2563EB;font-weight:800;text-decoration:none;font-size:15px}header h1{font-size:18px;margin:0;flex:1}' +
  '.tabs{display:flex;gap:8px;padding:12px 16px 0}.tabs button{flex:1;border:0;border-radius:12px;padding:11px;font-size:16px;font-weight:900;background:#E2E8F0;color:#475569}' +
  '.tabs button.on{background:#2563EB;color:#fff}main{padding:8px 16px 40px}h2{font-size:15px;margin:18px 2px 8px;color:#4338CA}' +
  '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}' +
  '.it{background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 1px 3px rgba(15,23,42,.08);cursor:pointer;position:relative}' +
  '.it .im{aspect-ratio:1;background:#E2E8F0 center/cover no-repeat}.it.v .im{aspect-ratio:9/16}' +
  '.it .pl{position:absolute;top:8px;right:8px;background:rgba(0,0,0,.55);color:#fff;border-radius:99px;padding:3px 9px;font-size:12px;font-weight:800}' +
  '.it .tx{padding:8px 10px}.it b{display:block;font-size:14px}.it small{color:#64748B;font-size:12px;font-weight:700}' +
  '.empty{text-align:center;color:#64748B;font-weight:700;padding:60px 10px;line-height:1.7}' +
  '#ov{position:fixed;inset:0;background:rgba(15,23,42,.92);display:none;flex-direction:column;align-items:center;justify-content:center;z-index:9;padding:16px}' +
  '#ov.show{display:flex}#ov img,#ov iframe{max-width:100%;max-height:78vh;border:0;border-radius:12px;background:#000}#ov iframe{width:min(420px,100%);aspect-ratio:9/16;height:auto}' +
  '#ov p{color:#fff;font-weight:700;text-align:center;margin:12px 0 0;line-height:1.5}#ov .x{position:absolute;top:12px;right:16px;font-size:32px;color:#fff;background:none;border:0}' +
  '</style></head><body>' +
  '<header><a href="__APP__" target="_top">‹ 앱으로</a><h1>🖼 꿈길 걷기 작품</h1></header>' +
  '<div class="tabs"><button data-k="photo">📷 사진전</button><button data-k="reels">🎬 릴스전</button></div>' +
  '<main id="m"></main><div id="ov"><button class="x" onclick="cl()">×</button><div id="ovb"></div><p id="ovp"></p></div>' +
  '<script>var D=__DATA__,K="__K__";' +
  'function es(s){return String(s||"").replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}[c]})}' +
  'function card(x,j){var v=x.kind==="reels";return \'<div class="it\'+(v?" v":"")+\'" data-j="\'+j+\'"><div class="im" data-id="\'+x.ids[0]+\'"></div>\'+(v?\'<span class="pl">▶ 재생</span>\':(x.ids.length>1?\'<span class="pl">+\'+(x.ids.length-1)+\'</span>\':""))+' +
  '\'<div class="tx"><b>\'+es(x.title||(v?"릴스":"사진"))+\'</b><small>\'+es(x.who)+(x.team?" 외":"")+\'</small></div></div>\'}' +
  'function draw(){document.querySelectorAll(".tabs button").forEach(function(b){b.className=b.dataset.k===K?"on":""});' +
  'var L=D.map(function(x,j){x.j=j;return x}).filter(function(x){return x.kind===K}),h="";' +
  'if(!L.length){h=\'<div class="empty">아직 공개된 작품이 없어요.<br>선생님이 고른 작품이 여기에 올라와요.</div>\'}' +
  'else if(K==="photo"){var g={};L.forEach(function(x){var m=x.mission||"기타";(g[m]=g[m]||[]).push(x)});' +
  'Object.keys(g).forEach(function(m){h+="<h2>"+es(m)+" · "+g[m].length+\'</h2><div class="grid">\'+g[m].map(function(x){return card(x,x.j)}).join("")+"</div>"})}' +
  'else{h=\'<div class="grid" style="margin-top:12px">\'+L.map(function(x){return card(x,x.j)}).join("")+"</div>"}' +
  'document.getElementById("m").innerHTML=h;document.querySelectorAll(".im").forEach(lazy);' +
  'document.querySelectorAll(".it").forEach(function(c){c.onclick=function(){op(D[c.dataset.j])}})}' +
  'var io=("IntersectionObserver" in window)?new IntersectionObserver(function(es2){es2.forEach(function(e){if(e.isIntersecting){io.unobserve(e.target);th(e.target)}})},{rootMargin:"300px"}):null;' +
  'function lazy(el){io?io.observe(el):th(el)}' +
  'function th(el){google.script.run.withSuccessHandler(function(u){if(u)el.style.backgroundImage=\'url("\'+u+\'")\'}).eThumb(el.dataset.id,false)}' +
  'function op(x){var b=document.getElementById("ovb");document.getElementById("ovp").innerHTML="<b>"+es(x.title)+"</b><br>"+es(x.who)+(x.team?" · "+es(x.team):"")+(x.desc?"<br>"+es(x.desc):"");' +
  'b.innerHTML=\'<p>불러오는 중…</p>\';document.getElementById("ov").classList.add("show");' +
  'if(x.kind==="reels"){google.script.run.withSuccessHandler(function(u){b.innerHTML=u?\'<iframe src="\'+u+\'" allow="autoplay" allowfullscreen></iframe>\':"<p>재생할 수 없어요</p>"}).ePlay(x.ids[0])}' +
  'else{var i=0;var sh=function(){google.script.run.withSuccessHandler(function(u){b.innerHTML=u?\'<img src="\'+u+\'">\'+(x.ids.length>1?\'<p>\'+(i+1)+" / "+x.ids.length+" · 눌러서 다음</p>":""):"<p>사진을 못 불러왔어요</p>";' +
  'var im=b.querySelector("img");if(im&&x.ids.length>1)im.onclick=function(){i=(i+1)%x.ids.length;sh()}}).eThumb(x.ids[i],true)};sh()}}' +
  'function cl(){document.getElementById("ov").classList.remove("show");document.getElementById("ovb").innerHTML=""}' +
  'document.querySelectorAll(".tabs button").forEach(function(b){b.onclick=function(){K=b.dataset.k;draw()}});draw();' +
  '</script></body></html>';
