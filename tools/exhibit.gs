/* ══════════════════════════════════════════════════════════════
   꿈길 걷기 — 사진전·릴스전 출품 시트 (Google Apps Script)

   ★ 흐름: 학생이 앱 「📤 출품하기」 → 구글 폼(파일 업로드) → 이 응답 시트
           → 올라오면 «바로» 앱 갤러리에 뜬다(사용자 결정 10/12: 당일엔 시트·폴더를 오가며 공개 체크할 틈이 없다)
           → 위험한 사진은 «발견한 선생님»이 갤러리에서 「🙈 숨기기」 (teacher숫자@ 계정만 버튼이 보이고, 서버도 막는다)
           → 「우수작」 체크 → 메뉴 한 번으로 학생회 전달 폴더 → 학생회 인스타 좋아요 투표
   ★ 사진은 공유 설정을 건드리지 않는다. 서버(이 스크립트)가 썸네일을 받아 그림 글자로 내려 준다.
     영상 재생만 「학교 계정 + 링크」 보기 권한을 연다(숨기지 않은 영상, 누를 때).

   쓰는 법 (학교 계정, 한 번만)
     1. 사진전 폼·릴스전 폼을 «같은 스프레드시트»에 응답 연결 (탭 이름이나 폼 제목에 「사진」「릴스」)
     2. 그 시트에서 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 저장
     3. 시트 새로고침 → 메뉴 「📸 꿈길 출품」 → ① 체크 칸 만들기
     4. Apps Script → 배포 → 새 배포 → «웹 앱» → 실행: 나 / 액세스: 영종고 사용자 → 웹 앱 URL 을 앱(stamps.json gallery)에
        ★ 코드를 고친 뒤엔 «배포 관리 → ✏ → 새 버전» (새 배포를 또 만들면 주소가 바뀐다)
   ══════════════════════════════════════════════════════════════ */

var E_APP = 'https://today119.github.io/kkumgil/walk/?tab=mis';
var E_COLS = ['숨김', '우수작', '숨긴 선생님'];
var E_TEACHER = /^teacher\d+@yeongjong\.icehs\.kr$/i;     // 숨기기 버튼 — 선생님 계정만

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📸 꿈길 출품')
    .addItem('① 체크 칸 만들기(숨김·우수작)', 'eSetup')
    .addItem('② 학생회 전달 폴더 만들기(우수작)', 'eCouncil')
    .addSeparator()
    .addItem('❓ 사용법', 'eHelp')
    .addToUi();
}

function eHelp() {
  SpreadsheetApp.getUi().alert('📸 꿈길 출품',
    '출품하면 바로 앱 갤러리(학교 계정만 볼 수 있음)에 떠요.\n' +
    '위험한 사진은 선생님 계정으로 갤러리를 열어 「🙈 숨기기」 — 또는 여기서 「숨김」 체크. 체크를 풀면 다시 보여요.\n' +
    '「우수작」 체크 → ② 학생회 전달 폴더 만들기를 누르면 원본 사본이 폴더 하나에 모여요. 그 폴더를 학생회에 공유하세요.\n\n' +
    '갤러리 주소: Apps Script → 배포 → 배포 관리 → 웹 앱 URL', SpreadsheetApp.getUi().ButtonSet.OK);
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
/* 머리줄에 칸이 없으면 맨 오른쪽에 만든다 → 열 번호(1부터) */
function eEnsureCol(sh, name, checkbox) {
  var last = Math.max(sh.getLastColumn(), 1);
  var H = sh.getRange(1, 1, 1, last).getDisplayValues()[0];
  var i = H.indexOf(name);
  if (i < 0) { i = last; sh.getRange(1, i + 1).setValue(name).setFontWeight('bold').setBackground('#FEF3C7'); }
  if (checkbox) sh.getRange(2, i + 1, 800, 1).insertCheckboxes();
  return i + 1;
}

function eSetup() {
  var tabs = eTabs(), msg = [];
  if (!tabs.length) { SpreadsheetApp.getUi().alert('사진전·릴스전 응답 탭을 못 찾았어요. 탭 이름에 「사진」「릴스」를 넣어 주세요.'); return; }
  tabs.forEach(function (t) {
    eEnsureCol(t.sh, '숨김', true); eEnsureCol(t.sh, '우수작', true); eEnsureCol(t.sh, '숨긴 선생님', false);
    msg.push(t.sh.getName() + (t.kind === 'reels' ? ' (릴스전)' : ' (사진전)'));
  });
  eFlush();
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
    var cHide = H.indexOf('숨김'), cBest = H.indexOf('우수작');
    var cNo = col(/학번/), cMail = col(/이메일/), cName = col(/^\d*\.?이름$|^\d*\.?성명$/), cMis = col(/미션|장소|어디서/),
        cTitle = col(/제목/), cDesc = col(/설명|소개/), cTeam = col(/함께|모둠/), cKind = col(/종류/);
    if (cName < 0) cName = col(/이름|성명/);
    var cFile = -1;
    for (var i = 0; i < H.length && cFile < 0; i++)
      for (var r = 1; r < v.length; r++) if (/drive\.google\.com\/open\?id=/.test(String(v[r][i]))) { cFile = i; break; }
    if (cFile < 0) return;
    v.slice(1).forEach(function (r, k) {
      var ids = (String(r[cFile]).match(/id=[\w-]+/g) || []).map(function (s) { return s.slice(3); });
      if (!ids.length) return;
      var mail = cMail >= 0 ? String(r[cMail]) : '', m = mail.match(/s\d\d(\d)(\d\d)\d\d@/), who = '';
      if (!m && cNo >= 0) m = String(r[cNo]).replace(/\D/g, '').match(/^(\d)(\d\d)\d\d$/);     // 학생 계정이 아니면 적어 낸 학번으로
      if (m) who = m[1] + '학년 ' + Number(m[2]) + '반 ';
      who += eMask(cName >= 0 ? r[cName] : '');
      out.push({
        kind: t.kind, tab: t.sh.getName(), row: k + 2, ids: ids, who: who.trim(),
        hidden: cHide >= 0 && r[cHide] === true, best: cBest >= 0 && r[cBest] === true,
        mission: cMis >= 0 ? String(r[cMis]) : '', title: cTitle >= 0 ? String(r[cTitle]) : '',
        desc: cDesc >= 0 ? String(r[cDesc]) : '', team: cTeam >= 0 ? String(r[cTeam]) : '',
        sort: cKind >= 0 ? String(r[cKind]) : ''
      });
    });
  });
  return out;
}
/* 파일 → {kind, hidden} — 썸네일마다 시트를 다시 읽지 않게 1분 캐시(숨기면 바로 지운다) */
function eIds() {
  var c = CacheService.getScriptCache(), hit = c.get('ids2');
  if (hit) return JSON.parse(hit);
  var o = {};
  eRows('').forEach(function (x) { x.ids.forEach(function (id) { o[id] = { k: x.kind, h: x.hidden }; }); });
  try { c.put('ids2', JSON.stringify(o), 60); } catch (e) {}
  return o;
}
function eFlush() { try { CacheService.getScriptCache().remove('ids2'); } catch (e) {} }

function eViewer() { try { return String(Session.getActiveUser().getEmail() || ''); } catch (e) { return ''; } }
function eIsTeacher() { return E_TEACHER.test(eViewer()); }
/* 볼 수 있는 파일인가 — 숨긴 것은 선생님만 */
function eCanSee(id) {
  var x = eIds()[id];
  return !!x && (!x.h || eIsTeacher()) ? x : null;
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
  var T = eIsTeacher();
  var items = eRows('').filter(function (x) { return T || !x.hidden; }).map(function (x) {
    return { kind: x.kind, tab: x.tab, row: x.row, ids: x.ids, who: x.who, mission: x.mission, title: x.title,
             desc: x.desc, team: x.team, sort: x.sort, hidden: x.hidden };
  });
  var data = JSON.stringify({ items: items, k: k, teacher: T, app: E_APP }).replace(/</g, '\\u003c');
  var html = E_HTML.replace('__CLIENT__', function () { return '(' + eClient.toString() + ')(' + data + ');'; });
  return HtmlService.createHtmlOutput(html).setTitle('꿈길 걷기 작품 보기')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

/* 숨기기 / 다시 보이기 — 선생님 계정만. 줄이 정렬로 밀렸을 수 있어 파일 ID 로 줄을 다시 찾는다 */
function eHide(tab, row, id, hide) {
  var me = eViewer();
  if (!E_TEACHER.test(me)) throw new Error('선생님 계정만 숨길 수 있어요.');
  var sh = SpreadsheetApp.getActive().getSheetByName(tab);
  if (!sh) throw new Error('탭을 찾지 못했어요: ' + tab);
  var cHide = eEnsureCol(sh, '숨김', false), cWho = eEnsureCol(sh, '숨긴 선생님', false);
  var v = sh.getDataRange().getValues(), r = row - 1;
  var has = function (i) { return v[i] && v[i].some(function (c) { return String(c).indexOf(id) >= 0; }); };
  if (!has(r)) { r = -1; for (var i = 1; i < v.length; i++) if (has(i)) { r = i; break; } }
  if (r < 1) throw new Error('그 작품 줄을 찾지 못했어요.');
  sh.getRange(r + 1, cHide).insertCheckboxes().setValue(!!hide);
  if (hide) { try { DriveApp.getFileById(id).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) {} }
  sh.getRange(r + 1, cWho).setValue(hide ? me + ' · ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm') : '');
  eFlush();
  return { ok: true, row: r + 1 };
}

/* 썸네일 — 드라이브 thumbnailLink 를 서버가 받아 data: 로 */
function eThumb(id, big) {
  if (!eCanSee(id)) return '';
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

/* 영상 재생 — 누를 때 그 파일만 «학교 계정 + 링크» 보기로 열고 드라이브 미리보기 주소를 준다.
   (폰 기본 플레이어 직접 재생은 불가: 드라이브 다운로드 주소가 Cross-Origin-Resource-Policy: same-site 라 다른 사이트의 <video> 에서 막힘) */
function ePlay(id) {
  var x = eCanSee(id);
  if (!x || x.k !== 'reels') return null;
  try { DriveApp.getFileById(id).setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { v: '', p: 'https://drive.google.com/file/d/' + id + '/preview' };
}

var E_HTML = '<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>' +
  '*{box-sizing:border-box}body{margin:0;font-family:-apple-system,"Apple SD Gothic Neo","Noto Sans KR",sans-serif;background:#F1F5FA;color:#0F172A}' +
  'header{position:sticky;top:0;background:#fff;border-bottom:1px solid #E2E8F0;padding:12px 16px;display:flex;align-items:center;gap:10px;z-index:5}' +
  'header a{color:#2563EB;font-weight:800;text-decoration:none;font-size:15px}header h1{font-size:18px;margin:0;flex:1}' +
  '.tm{background:#FEF3C7;color:#92400E;font-size:13px;font-weight:800;padding:8px 16px}' +
  '.tabs{display:flex;gap:8px;padding:12px 16px 0}.tabs button{flex:1;border:0;border-radius:12px;padding:11px;font-size:16px;font-weight:900;background:#E2E8F0;color:#475569}' +
  '.tabs button.on{background:#2563EB;color:#fff}main{padding:8px 16px 40px}h2{font-size:15px;margin:18px 2px 8px;color:#4338CA}' +
  '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}' +
  '.it{background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 1px 3px rgba(15,23,42,.08);cursor:pointer;position:relative}' +
  '.it.hid{opacity:.45;outline:3px dashed #DC2626}' +
  '.it .im{aspect-ratio:1;background:#E2E8F0 center/cover no-repeat}.it.v .im{aspect-ratio:9/16}' +
  '.it .pl{position:absolute;top:8px;right:8px;background:rgba(0,0,0,.55);color:#fff;border-radius:99px;padding:3px 9px;font-size:12px;font-weight:800}' +
  '.it .hb{position:absolute;top:8px;left:8px;border:0;border-radius:99px;padding:5px 10px;font-size:12px;font-weight:900;background:rgba(255,255,255,.92);color:#B91C1C}' +
  '.it .hb.sure{background:#DC2626;color:#fff}.it.hid .hb{background:#16A34A;color:#fff}' +
  '.it .tx{padding:8px 10px}.it b{display:block;font-size:14px}.it small{color:#64748B;font-size:12px;font-weight:700}' +
  '.empty{text-align:center;color:#64748B;font-weight:700;padding:60px 10px;line-height:1.7}' +
  '#ov{position:fixed;inset:0;background:rgba(15,23,42,.92);display:none;flex-direction:column;align-items:center;justify-content:center;z-index:9;padding:16px}' +
  '#ov.show{display:flex}#ov img,#ov iframe{max-width:100%;max-height:78vh;border:0;border-radius:12px;background:#000}#ovb{width:100%;display:flex;flex-direction:column;align-items:center}.vp{position:relative;width:calc(100vw - 32px);max-width:480px;display:flex;flex-direction:column;align-items:center;gap:14px}.vt{height:56vh;aspect-ratio:9/16;max-width:100%;border-radius:12px;background:#000 center/contain no-repeat}.pbtn{display:block;width:100%;text-align:center;background:#EF4444;color:#fff;font-size:20px;font-weight:900;padding:16px;border-radius:14px;text-decoration:none}.vh{color:#CBD5E1;font-size:12px;margin:6px 0 0;text-align:center}' +
  '#ov p{color:#fff;font-weight:700;text-align:center;margin:12px 0 0;line-height:1.5}#ov .x{position:absolute;top:12px;right:16px;font-size:32px;color:#fff;background:none;border:0}' +
  '#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#0F172A;color:#fff;font-weight:800;padding:10px 16px;border-radius:12px;display:none;z-index:20}' +
  '</style></head><body>' +
  '<header><a id="back" target="_top">‹ 앱으로</a><h1>🖼 꿈길 걷기 작품</h1></header><div id="tm"></div>' +
  '<div class="tabs"><button data-k="photo">📷 사진전</button><button data-k="reels">🎬 릴스전</button></div>' +
  '<main id="m"></main><div id="ov"><button class="x" id="ovx">×</button><div id="ovb"></div><p id="ovp"></p></div><div id="toast"></div>' +
  '<script>__CLIENT__</script></body></html>';

/* 브라우저에서 도는 쪽 — doGet 이 이 함수 소스를 그대로 페이지에 넣는다(문자열 이스케이프 지옥 피하기) */
function eClient(S) {
  var D = S.items, K = S.k, T = S.teacher, $ = function (id) { return document.getElementById(id); };
  $('back').href = S.app;
  if (T) $('tm').innerHTML = '<div class="tm">👩‍🏫 선생님 모드 — 위험한 사진은 「🙈 숨기기」(두 번 눌러 확정). 숨긴 것은 학생에게 안 보여요.</div>';
  else $('tm').innerHTML = '<div class="tm" style="background:#FEF2F2;color:#991B1B">⚠️ 친구 몰래 찍은·우스운 사진, 동의 없는 얼굴, 장난 사진, 개인정보·퍼 온 사진은 선생님이 숨겨요. 캡처해서 퍼뜨리지 마세요.</div>';
  function es(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function toast(m) { var t = $('toast'); t.textContent = m; t.style.display = 'block'; clearTimeout(toast.k); toast.k = setTimeout(function () { t.style.display = 'none'; }, 2200); }
  function card(x, j) {
    var v = x.kind === 'reels';
    return '<div class="it' + (v ? ' v' : '') + (x.hidden ? ' hid' : '') + '" data-j="' + j + '"><div class="im" data-id="' + x.ids[0] + '"></div>' +
      (v ? '<span class="pl">▶ 재생</span>' : (x.ids.length > 1 ? '<span class="pl">+' + (x.ids.length - 1) + '</span>' : '')) +
      (T ? '<button class="hb" data-h="' + j + '">' + (x.hidden ? '👁 다시 보이기' : '🙈 숨기기') + '</button>' : '') +
      '<div class="tx"><b>' + es(x.title || (v ? '릴스' : '사진')) + '</b><small>' + es(x.who) + (x.team ? ' 외' : '') + (x.hidden ? ' · 숨김' : '') + '</small></div></div>';
  }
  function draw() {
    document.querySelectorAll('.tabs button').forEach(function (b) { b.className = b.dataset.k === K ? 'on' : ''; });
    var L = D.map(function (x, j) { x.j = j; return x; }).filter(function (x) { return x.kind === K && (T || !x.hidden); }), h = '';
    if (!L.length) h = '<div class="empty">아직 올라온 작품이 없어요.<br>출품하면 여기에 바로 떠요.</div>';
    else if (K === 'photo') {
      var g = {}; L.forEach(function (x) { var m = x.mission || '기타'; (g[m] = g[m] || []).push(x); });
      Object.keys(g).forEach(function (m) { h += '<h2>' + es(m) + ' · ' + g[m].length + '</h2><div class="grid">' + g[m].map(function (x) { return card(x, x.j); }).join('') + '</div>'; });
    } else h = '<div class="grid" style="margin-top:12px">' + L.map(function (x) { return card(x, x.j); }).join('') + '</div>';
    $('m').innerHTML = h;
    document.querySelectorAll('.im').forEach(lazy);
    document.querySelectorAll('.it').forEach(function (c) { c.onclick = function () { op(D[c.dataset.j]); }; });
    document.querySelectorAll('.hb').forEach(function (b) {
      b.onclick = function (ev) {
        ev.stopPropagation();
        var x = D[b.dataset.h], hide = !x.hidden;
        if (hide && !b.classList.contains('sure')) { b.classList.add('sure'); b.textContent = '정말 숨길까요?'; setTimeout(function () { if (b.isConnected && b.classList.contains('sure')) { b.classList.remove('sure'); b.textContent = '🙈 숨기기'; } }, 3000); return; }
        b.disabled = true; b.textContent = '…';
        google.script.run.withSuccessHandler(function () { x.hidden = hide; toast(hide ? '숨겼어요 — 학생에게 안 보여요' : '다시 보이게 했어요'); draw(); })
          .withFailureHandler(function (e) { toast(e.message || '실패했어요'); draw(); })
          .eHide(x.tab, x.row, x.ids[0], hide);
      };
    });
  }
  var io = ('IntersectionObserver' in window) ? new IntersectionObserver(function (es2) { es2.forEach(function (e) { if (e.isIntersecting) { io.unobserve(e.target); th(e.target); } }); }, { rootMargin: '300px' }) : null;
  function lazy(el) { io ? io.observe(el) : th(el); }
  function th(el) { google.script.run.withSuccessHandler(function (u) { if (u) el.style.backgroundImage = 'url("' + u + '")'; }).eThumb(el.dataset.id, false); }
  function op(x) {
    var b = $('ovb');
    $('ovp').innerHTML = '<b>' + es(x.title) + '</b><br>' + es(x.who) + (x.team ? ' · ' + es(x.team) : '') + (x.desc ? '<br>' + es(x.desc) : '');
    b.innerHTML = '<p>불러오는 중…</p>'; $('ov').classList.add('show');
    if (x.kind === 'reels') {
      google.script.run.withSuccessHandler(function (r) {
        if (!r || !r.p) { b.innerHTML = '<p>재생할 수 없어요</p>'; return; }
        var url = r.p.replace('/preview', '/view');
        b.innerHTML = '<div class="vp"><div class="vt" id="vt"></div><a class="pbtn" href="' + url + '" target="_blank" rel="noopener">▶ 영상 재생</a></div><p class="vh">새 창에서 열려요 · 다 보면 창을 닫고 돌아오세요</p>';
        google.script.run.withSuccessHandler(function (u) { var t = $('vt'); if (t && u) t.style.backgroundImage = 'url("' + u + '")'; }).eThumb(x.ids[0], true);
      }).withFailureHandler(function () { b.innerHTML = '<p>재생할 수 없어요</p>'; }).ePlay(x.ids[0]);
    } else {
      var i = 0, sh = function () {
        google.script.run.withSuccessHandler(function (u) {
          b.innerHTML = u ? '<img src="' + u + '">' + (x.ids.length > 1 ? '<p>' + (i + 1) + ' / ' + x.ids.length + ' · 눌러서 다음</p>' : '') : '<p>사진을 못 불러왔어요</p>';
          var im = b.querySelector('img'); if (im && x.ids.length > 1) im.onclick = function () { i = (i + 1) % x.ids.length; sh(); };
        }).eThumb(x.ids[i], true);
      };
      sh();
    }
  }
  $('ovx').onclick = function () { $('ov').classList.remove('show'); $('ovb').innerHTML = ''; };
  document.querySelectorAll('.tabs button').forEach(function (b) { b.onclick = function () { K = b.dataset.k; draw(); }; });
  draw();
}
