/* ══════════════════════════════════════════════════════════════
   꿈길 걷기 — 인원 체크 폼 만들기 (Google Apps Script)

   ★ 왜 이렇게 만드나
     · 걸으면서 글자를 쓰기 어렵다 → 폼은 «전부 누르기만» 하게 만든다(반 → 이름 → 상황 → 이유 → 어디쯤).
     · 학생 명단은 «학교 구글 계정 안에만» 둔다. 꿈길 앱(공개 웹)에는 이름이 하나도 들어가지 않는다.
     · 출발 인원은 따로 세지 않는다. 출발 전에 «불참»만 찍으면 「명단 − 불참 = 참가」로 계산된다.
     · 한 학생의 «마지막» 기록이 지금 상태다(중도포기 뒤 「다시 합류」를 찍으면 다시 걷는 인원으로 센다).

   ★ 쓰는 법 (학교 계정으로, 컴퓨터에서 한 번만)
     1. 구글 드라이브 → 새로 만들기 → 구글 시트 (이름: 꿈길 걷기 인원 체크)
     2. 확장 프로그램 → Apps Script → 이 파일 내용을 통째로 붙여넣고 💾 저장
     3. 시트로 돌아와 새로고침 → 메뉴에 「🚶 꿈길 걷기」가 생긴다
     4. 🚶 꿈길 걷기 → 「① 명단 탭 만들기」 → 명단_1일차 / 명단_2일차 탭에 반·번호·이름을 붙여넣기
     5. 🚶 꿈길 걷기 → 「② 1일차 폼 만들기」(처음 한 번은 권한 허용 창이 뜬다 → 허용)
     6. 「안내」 탭에 폼 주소가 생긴다 → 선생님들께 그 주소(또는 QR)를 보내면 끝
     집계_1일차 탭은 폼이 제출될 때마다 저절로 새로 고쳐진다.
   ══════════════════════════════════════════════════════════════ */

var DAYS = {
  1: { title: '1일차 (10/30 금) · 영종도',
       places: ['출발 전', '백운산', '용궁사', '중산교차로', '박석공원', '영종진', '하늘신도시(점심)',
                '씨사이드파크', '족욕장', '학교 도착', '기타'] },
  2: { title: '2일차 (10/31 토) · 무의도',
       places: ['출발 전', '무의도 공영주차장', '잠진도', '마시안해변', '선녀바위', '을왕(점심)',
                '용유초', '기타'] }
};
var STATUS  = ['불참', '중도포기', '다시 합류'];
var REASONS = ['못 걷겠어요', '다쳤어요', '아파요', '결석·개인 사정', '기타'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚶 꿈길 걷기')
    .addItem('① 명단 탭 만들기', 'makeRosterTabs')
    .addSeparator()
    .addItem('② 1일차 폼 만들기', 'makeForm1')
    .addItem('② 2일차 폼 만들기', 'makeForm2')
    .addSeparator()
    .addItem('집계 새로 고치기', 'refreshAll')
    .addItem('📥 참가 조사 가져오기', 'importSurvey')
    .addItem('명단 확인하기', 'checkRoster')
    .addSeparator()
    .addItem('🔑 현황판 비번 정하기', 'setPin')
    .addItem('📱 현황판 주소 보기', 'showDashUrl')
    .addSeparator()
    .addItem('❓ 사용법', 'showHelp')
    .addToUi();
}

/* ❓ 사용법 — 메뉴마다 무엇을 하는지 팝업으로 */
function showHelp() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:sans-serif;font-size:14px;line-height:1.65;color:#0F172A;padding:4px 10px}' +
    'h3{margin:14px 0 4px;font-size:15px;color:#0369A1}b{color:#B45309}.k{background:#F1F5F9;border-radius:8px;padding:8px 10px;margin:6px 0}</style>' +
    '<p>이 시트는 <b>꿈길 걷기 당일 인원 체크</b>용이에요. 선생님들은 <b>폼에서 누르기만</b> 하고, 집계는 시트가 알아서 해요.</p>' +
    '<h3>① 명단 탭 만들기</h3>명단_1일차 · 명단_2일차 탭을 만들어요. 첫 줄 머리글은 <b>반 · 번호 · 이름</b>.' +
    '<div class="k">10개 반을 <b>옆으로 나란히</b>(반·번호·이름 묶음을 반마다) 붙여도, <b>세로로 이어</b> 붙여도 돼요.<br>반 칸은 그 반 첫 줄에만 적어도 돼요.</div>' +
    '<h3>명단 확인하기</h3>시트가 명단을 <b>어떻게 읽었는지</b> 반별 인원으로 보여 줘요. 폼 만들기 전에 한 번 눌러 보세요.' +
    '<h3>② 1일차 / 2일차 폼 만들기</h3>명단으로 <b>구글 폼을 자동으로</b> 만들어요.' +
    '<div class="k">반 고르기 → <b>그 반 명단</b>(여러 명 체크 가능) → 상황[불참/중도포기/다시 합류] → 이유 → 어디쯤 → 제출<br>' +
    '시각과 입력한 선생님은 자동 기록 · 학교 계정만 열림</div>' +
    '폼 주소와 QR 은 <b>「안내」 탭</b>에 생겨요. 교직원 단톡방에 올리면 끝. <b>배포는 필요 없어요.</b>' +
    '<h3>📱 현황판 (선생님 폰)</h3>꿈길 앱 홈의 「선생님용 · 학생 확인」 → 반 카드 → <b>학생 이름 버튼</b> → 중도포기·불참·다시 걷는 중. 폼 없이 여기서 바로 기록돼요(「기록_1일차」 탭).' +
    '<h3>집계 (자동)</h3>폼이 제출될 때마다 <b>집계_1일차</b> 탭이 새로 고쳐져요.' +
    '<div class="k">반별: 명단 · 불참 · 참가 · 중도포기 · <b>지금 걷는 인원</b><br>빠진 학생 목록(시각·어디쯤·이유) · 전체 기록</div>' +
    '한 학생의 <b>마지막 기록</b>이 지금 상태예요. 중도포기했다가 다시 걸으면 「다시 합류」를 누르면 돼요.' +
    '<h3>집계 새로 고치기</h3>자동 갱신이 늦을 때 손으로 다시 계산해요.' +
    '<h3>⚠️ 알아 둘 것</h3>폼을 만든 뒤 <b>명단을 고치면 ② 를 다시</b> 눌러 폼을 새로 만들어야 해요.'
  ).setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '🚶 꿈길 걷기 인원 체크 — 사용법');
}

/* 명단 확인하기 — 폼 만들기 전에 «어떻게 읽었는지» 보여 준다 */
function checkRoster() {
  var ui = SpreadsheetApp.getUi(), out = [];
  [1, 2].forEach(function (d) {
    try {
      var R = readRoster(d), n = 0;
      var line = R.classes.map(function (c) { n += R.byClass[c].length; return classLabel(c) + ' ' + R.byClass[c].length + '명'; });
      out.push('【' + d + '일차】 ' + R.classes.length + '개 반 · ' + n + '명\n' + line.join(' / ') +
               (R.excluded && R.excluded.length ? '\n뺀 학생 ' + R.excluded.length + '명(자퇴·전출 등): ' + R.excluded.join(', ') : ''));
    } catch (e) { out.push('【' + d + '일차】 ' + e.message); }
  });
  ui.alert('명단 확인', out.join('\n\n') + '\n\n맞으면 「② 폼 만들기」를 누르세요.', ui.ButtonSet.OK);
}

/* ══════════════════════════════════════════════════════════════
   📥 참가 조사 가져오기 — 담임 선생님들이 채운 「꿈길걷기_참가조사_담임입력용」 시트에서
   ★ 반별 탭(1-1 … 2-10)마다:  번호 | 이름 | 1일차 참가(체크) | 불참 사유 | 2일차 신청(체크) | 티셔츠
   · 1일차 = 1학년만 걷는다(2학년은 인천바로알기). 1학년 전원을 명단에 넣고, 체크가 풀린 학생은
     현황판 기록에 「불참」(사유 그대로)으로 미리 넣는다 → 아침에 선생님이 따로 안 찍어도 된다.
   · 2일차 = 희망자만. 1·2학년 모든 반에서 체크한 학생만 모아 명단을 만든다.
   · 몇 번을 눌러도 같다 — 지난번에 넣은 「참가조사」 기록은 지우고 새로 넣는다(선생님이 현장에서 누른 건 그대로).
   · 조사 시트 주소는 처음 한 번만 묻고 기억한다(코드에 안 적는다 — 깃허브가 공개라서).
   ══════════════════════════════════════════════════════════════ */
var SURVEY_TAG = '참가조사(자동)';
function importSurvey() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getDocumentProperties();
  var src = props.getProperty('surveyId');
  var r = ui.prompt('📥 참가 조사 가져오기',
    '담임 입력용 참가 조사 시트 주소(또는 ID)를 붙여넣어 주세요.' + (src ? '\n비워 두면 지난번 시트를 다시 읽어요.' : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = r.getResponseText().trim();
  if (t) { var m = t.match(/\/d\/([\w-]{20,})/); src = m ? m[1] : t; }
  if (!src) { ui.alert('시트 주소가 없어요.'); return; }
  var ss;
  try { ss = SpreadsheetApp.openById(src); } catch (e) { ui.alert('시트를 열 수 없어요. 주소와 공유 권한을 확인해 주세요.\n' + e.message); return; }
  props.setProperty('surveyId', src);
  var S = readSurvey(ss.getSheets().map(function (sh) { return { name: sh.getName(), grid: sh.getDataRange().getDisplayValues() }; }));
  if (!S.d1.length && !S.d2.length) { ui.alert('반별 탭(1-1, 2-3 …)에서 「이름」 머리줄을 못 찾았어요.'); return; }
  var me = SpreadsheetApp.getActive();
  writeRoster(me, '명단_1일차', S.d1);
  writeRoster(me, '명단_2일차', S.d2);
  var n = replaceSurveyAbsences(S.d1.filter(function (s) { return !s.go1; }));
  try { refreshAll(); } catch (e) {}
  var by = function (list) { var o = {}; list.forEach(function (s) { o[s.cls] = (o[s.cls] || 0) + 1; }); return Object.keys(o).sort(classOrder).map(function (c) { return c + ' ' + o[c]; }).join(' · '); };
  ui.alert('📥 가져왔어요',
    '【1일차】 1학년 ' + S.d1.length + '명 · 불참 ' + n + '명' + (n ? ' (현황판에 「불참」으로 미리 넣었어요)' : '') + '\n' +
    S.d1.filter(function (s) { return !s.go1; }).map(function (s) { return s.cls + ' ' + s.name + (s.why ? '(' + s.why + ')' : ''); }).join(', ') + '\n\n' +
    '【2일차】 신청 ' + S.d2.length + '명\n' + by(S.d2) + '\n\n' +
    '조사 시트가 바뀌면 이 버튼을 다시 누르면 돼요.', ui.ButtonSet.OK);
}

/* 반별 탭 읽기 — 탭 이름이 「1-3」 꼴인 것만. 머리줄은 「이름」 칸이 있는 줄 */
function readSurvey(tabs) {
  var d1 = [], d2 = [];
  tabs.forEach(function (tb) {
    var cls = String(tb.name).replace(/\s+/g, '');
    if (!/^\d+-\d+$/.test(cls)) return;
    var g = tb.grid, hr = -1;
    for (var i = 0; i < g.length; i++) if (g[i].some(function (v) { return String(v).trim() === '이름'; })) { hr = i; break; }
    if (hr < 0) return;
    var H = g[hr].map(function (v) { return String(v).replace(/\s+/g, ''); });
    var col = function (re) { for (var k = 0; k < H.length; k++) if (re.test(H[k])) return k; return -1; };
    var cNo = col(/^번호/), cNm = H.indexOf('이름'), c1 = col(/^1일차/), cWhy = col(/^불참사유/), c2 = col(/^2일차/);
    var grade = cls.split('-')[0];
    g.slice(hr + 1).forEach(function (row) {
      var nm = String(row[cNm] || '').trim();
      if (!nm) return;
      var yes = function (c) { return c >= 0 && /^(TRUE|O|○|◯|V|✓|✔|Y|예|참가|신청)$/i.test(String(row[c]).trim()); };
      var s = { cls: cls, no: Number(row[cNo]) || 0, name: nm, go1: yes(c1), go2: yes(c2), why: cWhy >= 0 ? String(row[cWhy] || '').trim() : '' };
      if (grade === '1') d1.push(s);
      if (s.go2) d2.push(s);
    });
  });
  return { d1: d1, d2: d2 };
}

function writeRoster(ss, name, list) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange('A:A').setNumberFormat('@');
  var rows = [['반', '번호', '이름']].concat(list.slice().sort(function (a, b) { return classOrder(a.cls, b.cls) || a.no - b.no; })
    .map(function (s) { return [s.cls, s.no, s.name]; }));
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E0F2FE');
  sh.setFrozenRows(1);
}

/* 1일차 불참을 「기록_1일차」에 넣는다 — 지난번 자동 기록은 지우고 */
function replaceSurveyAbsences(list) {
  var sh = logSheet(1, true), last = sh.getLastRow();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, LOG_HEAD.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) if (vals[i][6] === SURVEY_TAG) sh.deleteRow(i + 2);
  }
  if (!list.length) return 0;
  var R = readRoster(1), when = new Date(2026, 9, 30, 7, 0);   // 행사날 아침 7시로 적어 둔다 — 현장 기록보다 늘 앞서게
  var rows = list.map(function (s) {
    var lab = s.name;
    var cl = R.byClass[s.cls];
    if (cl) { var hit = cl.filter(function (x) { return x.name === s.name && x.no === s.no; })[0]; if (hit) lab = choiceLabel(cl, hit); }
    return [when, s.cls, lab, '불참', s.why || '', '출발 전', SURVEY_TAG];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_HEAD.length).setValues(rows);
  return rows.length;
}

/* ── ① 명단 탭 ── */
function makeRosterTabs() {
  var ss = SpreadsheetApp.getActive();
  [1, 2].forEach(function (d) {
    var name = '명단_' + d + '일차';
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 3).setValues([['반', '번호', '이름']]).setFontWeight('bold').setBackground('#E0F2FE');
    sh.getRange(2, 1, 1, 3).setValues([[d === 1 ? '1' : '1-1', '1', '(예시 — 지우고 붙여넣기)']]).setFontColor('#94A3B8');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');   // 「1-3」이 날짜로 바뀌지 않게
  });
  SpreadsheetApp.getUi().alert('명단 탭을 만들었어요',
    '명단_1일차 / 명단_2일차 탭에 반 · 번호 · 이름을 붙여넣으세요.\n\n' +
    '· 10개 반을 옆으로 나란히(반·번호·이름 묶음을 반마다) 붙여도, 세로로 이어 붙여도 돼요.\n' +
    '· 반 칸은 그 반 첫 줄에만 적어도 돼요.\n' +
    '· 2일차는 학년이 섞이니 반을 「1-3」, 「2-5」처럼 적어 주세요.\n\n' +
    '다 붙였으면 「명단 확인하기」 → 「② 폼 만들기」 순서로 누르세요.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/* 명단 읽기 — 붙이는 모양을 가리지 않는다
     · 옆으로 나란히: [반|번호|이름] [반|번호|이름] … 묶음 여러 개
     · 세로로 이어서: 한 묶음 아래로 반마다 쭉
     · 반 칸이 그 반 첫 줄에만 있고 아래가 비어 있어도 위의 반을 이어받는다
     · 중간에 다시 나온 「반·번호·이름」 머리줄, 빈 줄, 예시 줄은 건너뛴다 */
function readRoster(day) {
  var sh = SpreadsheetApp.getActive().getSheetByName('명단_' + day + '일차');
  if (!sh) throw new Error('명단_' + day + '일차 탭이 없어요. 「① 명단 탭 만들기」부터 해 주세요.');
  return parseRoster(sh.getDataRange().getDisplayValues(), day);
}
function parseRoster(grid, day) {
  var M = parseMyeongryeol(grid);          // 학교 「학생명렬표」를 통째로 붙여넣은 경우
  if (M) return M;
  var head = (grid[0] || []).map(function (v) { return String(v).trim(); });
  var blocks = [];
  head.forEach(function (v, c) {
    if (v !== '반') return;
    var nx = head.indexOf('반', c + 1); if (nx < 0) nx = head.length;
    var no = -1, nm = -1;
    for (var k = c + 1; k < nx; k++) { if (no < 0 && head[k] === '번호') no = k; if (nm < 0 && head[k] === '이름') nm = k; }
    if (nm >= 0) blocks.push({ c: c, no: no, nm: nm });
  });
  if (!blocks.length) throw new Error('명단_' + day + '일차 탭 첫 줄에 「반」「번호」「이름」 머리줄이 없어요.');
  var classes = [], byClass = {}, seen = {};
  blocks.forEach(function (b) {
    var cur = '';
    grid.slice(1).forEach(function (r) {
      var cls = String(r[b.c] || '').trim().replace(/반$/, ''), name = String(r[b.nm] || '').trim();
      if (cls === '반' || name === '이름') { cur = ''; return; }          // 중간 머리줄
      if (cls) cur = cls;
      if (!name || name.indexOf('예시') >= 0 || !cur) return;
      var no = b.no >= 0 ? Number(r[b.no]) || 0 : 0;
      var key = cur + '|' + no + '|' + name;
      if (seen[key]) return; seen[key] = 1;
      if (!byClass[cur]) { byClass[cur] = []; classes.push(cur); }
      byClass[cur].push({ no: no, name: name });
    });
  });
  if (!classes.length) throw new Error('명단_' + day + '일차 탭에 학생이 없어요.');
  classes.sort(classOrder);
  classes.forEach(function (c) { byClass[c].sort(function (a, b) { return a.no - b.no; }); });
  return { classes: classes, byClass: byClass };
}
/* 학교 「학생명렬표」 모양 그대로 읽기
     반 | | 1-1 | | 1-2 | …        ← 「반」 줄: 학급 이름이 두 칸마다
     번 호 | | 성 명 | 성별 | 성 명 | 성별 …
     1 | | 김다은 | 여 | 강율 | 남 …
   ★ 성별 칸이 「남·여」가 아닌 학생(자퇴·전출·공석 등)은 지금 학교에 없으므로 뺀다.
   ★ 아래 「남·여·계」 합계 줄은 번호가 아니라서 저절로 건너뛴다. */
function parseMyeongryeol(grid) {
  var T = function (v) { return String(v == null ? '' : v).replace(/\s+/g, ''); };
  var hr = -1;
  for (var i = 0; i < Math.min(grid.length, 10); i++) {
    var row = grid[i].map(T);
    if (row.indexOf('반') >= 0 && row.some(function (v) { return /^\d+-\d+$/.test(v); })) { hr = i; break; }
  }
  if (hr < 0) return null;
  var cols = [];
  grid[hr].forEach(function (v, c) { var t = T(v); if (/^\d+-\d+$/.test(t)) cols.push({ c: c, cls: t }); });
  var classes = [], byClass = {}, excluded = [];
  grid.slice(hr + 1).forEach(function (r) {
    var no = T(r[0]);
    if (!/^\d+$/.test(no)) return;
    cols.forEach(function (k) {
      var nm = String(r[k.c] == null ? '' : r[k.c]).trim(), sx = T(r[k.c + 1]);
      if (!nm) return;
      if (sx !== '남' && sx !== '여') { excluded.push(k.cls + ' ' + no + '번 ' + nm + '(' + (sx || '?') + ')'); return; }
      if (!byClass[k.cls]) { byClass[k.cls] = []; classes.push(k.cls); }
      byClass[k.cls].push({ no: Number(no), name: nm });
    });
  });
  if (!classes.length) return null;
  classes.sort(classOrder);
  return { classes: classes, byClass: byClass, excluded: excluded };
}

/* 「2」 < 「10」, 「1-3」 < 「2-1」 — 글자 순이 아니라 숫자 순으로 */
function classOrder(a, b) {
  var pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  for (var i = 0; i < Math.max(pa.length, pb.length); i++) {
    var x = pa[i] || 0, y = pb[i] || 0;
    if (x !== y) return x - y;
  }
  return a < b ? -1 : 1;
}
/* 같은 반에 동명이인이 있으면 번호를 붙여 구분한다 */
function choiceLabel(list, s) {
  var dup = list.filter(function (x) { return x.name === s.name; }).length > 1;
  return dup ? s.name + ' (' + s.no + '번)' : s.name;
}
function classLabel(c) { return /^\d+$/.test(c) ? c + '반' : c + '반'; }

/* ── ② 폼 만들기 ── */
function makeForm1() { makeForm(1); }
function makeForm2() { makeForm(2); }

function makeForm(day) {
  var ss = SpreadsheetApp.getActive(), props = PropertiesService.getDocumentProperties();
  var ui = SpreadsheetApp.getUi();
  if (props.getProperty('form' + day)) {
    var ok = ui.alert(day + '일차 폼이 이미 있어요. 새로 만들까요?\n(예전 폼은 지워지지 않고 연결만 끊겨요)', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  }
  var R = readRoster(day), D = DAYS[day];

  var form = FormApp.create('꿈길 걷기 인원 체크 — ' + D.title)
    .setDescription('누르기만 하면 돼요. 반 → 이름 → 상황 → 이유 → 어디쯤 → 제출.\n시각은 자동으로 기록됩니다.')
    .setConfirmationMessage('기록했어요 ✅  다른 학생도 있으면 아래 「다른 응답 제출」을 누르세요.')
    .setShowLinkToRespondAgain(true)
    .setAllowResponseEdits(false)
    .setProgressBar(false);
  try { form.setRequireLogin(true); } catch (e) {}     // 학교 계정만
  try { form.setCollectEmail(true); } catch (e) {}     // 누가 찍었는지

  // 1쪽: 반 고르기 → 그 반 쪽으로
  var classItem = form.addMultipleChoiceItem().setTitle('반').setRequired(true);
  var pages = {};
  R.classes.forEach(function (c) {
    var pg = form.addPageBreakItem().setTitle(classLabel(c));
    var list = R.byClass[c];
    form.addCheckboxItem()
      .setTitle('이름 · ' + classLabel(c))
      .setHelpText('여러 명을 한꺼번에 골라도 돼요')
      .setChoiceValues(list.map(function (s) { return choiceLabel(list, s); }))
      .setRequired(true);
    pages[c] = pg;
  });
  // 마지막 쪽: 상황·이유·어디쯤 (모든 반이 같이 쓴다)
  var common = form.addPageBreakItem().setTitle('상황');
  form.addMultipleChoiceItem().setTitle('상황').setChoiceValues(STATUS).setRequired(true);
  form.addMultipleChoiceItem().setTitle('이유').setChoiceValues(REASONS).setRequired(false);
  form.addMultipleChoiceItem().setTitle('어디쯤').setChoiceValues(D.places).setRequired(false);
  R.classes.forEach(function (c) { pages[c].setGoToPage(common); });
  classItem.setChoices(R.classes.map(function (c) { return classItem.createChoice(classLabel(c), pages[c]); }));

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty('form' + day, form.getId());

  // 응답 탭 이름 정리 + 폼 제출 때마다 집계
  SpreadsheetApp.flush();
  ss.getSheets().forEach(function (sh) {
    var u = sh.getFormUrl();
    if (u && u.indexOf(form.getId()) >= 0) sh.setName('응답_' + day + '일차');
  });
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'onSubmit'; });
  if (!has) ScriptApp.newTrigger('onSubmit').forSpreadsheet(ss).onFormSubmit().create();

  writeGuide(day, form);
  refresh(day);
  ui.alert(day + '일차 폼을 만들었어요 ✅',
    '· 「안내」 탭에 폼 주소와 QR 이 있어요 → 교직원 단톡방에 올려 주세요.\n' +
    '· 응답은 「응답_' + day + '일차」 탭, 집계는 「집계_' + day + '일차」 탭에 저절로 쌓여요.\n' +
    '· 시험 삼아 한 번 제출해 보고, 집계에 뜨는지 확인해 보세요(시험 응답은 폼 → 응답 탭에서 지울 수 있어요).',
    ui.ButtonSet.OK);
}

function writeGuide(day, form) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('안내') || ss.insertSheet('안내', 0);
  var url = form.getPublishedUrl();
  var row = day === 1 ? 2 : 6;
  sh.getRange(1, 1).setValue('선생님용 폼 주소 — 이 주소나 QR 을 교직원 단톡방에 올려 주세요').setFontWeight('bold');
  sh.getRange(row, 1, 3, 1).setValues([[DAYS[day].title], [url], ['']]);
  sh.getRange(row, 1).setFontWeight('bold').setFontSize(13);
  sh.getRange(row + 2, 1).setFormula('=IMAGE("https://quickchart.io/qr?size=220&text=" & ENCODEURL(A' + (row + 1) + '))');
  sh.setRowHeight(row + 2, 230);
  sh.setColumnWidth(1, 420);
}

/* ── 집계 ── */
function onSubmit() { refreshAll(); }
function refreshAll() {
  [1, 2].forEach(function (d) {
    var has = PropertiesService.getDocumentProperties().getProperty('form' + d) || logSheet(d, false);
    if (has) { try { refresh(d); } catch (e) {} }
  });
}

function readEvents(day) {
  return readFormEvents(day).concat(readLog(day));
}
/* 현황판에서 누른 기록 — 「기록_N일차」 탭 */
var LOG_HEAD = ['시각', '반', '이름', '상황', '이유', '어디쯤', '입력한 선생님'];
function logSheet(day, make) {
  var ss = SpreadsheetApp.getActive(), name = '기록_' + day + '일차';
  var sh = ss.getSheetByName(name);
  if (!sh && make) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]).setFontWeight('bold').setBackground('#E0F2FE');
    sh.setFrozenRows(1);
    sh.getRange('B:C').setNumberFormat('@');
  }
  return sh;
}
function readLog(day) {
  var sh = logSheet(day, false);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEAD.length).getValues()
    .filter(function (r) { return r[0] && r[1] && r[2]; })
    .map(function (r) {
      return { time: new Date(r[0]), cls: String(r[1]), names: [String(r[2])], status: String(r[3]),
               reason: String(r[4] || ''), place: String(r[5] || ''), by: String(r[6] || '') };
    });
}
function readFormEvents(day) {
  var id = PropertiesService.getDocumentProperties().getProperty('form' + day);
  if (!id) return [];
  var form;
  try { form = FormApp.openById(id); } catch (e) { return []; }
  return form.getResponses().map(function (r) {
    var e = { time: r.getTimestamp(), by: '', cls: '', names: [], status: '', reason: '', place: '' };
    try { e.by = r.getRespondentEmail(); } catch (x) {}
    r.getItemResponses().forEach(function (ir) {
      var t = ir.getItem().getTitle(), v = ir.getResponse();
      if (t === '반') e.cls = String(v).replace(/반$/, '');
      else if (t.indexOf('이름') === 0) e.names = [].concat(v);
      else if (t === '상황') e.status = v;
      else if (t === '이유') e.reason = v || '';
      else if (t === '어디쯤') e.place = v || '';
    });
    return e;
  });
}

/* 명단 + 기록 → 학생별 현재 상태 (같은 학생은 마지막 기록이 이긴다) */
function computeStatus(R, events) {
  events = events.slice().sort(function (a, b) { return a.time - b.time; });
  var now = {};   // key 반|이름표시 → 마지막 기록
  events.forEach(function (e) {
    e.names.forEach(function (n) { now[e.cls + '|' + n] = e; });
  });
  var table = [], people = [];
  R.classes.forEach(function (c) {
    var list = R.byClass[c], absent = 0, drop = 0;
    list.forEach(function (s) {
      var lab = choiceLabel(list, s), e = now[c + '|' + lab];
      var st = !e ? '걷는 중' : e.status === '불참' ? '불참' : e.status === '중도포기' ? '중도포기' : '걷는 중';
      if (st === '불참') absent++;
      if (st === '중도포기') drop++;
      people.push([classLabel(c), s.no, s.name, st, e ? e.time : '', e ? e.place : '', e ? e.reason : '', c, lab, e ? e.by : '']);
    });
    var go = list.length - absent;
    table.push([classLabel(c), list.length, absent, go, drop, go - drop]);
  });
  return { table: table, people: people, events: events };
}

function refresh(day) {
  var R = readRoster(day), S = computeStatus(R, readEvents(day));
  var ss = SpreadsheetApp.getActive(), name = '집계_' + day + '일차';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name, 1);
  sh.clear();
  var r = 1;
  sh.getRange(r, 1).setValue('꿈길 걷기 ' + DAYS[day].title + ' — 마지막 갱신 ' +
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm:ss')).setFontWeight('bold').setFontSize(13);
  r += 2;
  var head = ['반', '명단', '불참', '참가(출발)', '중도포기', '지금 걷는 인원'];
  var tot = S.table.reduce(function (a, x) { return a.map(function (v, i) { return i ? v + x[i] : v; }); }, ['합계', 0, 0, 0, 0, 0]);
  var rows = [head].concat(S.table).concat([tot]);
  sh.getRange(r, 1, rows.length, head.length).setValues(rows).setHorizontalAlignment('center');
  sh.getRange(r, 1, 1, head.length).setFontWeight('bold').setBackground('#E0F2FE');
  sh.getRange(r + rows.length - 1, 1, 1, head.length).setFontWeight('bold').setBackground('#F1F5F9');
  sh.getRange(r + 1, 5, rows.length - 1, 1).setFontColor('#B91C1C').setFontWeight('bold');
  sh.getRange(r + 1, 6, rows.length - 1, 1).setFontColor('#166534').setFontWeight('bold');
  r += rows.length + 2;

  // 중도포기·불참 학생 (나중에 명단 확인용)
  var out = S.people.filter(function (p) { return p[3] !== '걷는 중'; });
  sh.getRange(r, 1).setValue('빠진 학생 ' + out.length + '명').setFontWeight('bold').setFontSize(12);
  r++;
  var h2 = ['반', '번호', '이름', '상태', '시각', '어디쯤', '이유'];
  var body = [h2].concat(out.map(function (p) { return [p[0], p[1], p[2], p[3], p[4] ? fmt(p[4]) : '', p[5], p[6]]; }));
  sh.getRange(r, 1, body.length, h2.length).setValues(body);
  sh.getRange(r, 1, 1, h2.length).setFontWeight('bold').setBackground('#FEF3C7');
  r += body.length + 2;

  // 전체 기록 (시각 순)
  sh.getRange(r, 1).setValue('전체 기록 (시각 순)').setFontWeight('bold').setFontSize(12);
  r++;
  var h3 = ['시각', '반', '이름', '상황', '이유', '어디쯤', '입력한 선생님'];
  var log = [h3].concat(S.events.map(function (e) {
    return [fmt(e.time), classLabel(e.cls), e.names.join(', '), e.status, e.reason, e.place, e.by];
  }));
  sh.getRange(r, 1, log.length, h3.length).setValues(log);
  sh.getRange(r, 1, 1, h3.length).setFontWeight('bold').setBackground('#F1F5F9');
  sh.setFrozenRows(0);
  sh.setColumnWidths(1, 7, 100);   // 열 넓이 100 고정(자동 맞춤은 빈 열을 좁게 접어 버렸다)
}
function fmt(d) { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'HH:mm'); }

/* ══════════════════════════════════════════════════════════════
   📱 선생님 현황판 (웹앱)
   ★ 단톡방에 「몇 명 빠졌어요」를 계속 올리지 않아도, 아무 선생님이나 폰으로 지금 상태를 본다.
   ★ 기본 잠금 = 학교 계정 로그인(배포 때 「학교 도메인 내 모든 사용자」). 학생 이름은 구글 안에서만 보인다.
     꿈길 앱(공개 웹)에는 이 현황판으로 가는 «버튼»만 둔다.
   ★ 비번은 «선택» — 🔑 메뉴로 정하면 로그인에 더해 비번도 묻는다(코드엔 안 쓰고 스크립트 속성에 둔다).
     4자리라 막 눌러 맞히지 못하게 10분 안에 10번 틀리면 10분 동안 잠근다.
   배포(한 번만): Apps Script 오른쪽 위 「배포 → 새 배포 → ⚙ 웹 앱」
       실행 계정: 나 / 액세스 권한: 「(학교 도메인) 내 모든 사용자」 → 배포 → 웹 앱 URL 복사
   코드를 고친 뒤엔: 「배포 → 배포 관리 → ✏️ → 버전: 새 버전 → 배포」 (주소는 그대로)
   ══════════════════════════════════════════════════════════════ */
function doGet() {
  return HtmlService.createHtmlOutput(DASH_HTML)
    .setTitle('꿈길 걷기 현황판')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setPin() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('🔑 현황판 비번 (선택)', '선생님들이 같이 쓸 비밀번호(숫자 4자리 이상).\n비워 두고 확인하면 비번 없이 학교 계정 로그인만으로 열려요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var p = r.getResponseText().trim();
  if (!p) { PropertiesService.getScriptProperties().deleteProperty('pin'); ui.alert('비번을 껐어요. 학교 계정 로그인만으로 열려요.'); return; }
  if (p.length < 4) { ui.alert('4자리 이상으로 넣어 주세요.'); return; }
  PropertiesService.getScriptProperties().setProperty('pin', p);
  ui.alert('비번을 저장했어요. 현황판을 열면 이 비번을 물어요.');
}

function checkPin_(pin) {
  var real = PropertiesService.getScriptProperties().getProperty('pin');
  if (!real) return '';                                     // 비번을 안 정했으면 학교 계정 로그인만으로 연다
  var cache = CacheService.getScriptCache();
  if (cache.get('pinlock')) return '비번을 여러 번 틀려서 잠겼어요. 10분 뒤에 다시 해 주세요.';
  if (String(pin || '') === real) return '';
  if (!pin) return '';                                     // 처음 열 때(비번 칸 보여 주기)
  var n = Number(cache.get('pinfail') || 0) + 1;
  cache.put('pinfail', String(n), 600);
  if (n >= 10) { cache.put('pinlock', '1', 600); cache.remove('pinfail'); }
  return '비번이 맞지 않아요.';
}

/* 현황판이 1분마다 부른다 — 폼이 없어도 명단만 있으면 된다 */
function getStatus(day, pin) {
  var bad = checkPin_(pin);
  var real = PropertiesService.getScriptProperties().getProperty('pin');
  if (real && (bad || String(pin || '') !== real)) return { ok: false, need: 'pin', msg: bad };
  day = Number(day) === 2 ? 2 : 1;
  var R;
  try { R = readRoster(day); } catch (e) { return { ok: false, msg: e.message }; }
  var S = computeStatus(R, readEvents(day));
  var tot = { list: 0, absent: 0, go: 0, drop: 0, walk: 0 };
  var byCls = {};
  S.people.forEach(function (p) { (byCls[p[7]] = byCls[p[7]] || []).push({ name: p[8], st: p[3] }); });
  var classes = S.table.map(function (t, i) {
    tot.list += t[1]; tot.absent += t[2]; tot.go += t[3]; tot.drop += t[4]; tot.walk += t[5];
    var c = R.classes[i];
    return { cls: c, label: t[0], list: t[1], absent: t[2], go: t[3], drop: t[4], walk: t[5], students: byCls[c] || [] };
  });
  var out = S.people.filter(function (p) { return p[3] !== '걷는 중'; })
    .sort(function (a, b) { return (b[4] || 0) - (a[4] || 0); })
    .map(function (p) {
      return { cls: p[7], label: p[0], name: p[8], st: p[3], time: p[4] ? fmt(p[4]) : '', place: p[5], reason: p[6],
               by: String(p[9] || '').split('@')[0] };
    });
  return { ok: true, title: DAYS[day].title, time: Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm'),
           total: tot, classes: classes, out: out,
           reasons: REASONS.filter(function (r) { return r !== '결석·개인 사정'; }),
           places: DAYS[day].places.filter(function (p) { return p !== '출발 전'; }) };
}

/* 현황판에서 학생(여러 명 가능)의 상황을 기록한다 → 「기록_N일차」 탭에 한 줄 */
function record(day, pin, cls, names, status, reason, place) {
  names = [].concat(names).map(String).filter(Boolean);
  var real = PropertiesService.getScriptProperties().getProperty('pin');
  if (real && String(pin || '') !== real) return { ok: false, need: 'pin', msg: checkPin_(pin) || '비번을 다시 넣어 주세요.' };
  day = Number(day) === 2 ? 2 : 1;
  if (STATUS.indexOf(status) < 0) return { ok: false, msg: '알 수 없는 상황이에요.' };
  var R = readRoster(day), list = R.byClass[cls];
  if (!list) return { ok: false, msg: '명단에 없는 반이에요.' };
  var labels = list.map(function (s) { return choiceLabel(list, s); });
  var bad = names.filter(function (n) { return labels.indexOf(n) < 0; });
  if (!names.length || bad.length) return { ok: false, msg: '명단에 없는 학생이에요: ' + bad.join(', ') };
  var who = '';
  try { who = Session.getActiveUser().getEmail(); } catch (e) {}
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = logSheet(day, true), now = new Date();
    var rows = names.map(function (n) { return [now, cls, n, status, reason || '', place || '', who]; });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  try { refresh(day); } catch (e) {}
  return { ok: true, status: getStatus(day, pin) };
}

function showDashUrl() {
  var u = '';
  try { u = ScriptApp.getService().getUrl(); } catch (e) {}
  SpreadsheetApp.getUi().alert('📱 현황판 주소',
    u ? u + '\n\n이 주소를 꿈길 앱 담당 선생님께 보내 주세요. 앱의 「👩‍🏫 학생 확인」 버튼에 연결됩니다.'
      : '아직 배포하지 않았어요.\n\nApps Script 오른쪽 위 「배포 → 새 배포 → ⚙ 웹 앱」\n실행 계정: 나 / 액세스 권한: (학교 도메인) 내 모든 사용자 → 배포',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

var DASH_HTML = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box;margin:0}
body{font-family:-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;background:#F1F5FA;color:#0F172A;-webkit-text-size-adjust:100%}
button{font-family:inherit;cursor:pointer}
.hd{position:sticky;top:0;z-index:5;background:#0B4F82;color:#fff;padding:14px 16px 12px}
.hd h1{font-size:19px;font-weight:900}
.hd .t{font-size:13px;opacity:.85;margin-top:2px}
.seg{display:flex;gap:4px;background:rgba(255,255,255,.18);border-radius:12px;padding:3px;margin-top:10px}
.seg button{flex:1;border:0;background:none;color:#fff;font-size:15px;font-weight:800;padding:9px;border-radius:9px}
.seg button.on{background:#fff;color:#0B4F82}
.rf{border:0;background:rgba(255,255,255,.18);color:#fff;font-size:13px;font-weight:800;padding:5px 10px;border-radius:999px;float:right}
.wrap{padding:14px 14px 100px}
.big{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.big div{background:#fff;border-radius:16px;padding:12px 6px;text-align:center;box-shadow:0 2px 8px rgba(15,23,42,.06)}
.big b{display:block;font-size:30px;font-weight:900;line-height:1.15}
.big span{font-size:13px;font-weight:800;color:#64748B}
.big .walk b{color:#15803D}.big .drop b{color:#B91C1C}
.h{font-size:16px;font-weight:900;margin:20px 2px 8px;display:flex;justify-content:space-between;align-items:baseline}
.h small{font-size:13px;color:#0284C7;font-weight:800}
.cls{display:grid;grid-template-columns:repeat(2,1fr);gap:8px}
.c{border:0;text-align:left;background:#fff;border-radius:14px;padding:10px 12px;box-shadow:0 2px 8px rgba(15,23,42,.05);color:inherit}
.c .n{font-size:16px;font-weight:900;display:flex;justify-content:space-between;align-items:baseline}
.c .n em{font-style:normal;font-size:20px;color:#15803D}
.c .s{font-size:13px;color:#64748B;font-weight:700;margin-top:2px}
.c .s i{font-style:normal;color:#B91C1C}
.c.alert{outline:2px solid #FCA5A5}
.p{border:0;width:100%;text-align:left;background:#fff;border-radius:14px;padding:11px 13px;margin-bottom:7px;display:flex;gap:10px;align-items:center;box-shadow:0 2px 8px rgba(15,23,42,.05);color:inherit}
.tg{flex:0 0 auto;font-size:12px;font-weight:900;padding:4px 8px;border-radius:999px}
.tg.d{background:#FEE2E2;color:#B91C1C}.tg.a{background:#E2E8F0;color:#334155}
.p .nm{font-size:16px;font-weight:900}
.p .mt{font-size:13px;color:#64748B;font-weight:700;margin-top:1px}
.empty{background:#fff;border-radius:14px;padding:16px;text-align:center;color:#64748B;font-weight:700}
.err{background:#FEF3C7;color:#92400E;border-radius:14px;padding:14px;font-weight:700;line-height:1.6}
/* 반 화면 · 학생 화면 (아래에서 올라오는 판) */
.sh{position:fixed;inset:0;z-index:20;background:rgba(15,23,42,.45);display:none;align-items:flex-end}
.sh.show{display:flex}
.pn{width:100%;max-height:92vh;overflow:auto;background:#F8FAFC;border-radius:22px 22px 0 0;padding:16px 14px 26px}
.pn .top{display:flex;align-items:center;gap:10px;margin-bottom:12px}
.pn .top h2{font-size:20px;font-weight:900;flex:1}
.pn .top .x{border:0;background:#E2E8F0;width:40px;height:40px;border-radius:12px;font-size:22px}
.pn .sub{font-size:14px;color:#64748B;font-weight:700;margin:-6px 0 12px}
.names{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.nb{border:2px solid #E2E8F0;background:#fff;border-radius:14px;padding:12px 4px;font-size:16px;font-weight:800;color:#0F172A;line-height:1.2}
.nb small{display:block;font-size:11px;font-weight:800;margin-top:3px;color:#94A3B8}
.nb.absent{background:#F1F5F9;color:#94A3B8;border-color:#E2E8F0}
.nb.drop{background:#FEE2E2;border-color:#FCA5A5;color:#991B1B}
.nb.drop small{color:#B91C1C}
.nb.on{background:#DBEAFE;border-color:#2563EB;color:#1D4ED8;box-shadow:0 0 0 2px #2563EB inset}
.bar{display:none;position:sticky;bottom:-26px;margin:14px -14px -26px;padding:12px 14px 18px;background:#fff;border-top:1px solid #E2E8F0;box-shadow:0 -6px 16px rgba(15,23,42,.08)}
.bar .bt{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;font-size:16px}
.bar .clr{border:0;background:#F1F5F9;border-radius:999px;padding:6px 12px;font-weight:800;color:#475569}
.acts.three{grid-template-columns:1fr 1fr 1fr}
.q{font-size:15px;font-weight:900;margin:14px 2px 8px;color:#334155}
.acts{display:grid;gap:8px}
.acts.two{grid-template-columns:1fr 1fr}
.ab{border:0;border-radius:16px;padding:16px 10px;font-size:18px;font-weight:900;background:#fff;color:#0F172A;box-shadow:0 2px 8px rgba(15,23,42,.08)}
.ab.red{background:#DC2626;color:#fff}.ab.gray{background:#475569;color:#fff}.ab.green{background:#16A34A;color:#fff}
.ab.sm{font-size:16px;padding:14px 8px}
/* 아래 바로가기 — 꿈길 앱의 하단 탭과 같은 모양. 현황판은 구글 쪽 화면이라 «앱 주소로» 이동한다 */
.tabs{position:fixed;left:0;right:0;bottom:0;z-index:10;display:grid;grid-template-columns:repeat(5,1fr);background:#fff;border-top:1px solid #DBE4EF;box-shadow:0 -4px 14px rgba(15,23,42,.06);padding-bottom:env(safe-area-inset-bottom,0px)}
.tabs a{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;min-height:64px;color:#667A91;font-size:14px;font-weight:800;text-decoration:none}
.tabs a .e{font-size:22px;line-height:1}
.tabs a.on{color:#0284C7}
.toast{position:fixed;left:50%;bottom:84px;transform:translateX(-50%);background:#0F172A;color:#fff;font-weight:800;padding:12px 18px;border-radius:999px;z-index:40;display:none}
.pin{max-width:340px;margin:40px auto 0;background:#fff;border-radius:20px;padding:24px 20px;text-align:center;box-shadow:0 6px 20px rgba(15,23,42,.08)}
.pin h2{font-size:20px;font-weight:900;margin-bottom:6px}.pin p{font-size:14px;color:#64748B;font-weight:700;line-height:1.5}
.pin input{width:100%;margin-top:16px;font-size:30px;font-weight:900;letter-spacing:.4em;text-align:center;padding:12px;border:2px solid #CBD5E1;border-radius:14px}
.pin button{width:100%;margin-top:10px;font-size:18px;font-weight:900;padding:14px;border:0;border-radius:14px;background:#0B4F82;color:#fff}
.pin .m{margin-top:10px;color:#B91C1C;font-weight:800;font-size:14px;min-height:20px}
</style></head><body>
<div class="hd"><button class="rf" id="rf">↻ 새로고침</button><h1>🚶 꿈길 걷기 현황판</h1><div class="t" id="tm">불러오는 중…</div>
<div class="seg"><button data-d="1" class="on">1일차</button><button data-d="2">2일차</button></div></div>
<div class="wrap" id="w"></div>
<nav class="tabs">
  <a href="https://today119.github.io/kkumgil/walk/" target="_top"><span class="e">🏠</span>홈</a>
  <a href="https://today119.github.io/kkumgil/walk/?tab=course" target="_top"><span class="e">🗺</span>코스</a>
  <a href="https://today119.github.io/kkumgil/walk/?tab=prep" target="_top"><span class="e">🎒</span>준비물</a>
  <a href="https://today119.github.io/kkumgil/walk/?tab=stamp" target="_top"><span class="e">🏅</span>스탬프</a>
  <a class="on" href="#" onclick="load();return false"><span class="e">👩‍🏫</span>현황판</a>
</nav>
<div class="sh" id="sh"><div class="pn" id="pn"></div></div>
<div class="toast" id="toast"></div>
<script>
var day = 1, pin = '', S = null, openCls = null, busy = false;
try { pin = localStorage.getItem('kk_pin') || ''; } catch (e) {}
function $(id){ return document.getElementById(id); }
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function toast(m){ var t = $('toast'); t.textContent = m; t.style.display = 'block'; clearTimeout(toast.k); toast.k = setTimeout(function(){ t.style.display = 'none'; }, 2200); }

function load(){
  if(busy) return;
  $('tm').textContent = '불러오는 중…';
  google.script.run.withSuccessHandler(draw).withFailureHandler(function(e){
    $('w').innerHTML = '<div class="err">불러오지 못했어요: ' + esc(e.message) + '</div>';
  }).getStatus(day, pin);
}
function draw(R){
  if(R.need === 'pin'){ askPin(R.msg); return; }
  if(!R.ok){ $('w').innerHTML = '<div class="err">' + esc(R.msg) + '</div>'; $('tm').textContent = ''; return; }
  S = R;
  $('tm').textContent = S.title + ' · ' + S.time + ' 기준 · 1분마다 갱신';
  var T = S.total, h = '';
  h += '<div class="big"><div><b>' + T.go + '</b><span>참가</span></div>' +
       '<div class="drop"><b>' + T.drop + '</b><span>중도포기</span></div>' +
       '<div class="walk"><b>' + T.walk + '</b><span>지금 걷는 중</span></div></div>';
  h += '<div class="h">반별 <small>반을 누르면 기록할 수 있어요</small></div><div class="cls">' + S.classes.map(function(c, i){
    return '<button class="c' + (c.drop ? ' alert' : '') + '" data-ci="' + i + '"><div class="n">' + esc(c.label) + '<em>' + c.walk + '</em></div>' +
           '<div class="s">명단 ' + c.list + ' · 불참 ' + c.absent + ' · <i>포기 ' + c.drop + '</i></div></button>';
  }).join('') + '</div>';
  h += '<div class="h">빠진 학생 ' + S.out.length + '명</div>';
  h += S.out.length ? S.out.map(function(p){
    return '<button class="p" data-who="' + esc(p.cls) + '|' + esc(p.name) + '"><span class="tg ' + (p.st === '중도포기' ? 'd' : 'a') + '">' + esc(p.st) + '</span><div><div class="nm">' +
           esc(p.label) + ' ' + esc(p.name) + '</div><div class="mt">' + [p.time, p.place, p.reason, p.by].filter(Boolean).map(esc).join(' · ') + '</div></div></button>';
  }).join('') : '<div class="empty">아직 없어요 👍</div>';
  $('w').innerHTML = h;
  $('w').querySelectorAll('[data-ci]').forEach(function(b){ b.onclick = function(){ openClass(+b.dataset.ci); }; });
  $('w').querySelectorAll('[data-who]').forEach(function(b){
    b.onclick = function(){ var k = b.dataset.who.split('|'); openStudent(k[0], k[1]); };
  });
  if(openCls !== null && $('sh').classList.contains('show') && $('pn').dataset.mode === 'class') openClass(openCls, true);
}

/* ── 반 화면: 학생 이름 버튼 — 눌러서 «여러 명 선택» → 아래 줄에서 한꺼번에 처리 ── */
var sel = {};
function clsByKey(k){ for(var i = 0; i < S.classes.length; i++) if(S.classes[i].cls === k) return i; return -1; }
function selNames(){ return Object.keys(sel).filter(function(k){ return sel[k]; }); }
function who(c, names){ return esc(c.label) + ' ' + esc(names[0]) + (names.length > 1 ? ' 외 ' + (names.length - 1) + '명' : ''); }
function openClass(i, keep){
  var c = S.classes[i];
  if(openCls !== i || !keep) sel = {};
  openCls = i;
  var h = '<div class="top"><h2>' + esc(c.label) + '</h2><button class="x" data-x>×</button></div>' +
          '<div class="sub">걷는 중 ' + c.walk + ' · 불참 ' + c.absent + ' · 포기 ' + c.drop + ' — 학생을 골라 주세요(여러 명 가능)</div><div class="names">' +
          c.students.map(function(s){
            var k = s.st === '불참' ? 'absent' : s.st === '중도포기' ? 'drop' : '';
            return '<button class="nb ' + k + (sel[s.name] ? ' on' : '') + '" data-nm="' + esc(s.name) + '">' + (sel[s.name] ? '✓ ' : '') + esc(s.name) +
                   (k ? '<small>' + esc(s.st) + '</small>' : '') + '</button>';
          }).join('') + '</div><div class="bar" id="bar"></div>';
  show(h, 'class');
  $('pn').querySelectorAll('[data-nm]').forEach(function(b){
    b.onclick = function(){ sel[b.dataset.nm] = !sel[b.dataset.nm]; openClass(i, true); };
  });
  paintBar(c);
}
function paintBar(c){
  var n = selNames(), bar = $('bar'); if(!bar) return;
  if(!n.length){ bar.style.display = 'none'; return; }
  bar.style.display = 'block';
  bar.innerHTML = '<div class="bt"><b>' + n.length + '명 선택</b><button class="clr" id="clr">선택 해제</button></div>' +
    '<div class="acts three"><button class="ab red sm" data-st="중도포기">🚑 중도포기</button>' +
    '<button class="ab gray sm" data-st="불참">🚫 불참</button><button class="ab green sm" data-st="다시 합류">↩ 걷는 중</button></div>';
  $('clr').onclick = function(){ sel = {}; openClass(openCls, true); };
  bar.querySelectorAll('[data-st]').forEach(function(b){
    b.onclick = function(){
      var st = b.dataset.st, names = selNames();
      if(st === '중도포기') askReason(c, names);
      else save(c, names, st, '', st === '불참' ? '출발 전' : '');
    };
  });
}
/* 빠진 학생 목록에서 이름을 누르면 → 그 반 화면에서 그 학생이 골라진 채로 */
function openStudent(cls, name){
  var i = clsByKey(cls); if(i < 0) return;
  openCls = i; sel = {}; sel[name] = true;
  openClass(i, true);
}
function askReason(c, names){
  var h = '<div class="top"><h2>' + who(c, names) + '</h2><button class="x" data-back>‹</button></div>' +
          '<div class="q">🚑 중도포기 이유</div><div class="acts two">' +
          S.reasons.map(function(r){ return '<button class="ab sm" data-r="' + esc(r) + '">' + esc(r) + '</button>'; }).join('') + '</div>';
  show(h, 'student');
  $('pn').querySelector('[data-back]').onclick = function(){ openClass(openCls, true); };
  $('pn').querySelectorAll('[data-r]').forEach(function(b){ b.onclick = function(){ askPlace(c, names, b.dataset.r); }; });
}
function askPlace(c, names, reason){
  var h = '<div class="top"><h2>' + who(c, names) + '</h2><button class="x" data-back>‹</button></div>' +
          '<div class="sub">이유: ' + esc(reason) + '</div><div class="q">📍 어디쯤인가요?</div><div class="acts two">' +
          S.places.map(function(p){ return '<button class="ab sm" data-p="' + esc(p) + '">' + esc(p) + '</button>'; }).join('') + '</div>';
  show(h, 'student');
  $('pn').querySelector('[data-back]').onclick = function(){ askReason(c, names); };
  $('pn').querySelectorAll('[data-p]').forEach(function(b){ b.onclick = function(){ save(c, names, '중도포기', reason, b.dataset.p); }; });
}
function save(c, names, st, reason, place){
  busy = true;
  $('pn').innerHTML = '<div class="empty">저장하는 중…</div>';
  google.script.run.withSuccessHandler(function(R){
    busy = false;
    if(R && R.need === 'pin'){ hide(); askPin(R.msg); return; }
    if(!R || !R.ok){ toast((R && R.msg) || '저장하지 못했어요'); openClass(openCls, true); return; }
    toast('✅ ' + c.label + ' ' + names.length + '명 — ' + (st === '다시 합류' ? '걷는 중' : st));
    sel = {};
    draw(R.status);
    var i = clsByKey(c.cls); if(i >= 0) openClass(i, true);
  }).withFailureHandler(function(e){
    busy = false; toast('저장하지 못했어요: ' + e.message); openClass(openCls, true);
  }).record(day, pin, c.cls, names, st, reason, place);
}

function show(h, mode){ $('pn').innerHTML = h; $('pn').dataset.mode = mode; $('sh').classList.add('show');
  var x = $('pn').querySelector('[data-x]'); if(x) x.onclick = hide; }
function hide(){ $('sh').classList.remove('show'); openCls = null; }
$('sh').onclick = function(e){ if(e.target === $('sh')) hide(); };

/* 비번 — 확인은 구글 서버에서 한다. 맞으면 이 폰에 기억 */
function askPin(msg){
  $('tm').textContent = '';
  $('w').innerHTML = '<div class="pin"><h2>🔒 선생님 확인</h2><p>선생님들이 같이 쓰는 비밀번호를 넣어 주세요.</p>' +
    '<input id="pi" type="password" inputmode="numeric" autocomplete="off" maxlength="12"><button id="pb">확인</button><div class="m">' + esc(msg || '') + '</div></div>';
  var go = function(){ pin = $('pi').value.trim(); try { localStorage.setItem('kk_pin', pin); } catch (e) {} load(); };
  $('pb').onclick = go;
  $('pi').onkeydown = function(e){ if(e.key === 'Enter') go(); };
  $('pi').focus();
}
document.querySelectorAll('[data-d]').forEach(function(b){
  b.onclick = function(){ day = +b.dataset.d; hide(); document.querySelectorAll('[data-d]').forEach(function(x){ x.className = x === b ? 'on' : ''; }); load(); };
});
$('rf').onclick = load;
setInterval(function(){ if(!$('sh').classList.contains('show') || $('pn').dataset.mode === 'class') load(); }, 60000);
load();
</script></body></html>
`;
