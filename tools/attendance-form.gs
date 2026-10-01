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
    .addToUi();
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
  SpreadsheetApp.getUi().alert('명단_1일차 / 명단_2일차 탭을 만들었어요.\n반 · 번호 · 이름 순서로 붙여넣으세요.\n(2일차는 학년이 섞이니 반을 「1-3」, 「2-5」처럼 적어 주세요)');
}

function readRoster(day) {
  var sh = SpreadsheetApp.getActive().getSheetByName('명단_' + day + '일차');
  if (!sh) throw new Error('명단_' + day + '일차 탭이 없어요. 「① 명단 탭 만들기」부터 해 주세요.');
  var rows = sh.getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return r[0] && r[2] && r[2].indexOf('예시') < 0; });
  if (!rows.length) throw new Error('명단_' + day + '일차 탭에 학생이 없어요.');
  var classes = [], byClass = {};
  rows.forEach(function (r) {
    var c = String(r[0]).trim(), nm = String(r[2]).trim();
    if (!byClass[c]) { byClass[c] = []; classes.push(c); }
    byClass[c].push({ no: Number(r[1]) || 0, name: nm });
  });
  classes.sort(classOrder);
  classes.forEach(function (c) { byClass[c].sort(function (a, b) { return a.no - b.no; }); });
  return { classes: classes, byClass: byClass };
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
  ui.alert(day + '일차 폼을 만들었어요.\n「안내」 탭에 주소와 QR 이 있어요.');
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
function refreshAll() { [1, 2].forEach(function (d) { if (PropertiesService.getDocumentProperties().getProperty('form' + d)) refresh(d); }); }

function readEvents(day) {
  var id = PropertiesService.getDocumentProperties().getProperty('form' + day);
  var form = FormApp.openById(id);
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
      people.push([classLabel(c), s.no, s.name, st, e ? e.time : '', e ? e.place : '', e ? e.reason : '']);
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
  sh.autoResizeColumns(1, 7);
}
function fmt(d) { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'HH:mm'); }
