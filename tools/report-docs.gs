/* ══════════════════════════════════════════════════════════════
   꿈길 걷기 — 학생별 탐방보고서 문서 만들기 (Google Apps Script)

   ★ 왜: 백지 보고서는 「재밌었다」로 끝난다. 자기 «탐방계획서»가 맨 앞에 채워져 있으면
          「계획 → 실제로 본 것 → 달라진 점 → 내 꿈」을 저절로 쓰게 된다. 사진도 문서에 바로 넣는다.
   ★ 결과물이 그대로 작품집 원고가 된다 — 「내가 고른 랜드마크」 칸으로 코스 순서 편집이 된다.

   쓰는 법 (학교 계정, 한 번만)
     1. 탐방계획서 «응답 시트»(구글 폼 응답이 쌓인 시트)를 연다 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 저장
     2. 시트 새로고침 → 메뉴 「📝 꿈길 보고서」 → ① 설정 → ② 미리 1명만 만들어 보기 → ③ 전체 만들기
     3. 「보고서 목록」 탭에 학생별 문서 링크가 생긴다(반별 폴더에도 정리됨)
   ══════════════════════════════════════════════════════════════ */

var R_TEST = ['이영종', '이길동'];          // 시험 응답 — 만들지 않는다
var R_PLACES = {
  '1': ['백운산·용궁사', '하늘도시', '영종진·구읍뱃터', '씨사이드파크·인천대교', '기타'],
  '2': ['인천의 대학', '개항장 일대', '인천의 오늘과 미래(송도)', '기타']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📝 꿈길 보고서')
    .addItem('① 설정', 'rSetup')
    .addItem('② 미리 1명만 만들어 보기', 'rPreviewOne')
    .addItem('③ 전체 만들기(이어서 만들기)', 'rMakeAll')
    .addSeparator()
    .addItem('❓ 사용법', 'rHelp')
    .addToUi();
}

function rHelp() {
  SpreadsheetApp.getUi().alert('📝 꿈길 보고서',
    '이 시트의 탐방계획서 응답으로 학생마다 「계획서가 채워진 보고서 문서」를 만들어요.\n\n' +
    '① 설정 — 학년, 결과를 모을 폴더 이름, 학생 계정 규칙(예: s26{학번}@yeongjong.icehs.kr)\n' +
    '② 미리 1명만 — 첫 학생 문서 하나만 만들어 모양을 확인\n' +
    '③ 전체 만들기 — 나머지를 만들어요. 6분 제한이 있어 중간에 멈추면 다시 누르면 이어서 만들어요.\n\n' +
    '계정 규칙을 넣으면 학생 계정에 «편집» 권한을 줘요. 비우면 문서만 만들고 공유는 안 해요.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function rProps() { return PropertiesService.getDocumentProperties(); }

function rSetup() {
  var ui = SpreadsheetApp.getUi(), P = rProps();
  var a = ui.prompt('① 학년', '이 응답 시트는 몇 학년인가요? (1 또는 2)', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var b = ui.prompt('② 폴더 이름', '학생 문서를 모을 드라이브 폴더 이름(없으면 만들어요)', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  var c = ui.prompt('③ 학생 계정 규칙 (선택)', '예: s26{학번}@yeongjong.icehs.kr (1·2학년 모두 지금 학번 앞에 s26)\n비워 두면 공유하지 않아요.', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  P.setProperty('grade', String(a.getResponseText()).trim() === '2' ? '2' : '1');
  P.setProperty('folder', b.getResponseText().trim() || '꿈길 보고서');
  P.setProperty('acct', c.getResponseText().trim());
  ui.alert('저장했어요. 「② 미리 1명만 만들어 보기」로 모양을 확인해 보세요.');
}

/* 응답 시트에서 학생 목록 — 같은 학번이 여러 번 냈으면 마지막 것 */
function rStudents() {
  var sh = SpreadsheetApp.getActive().getSheets()[0];
  var v = sh.getDataRange().getDisplayValues(), H = v[0].map(function (x) { return String(x).replace(/\s+/g, ''); });
  var col = function (re) { for (var i = 0; i < H.length; i++) if (re.test(H[i])) return i; return -1; };
  var cNo = col(/학번/), cNm = col(/^성명|^이름/), cJob = col(/희망진로/), cPlace = col(/특징/), cPlan = col(/연계|계획/), cGo = col(/참가희망/);
  var by = {};
  v.slice(1).forEach(function (r) {
    var no = String(r[cNo] || '').replace(/\.0$/, '').trim(), nm = String(r[cNm] || '').trim();
    if (!/^\d{5}$/.test(no) || !nm || R_TEST.indexOf(nm) >= 0) return;
    by[no] = { no: no, name: nm, job: r[cJob] || '', place: r[cPlace] || '', plan: r[cPlan] || '', go2: cGo >= 0 ? r[cGo] : '' };
  });
  return Object.keys(by).sort().map(function (k) { return by[k]; });
}

function rFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function rSub(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* 한 학생 문서 */
function rMakeDoc(s, P) {
  var grade = s.no.charAt(0), cls = String(Number(s.no.slice(1, 3))), num = String(Number(s.no.slice(3)));
  var title = grade + '-' + cls + ' ' + s.name + ' 꿈길 걷기 탐방보고서';
  var doc = DocumentApp.create(title), b = doc.getBody();
  b.setMarginTop(50).setMarginBottom(50).setMarginLeft(56).setMarginRight(56);
  var H = function (t) { return b.appendParagraph(t).setHeading(DocumentApp.ParagraphHeading.HEADING2); };
  var guide = function (t) { var p = b.appendParagraph(t); p.editAsText().setItalic(true).setForegroundColor('#64748B').setFontSize(10); return p; };
  var blank = function (n) { for (var i = 0; i < n; i++) b.appendParagraph(''); };

  b.appendParagraph('2026 제10회 꿈길 걷기 탐방보고서').setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph(grade + '학년 ' + cls + '반 ' + num + '번  ' + s.name + (s.job ? '   ·   희망 진로: ' + s.job : '')).editAsText().setBold(true);

  H('내가 고른 랜드마크');
  guide('하나만 남기고 나머지는 지우세요.');
  b.appendParagraph((R_PLACES[grade] || R_PLACES['1']).join('   /   '));

  H('① 가기 전에 — 나의 탐방계획서');
  guide('9월에 내가 쓴 계획서예요. 고치지 말고 그대로 두세요.');
  var t = b.appendTable([['장소 특징 조사', String(s.place || '(제출한 계획서가 없어요)')], ['진로 연계 계획', String(s.plan || '')]]);
  t.setBorderColor('#CBD5E1');
  for (var i = 0; i < t.getNumRows(); i++) {
    t.getRow(i).getCell(0).setWidth(110).setBackgroundColor('#F1F5F9').editAsText().setBold(true).setFontSize(10);
    t.getRow(i).getCell(1).editAsText().setFontSize(10);
  }

  H('② 다녀와서 — 실제로 보고 알게 된 것');
  guide('계획서에서 보려던 것을 실제로 봤나요? 무엇을, 어디서, 어떻게 봤는지 구체적으로.');
  blank(4);
  H('📷 사진 (2~4장)');
  guide('미션 사진·인증샷을 넣으세요. 폰 구글 문서 앱: ＋ → 이미지. 사진 아래에 한 줄 설명(어디서, 무엇).');
  var ph = b.appendTable([['사진 1\n\n\n\n설명:', '사진 2\n\n\n\n설명:']]);
  ph.setBorderColor('#CBD5E1');
  H('③ 계획과 달라진 점');
  guide('못 한 것, 뜻밖에 발견한 것, 생각이 바뀐 것.');
  blank(3);
  H('④ 내 꿈(진로)과 이어지는 점');
  blank(3);
  H('⑤ 한 줄 소감');
  guide('작품집에 실릴 수 있어요. 나만 쓸 수 있는 한 문장으로.');
  blank(1);
  doc.saveAndClose();

  var file = DriveApp.getFileById(doc.getId());
  var root = rFolder(P.getProperty('folder') || '꿈길 보고서');
  rSub(root, grade + '학년 ' + cls + '반').addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  var acct = P.getProperty('acct'), shared = '';
  if (acct) {
    var mail = acct.replace('{학번}', s.no);
    try { file.addEditor(mail); shared = mail; } catch (e) { shared = '공유 실패: ' + mail; }
  }
  return { url: doc.getUrl(), shared: shared, title: title };
}

function rIndex() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('보고서 목록');
  if (!sh) {
    sh = ss.insertSheet('보고서 목록');
    sh.appendRow(['학번', '이름', '문서', '공유', '만든 시각']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E0F2FE');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rDone(sh) {
  var v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues() : [];
  var o = {}; v.forEach(function (r) { o[r[0]] = 1; }); return o;
}

function rPreviewOne() {
  var P = rProps(), list = rStudents(), sh = rIndex(), done = rDone(sh);
  var s = list.filter(function (x) { return !done[x.no]; })[0];
  if (!s) { SpreadsheetApp.getUi().alert('만들 학생이 없어요.'); return; }
  var r = rMakeDoc(s, P);
  sh.appendRow([s.no, s.name, r.url, r.shared, new Date()]);
  SpreadsheetApp.getUi().alert('미리 만들었어요: ' + r.title + '\n「보고서 목록」 탭의 링크로 열어 보세요.\n괜찮으면 「③ 전체 만들기」를 누르세요.');
}

/* 6분 실행 제한 — 5분이 되면 멈추고, 다시 누르면 이어서 만든다 */
function rMakeAll() {
  var P = rProps(), list = rStudents(), sh = rIndex(), done = rDone(sh), t0 = Date.now(), n = 0;
  var todo = list.filter(function (x) { return !done[x.no]; });
  for (var i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > 5 * 60 * 1000) break;
    var r = rMakeDoc(todo[i], P);
    sh.appendRow([todo[i].no, todo[i].name, r.url, r.shared, new Date()]);
    n++;
  }
  var left = todo.length - n;
  SpreadsheetApp.getUi().alert(n + '명 만들었어요.' + (left ? '\n남은 ' + left + '명은 「③ 전체 만들기」를 한 번 더 누르면 이어서 만들어요.' : '\n모두 끝났어요.'));
}
