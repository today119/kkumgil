#!/usr/bin/env python3
"""attendance-form.gs 의 현황판 화면(DASH_HTML) → walk/dash.html (꿈길 앱 안의 현황판).

  python3 tools/build_dash.py

★ 왜 따로 두나(10/28 선생님 폰에서 「현재 파일을 열 수 없습니다」):
  폰 브라우저에 구글 계정이 여러 개 로그인돼 있으면 Apps Script 웹앱 화면 자체가 안 열린다.
  앱 주소(github.io)에서 같은 화면을 띄우고, 자료만 doPost(act: status / record)로 주고받으면
  구글 로그인 쿠키가 아예 안 실려서 누구 폰이든 열린다.
★ 화면은 DASH_HTML 그대로 쓴다 — google.script.run 자리에 fetch 로 바꿔 끼우는 심(shim)만 넣는다.
  현황판 화면을 고치면 이 스크립트를 다시 돌린다.
"""
import json
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parent.parent
gs = (ROOT / 'tools/attendance-form.gs').read_text()
h = gs[gs.index('var DASH_HTML = `') + len('var DASH_HTML = `'):]
h = h[:h.index('`;')]

url = json.loads((ROOT / 'walk/courses.json').read_text())['teacher']['url']
EXEC = re.sub(r'/a/macros/[^/]+/', '/macros/', url)        # 학교 로그인 없이 받는 주소

SHIM = """<script>
/* google.script.run 대신 — 같은 이름으로 불러서 doPost 로 보낸다(구글 로그인 쿠키 없이) */
var EXEC = %s;
function gsr_(){
  var ok = function(){}, fail = function(){}, o = {
    withSuccessHandler: function(f){ ok = f; return o; },
    withFailureHandler: function(f){ fail = f; return o; },
    getStatus: function(day, pin){ call({act: 'status', day: day, pin: pin}); },
    record: function(day, pin, cls, names, st, reason, place, me){
      call({act: 'record', day: day, pin: pin, cls: cls, names: names, st: st, reason: reason, place: place, me: me}); }
  };
  function call(b){
    fetch(EXEC, {method: 'POST', body: JSON.stringify(b)})
      .then(function(r){ return r.json(); })
      .then(ok, function(){ fail({message: '인터넷 연결을 확인해 주세요'}); });
  }
  return o;
}
var google = {script: {}};
Object.defineProperty(google.script, 'run', {get: gsr_});
</script>""" % json.dumps(EXEC)

h = h.replace('<script>', SHIM + '\n<script>', 1)
h = h.replace('</head>', '<title>꿈길 걷기 현황판</title>\n<link rel="manifest" href="manifest.json">\n</head>', 1)
h = h.replace('https://today119.github.io/kkumgil/walk/', './')
h = h.replace('target="_top"', '')
(ROOT / 'walk/dash.html').write_text(h)
print('walk/dash.html', len(h), 'chars →', EXEC[:60] + '…')
