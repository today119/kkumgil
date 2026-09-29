#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""앱 코스(course-dayN.json) → GPX.

  python3 tools/to_gpx.py walk/course-day1.json /tmp/1일차.gpx

gpx.studio 에서 「열기」로 불러 «이어 그리기» 할 때 쓴다.
답사 지점은 웨이포인트로 같이 넣어, 그리다가 어디가 갈림길이었는지 보이게 한다.
"""
import json, sys, html, datetime

KIND = {"fork": "갈림길", "spot": "명소", "wc": "화장실", "rest": "휴식",
        "risk": "주의", "car": "차량진입"}

def main(src, out):
    d = json.load(open(src, encoding="utf-8"))
    name = d.get("title", "코스")
    L = ['<?xml version="1.0" encoding="UTF-8"?>',
         '<gpx version="1.1" creator="kkumgil" xmlns="http://www.topografix.com/GPX/1/1">',
         '  <metadata><name>%s</name><time>%sZ</time></metadata>'
         % (html.escape(name), datetime.datetime.utcnow().replace(microsecond=0).isoformat())]
    for i, p in enumerate(d.get("points", []), 1):
        lab = "%s %s" % (KIND.get(p.get("type"), p.get("name", "")), 
                         ("(%s)" % p["dirName"]) if p.get("dirName") else "")
        L.append('  <wpt lat="%.6f" lon="%.6f"><name>%d. %s</name><desc>%.2fkm 지점</desc></wpt>'
                 % (p["lat"], p["lng"], p.get("no", i), html.escape(lab.strip()), p.get("m", 0)/1000))
    L.append('  <trk><name>%s</name><trkseg>' % html.escape(name))
    for a, b in d["course"]:
        L.append('    <trkpt lat="%.6f" lon="%.6f"></trkpt>' % (a, b))
    L += ['  </trkseg></trk>', '</gpx>']
    open(out, "w", encoding="utf-8").write("\n".join(L))
    print("%s → %s  (경로 %d점 · 지점 %d개 · %.2fkm)"
          % (src, out, len(d["course"]), len(d.get("points", [])), d["lengthM"]/1000))

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
