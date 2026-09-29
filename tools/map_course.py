#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""코스를 OpenStreetMap 지도 위에 그려 PNG 로 저장한다.

  python3 tools/map_course.py walk/course-day1.json /tmp/day1.png "1일차 코스"

⚠️ OSM 타일 정책 — 대량으로 긁지 말고, 실제 User-Agent 를 보낼 것.
   한 번에 수십 장 수준이라 문제 없지만 자동 반복 실행은 하지 않는다.
"""
import io, json, math, os, sys, time, urllib.request
from PIL import Image, ImageDraw, ImageFont

TILE, UA = 256, "kkumgil-course-map/1.0 (school walking event; contact: teacher51@yeongjong.icehs.kr)"
FONT = "/System/Library/Fonts/AppleSDGothicNeo.ttc"


def deg2num(lat, lng, z):
    la = math.radians(lat); n = 2.0 ** z
    return ((lng + 180.0) / 360.0 * n,
            (1.0 - math.asinh(math.tan(la)) / math.pi) / 2.0 * n)


def fetch(z, x, y, cache="/tmp/osmtiles"):
    os.makedirs(cache, exist_ok=True)
    p = os.path.join(cache, "%d_%d_%d.png" % (z, x, y))
    if os.path.exists(p):
        return Image.open(p).convert("RGB")
    req = urllib.request.Request("https://tile.openstreetmap.org/%d/%d/%d.png" % (z, x, y),
                                 headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = r.read()
    open(p, "wb").write(data)
    time.sleep(0.12)                       # 서버를 몰아치지 않는다
    return Image.open(io.BytesIO(data)).convert("RGB")


def render(course, out, title, points=None, zoom=None, pad=0.06):
    lats = [p[0] for p in course]; lngs = [p[1] for p in course]
    la0, la1, ln0, ln1 = min(lats), max(lats), min(lngs), max(lngs)
    dla, dln = (la1 - la0) or 1e-4, (ln1 - ln0) or 1e-4
    la0 -= dla * pad; la1 += dla * pad; ln0 -= dln * pad; ln1 += dln * pad

    if zoom is None:                       # 2000px 안에 들어오는 가장 큰 배율
        for z in range(17, 9, -1):
            x0, y1 = deg2num(la0, ln0, z); x1, y0 = deg2num(la1, ln1, z)
            if (x1 - x0) * TILE <= 2100 and (y1 - y0) * TILE <= 2100:
                zoom = z; break
        else:
            zoom = 12
    z = zoom
    x0, y1 = deg2num(la0, ln0, z); x1, y0 = deg2num(la1, ln1, z)
    tx0, tx1, ty0, ty1 = int(x0), int(x1), int(y0), int(y1)
    W = (tx1 - tx0 + 1) * TILE; H = (ty1 - ty0 + 1) * TILE
    img = Image.new("RGB", (W, H), "#eee")
    n = 0
    for tx in range(tx0, tx1 + 1):
        for ty in range(ty0, ty1 + 1):
            try:
                img.paste(fetch(z, tx, ty), ((tx - tx0) * TILE, (ty - ty0) * TILE)); n += 1
            except Exception as e:
                print("  타일 실패", tx, ty, e)
    print("  타일 %d장 · 배율 z%d · %dx%d" % (n, z, W, H))

    def xy(lat, lng):
        px, py = deg2num(lat, lng, z)
        return ((px - tx0) * TILE, (py - ty0) * TILE)

    d = ImageDraw.Draw(img, "RGBA")
    path = [xy(*p) for p in course]
    d.line(path, fill=(255, 255, 255, 220), width=11, joint="curve")
    d.line(path, fill=(0, 110, 255, 235), width=6, joint="curve")

    f = lambda s: ImageFont.truetype(FONT, s) if os.path.exists(FONT) else ImageFont.load_default()
    for p, col, lab in ((path[0], (16, 160, 70), "출발"), (path[-1], (220, 40, 40), "끝")):
        d.ellipse([p[0]-13, p[1]-13, p[0]+13, p[1]+13], fill=col, outline="white", width=4)
        d.text((p[0]+18, p[1]-11), lab, fill=col, font=f(30), stroke_width=5, stroke_fill="white")

    for q in (points or []):
        p = xy(q["lat"], q["lng"])
        d.ellipse([p[0]-7, p[1]-7, p[0]+7, p[1]+7], fill=(255, 170, 0), outline="white", width=3)

    if title:
        t = f(34)
        w = d.textlength(title, font=t)
        d.rectangle([0, 0, w + 34, 56], fill=(255, 255, 255, 235))
        d.text((17, 11), title, fill=(20, 20, 20), font=t)
    cr = f(20)
    c = "지도 © OpenStreetMap 기여자"
    d.rectangle([W - d.textlength(c, font=cr) - 22, H - 34, W, H], fill=(255, 255, 255, 220))
    d.text((W - d.textlength(c, font=cr) - 11, H - 29), c, fill=(60, 60, 60), font=cr)
    img.save(out)
    print("  저장:", out)


if __name__ == "__main__":
    src = sys.argv[1]; out = sys.argv[2]
    d = json.load(open(src, encoding="utf-8"))
    render(d["course"], out, sys.argv[3] if len(sys.argv) > 3 else d.get("title", ""), d.get("points"))
