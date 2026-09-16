#!/usr/bin/env python3
"""답사 결과·GPX → 학생 앱 데이터(walk/course.json + walk/photos/).

  python3 tools/build_walk.py

코스는 SEGMENTS 순서대로 이어 붙인다. 구간은 두 종류다.
  {"survey": 가공본 json, "photos": 사진 폴더}  — 현장 답사 구간 (경로 + 지점 + 사진)
  {"gpx": gpx 파일}                            — 컴퓨터로 그린 구간 (경로만, <wpt>가 있으면 지점으로)

- 지점 km는 앞 구간 길이만큼 밀어서 「전체 코스 기준 m」로 바꾼다.
- 구간 사이가 GAP_WARN(m) 넘게 떨어져 있으면 경고한다. 틈이 있으면 그 사이에서 이탈 경보가 울린다.
- 사진은 긴 변 800px·JPEG 품질 70으로 줄인다(원본 평균 517KB → 약 100KB).
  학생 281명이 산에서 데이터로 받으므로 작을수록 좋다.
"""
import json
import math
import os
import unicodedata
import xml.etree.ElementTree as ET

from PIL import Image, ImageOps

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SEGMENTS = [
    {"survey": "data/survey-20260911/course_seg1.json", "photos": "data/survey-20260911/photos"},
    # 2026-09-16 사용자가 gpx.studio 로 직접 그린 구간(용궁사 앞~중산교차로~하늘초 앞, 2.78km)
    {"gpx": "data/gpx/gap_1to2_20260916_joined.gpx"},
    {"survey": "data/survey-20260915/course_seg2.json", "photos": "data/survey-20260915/photos"},
]
OUT = os.path.join(ROOT, "walk")
MAX_SIDE, QUALITY = 800, 70
GAP_WARN = 30
DP_TOL = 3  # GPX 경로 단순화 허용 오차(m)

# 지점 종류별 학생에게 보일 기본 이름 (메모가 있으면 메모를 쓴다)
KIND = {"fork": "갈림길", "spot": "명소", "wc": "화장실", "rest": "휴식", "risk": "주의", "car": "차량 진입"}
# 현장 메모 오타 — 원본은 그대로 두고 여기서만 고친다
TYPO = {"정산": "정상", "세터": "센터", "휴식장소": "휴식 장소"}


def fix(s):
    for a, b in TYPO.items():
        s = s.replace(a, b)
    return s


def nfc(s):
    return unicodedata.normalize("NFC", s)


def meters(a, b):
    kx = 111320 * math.cos(math.radians((a[0] + b[0]) / 2))
    return math.hypot((b[1] - a[1]) * kx, (b[0] - a[0]) * 111320)


def length(path):
    return sum(meters(path[i - 1], path[i]) for i in range(1, len(path)))


def simplify(path, tol):
    """Douglas-Peucker. 위경도를 평면(m)으로 펴서 계산한다."""
    if len(path) < 3:
        return path
    lat0 = path[0][0]
    kx = 111320 * math.cos(math.radians(lat0))
    xy = [((p[1] - path[0][1]) * kx, (p[0] - lat0) * 111320) for p in path]
    keep = [False] * len(path)
    keep[0] = keep[-1] = True
    stack = [(0, len(path) - 1)]
    while stack:
        i, j = stack.pop()
        (ax, ay), (bx, by) = xy[i], xy[j]
        L = math.hypot(bx - ax, by - ay) or 1e-9
        far, idx = 0, None
        for k in range(i + 1, j):
            d = abs((bx - ax) * (ay - xy[k][1]) - (ax - xy[k][0]) * (by - ay)) / L
            if d > far:
                far, idx = d, k
        if idx is not None and far > tol:
            keep[idx] = True
            stack += [(i, idx), (idx, j)]
    return [p for p, k in zip(path, keep) if k]


def snap_m(path, lat, lng):
    """지점을 경로에 내렸을 때 구간 시작부터의 거리(m)와 떨어진 거리(m)."""
    best, run = (0, float("inf")), 0.0
    for i in range(1, len(path)):
        a, b = path[i - 1], path[i]
        seg = meters(a, b)
        for t in (k / 10 for k in range(11)):
            q = (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
            d = meters(q, (lat, lng))
            if d < best[1]:
                best = (run + seg * t, d)
        run += seg
    return best


def shrink(src, dst):
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    im.thumbnail((MAX_SIDE, MAX_SIDE))
    im.save(dst, "JPEG", quality=QUALITY, optimize=True, progressive=True)


def load_survey(seg, no):
    data = json.load(open(os.path.join(ROOT, seg["survey"]), encoding="utf-8"))
    photo_dir = os.path.join(ROOT, seg["photos"])
    files = {nfc(f): f for f in os.listdir(photo_dir)}
    points = []
    for p in data["points"]:
        photos = []
        for name in p["photos"]:
            real = files.get(nfc(name))
            if not real:
                print("  사진 없음:", name)
                continue
            out_name = "s%d_%s" % (no, nfc(name))
            shrink(os.path.join(photo_dir, real), os.path.join(OUT, "photos", out_name))
            photos.append(out_name)
        points.append({
            "type": p["type"],
            "name": fix(p["memo"]) or KIND.get(p["type"], p["type"]),
            "m": p["km"] * 1000,
            "lat": p["lat"], "lng": p["lng"],
            "dir": p.get("dir", ""),
            "photos": photos,
        })
    return data["course"], data["lengthM"], points


def load_gpx(seg):
    root = ET.parse(os.path.join(ROOT, seg["gpx"])).getroot()
    local = lambda el: el.tag.split("}")[-1]
    path = [(float(e.get("lat")), float(e.get("lon"))) for e in root.iter() if local(e) in ("trkpt", "rtept")]
    path = [list(p) for p in simplify(path, DP_TOL)]
    points = []
    for w in (e for e in root.iter() if local(e) == "wpt"):
        kids = {local(k): (k.text or "").strip() for k in w}
        lat, lng = float(w.get("lat")), float(w.get("lon"))
        m, off = snap_m(path, lat, lng)
        if off > 50:
            print("  경로에서 %dm 떨어진 지점은 뺌: %s" % (off, kids.get("name")))
            continue
        typ = kids.get("type") if kids.get("type") in KIND else "spot"
        points.append({"type": typ, "name": kids.get("name") or KIND[typ], "m": m,
                       "lat": lat, "lng": lng, "dir": "", "photos": []})
    return path, length(path), points


def main():
    os.makedirs(os.path.join(OUT, "photos"), exist_ok=True)
    for f in os.listdir(os.path.join(OUT, "photos")):
        os.remove(os.path.join(OUT, "photos", f))
    course, points, offset_m = [], [], 0.0
    for no, seg in enumerate(SEGMENTS, 1):
        path, seg_len, seg_points = load_survey(seg, no) if "survey" in seg else load_gpx(seg)
        name = seg.get("survey") or seg.get("gpx")
        if course:
            gap = meters(course[-1], path[0])
            if gap > GAP_WARN:
                print("  ⚠️ %d구간 시작이 앞 구간 끝에서 %dm 떨어짐 — 틈을 직선으로 잇습니다 (%s)" % (no, gap, name))
            offset_m += gap
            if gap < 1:
                path = path[1:]
        course += path
        for p in seg_points:
            p["m"] = round(offset_m + p["m"])
        points += seg_points
        print("  %d구간 %.2fkm · 지점 %d개 (%s)" % (no, seg_len / 1000, len(seg_points), name))
        offset_m += seg_len

    for p in points:
        p["desc"] = ""  # 명소 해설: 출처 확인이 끝난 것만 채울 것
    data = {
        "title": "꿈길 걷기 1일차",
        "lengthM": round(offset_m),
        "course": course,
        "points": sorted(points, key=lambda x: x["m"]),
    }
    with open(os.path.join(OUT, "course.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    photos = os.listdir(os.path.join(OUT, "photos"))
    size = sum(os.path.getsize(os.path.join(OUT, "photos", x)) for x in photos)
    print("코스 %d점 · %.2fkm · 지점 %d개 · 사진 %d장 %.1fMB" % (
        len(course), offset_m / 1000, len(points), len(photos), size / 1e6))


if __name__ == "__main__":
    main()
