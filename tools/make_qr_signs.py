"""스탬프 안내판 QR — stamps.json 을 읽어 지점마다 A4 세로 한 장(PNG) + 묶음 PDF 를 만든다.

    python3 tools/make_qr_signs.py            # 전부
    python3 tools/make_qr_signs.py 2          # 2일차만
결과: ~/kkumgil/qr/<날>일차_<번호>_<이름>.png, ~/kkumgil/qr/스탬프_QR_<날>일차.pdf
QR 아래 4자리 = 카메라가 안 되는 폰은 앱 「QR 찍기」에서 숫자로 연다.
"""
import json, os, sys
import qrcode
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.expanduser("~/kkumgil/qr")
BASE = "https://today119.github.io/kkumgil/walk/?stamp={id}&k={code}"
FONT = "/System/Library/Fonts/AppleSDGothicNeo.ttc"
W, H = 2480, 3508  # A4 300dpi


def font(size, bold=True):
    return ImageFont.truetype(FONT, size, index=6 if bold else 0)


def center(d, y, text, f, fill):
    w = d.textlength(text, font=f)
    d.text(((W - w) / 2, y), text, font=f, fill=fill)


def sign(s):
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 520], fill="#1E3A8A")
    center(d, 110, "꿈길 걷기 스탬프 투어", font(120), "#BFDBFE")
    center(d, 270, f"{s['day']}일차 · {s['no']}번", font(150), "white")
    center(d, 640, s["name"], font(230), "#0F172A")
    q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=40, border=2)
    q.add_data(BASE.format(**s))
    q.make(fit=True)
    qi = q.make_image(fill_color="black", back_color="white").convert("RGB").resize((1600, 1600), Image.NEAREST)
    im.paste(qi, ((W - 1600) // 2, 1000))
    center(d, 2660, f"숫자 {s['code']}", font(200), "#1E3A8A")
    center(d, 2950, "꿈길 앱 → 스탬프 → QR 찍기", font(100, False), "#334155")
    center(d, 3110, "(카메라가 안 되면 숫자 4자리를 넣으세요)", font(80, False), "#64748B")
    return im


def main():
    day = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    data = json.load(open(os.path.join(ROOT, "walk", "stamps.json"), encoding="utf-8"))
    os.makedirs(OUT, exist_ok=True)
    by = {}
    for s in data["stamps"]:
        if day and s["day"] != day:
            continue
        im = sign(s)
        im.save(os.path.join(OUT, f"{s['day']}일차_{s['no']:02d}_{s['name']}.png"))
        by.setdefault(s["day"], []).append(im)
    for d, ims in by.items():
        p = os.path.join(OUT, f"스탬프_QR_{d}일차.pdf")
        ims[0].save(p, save_all=True, append_images=ims[1:], resolution=300)
        print(p, len(ims), "장")


if __name__ == "__main__":
    main()
