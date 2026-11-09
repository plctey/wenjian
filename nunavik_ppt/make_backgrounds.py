"""Generate aurora-themed background images used by build_ppt.py."""
import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

W, H = 1920, 1080
OUT = Path(__file__).parent / "assets"
OUT.mkdir(exist_ok=True)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def sky(top, bottom):
    img = Image.new("RGB", (W, H))
    d = ImageDraw.Draw(img)
    for y in range(H):
        d.line([(0, y), (W, y)], fill=lerp(top, bottom, y / H))
    return img


def stars(img, n, seed, max_alpha=255):
    rnd = random.Random(seed)
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    for _ in range(n):
        x, y = rnd.randrange(W), rnd.randrange(int(H * 0.75))
        r = rnd.choice([1, 1, 1, 2])
        a = rnd.randint(80, max_alpha)
        d.ellipse([x - r, y - r, x + r, y + r], fill=(255, 255, 255, a))
    img.alpha_composite(layer)


def aurora(img, curtains, intensity=1.0, blur=28):
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    for c in curtains:
        x0, x1 = c["x"]
        for x in range(x0, x1, 2):
            t = (x - x0) / max(1, (x1 - x0))
            fade = math.sin(math.pi * t) ** 0.8
            cy = c["y"] + c["amp"] * math.sin(x / c["wl"] + c["ph"]) + 0.5 * c["amp"] * math.sin(x / (c["wl"] * 0.37) + 1.3)
            height = c["h"] * (0.6 + 0.4 * math.sin(x / (c["wl"] * 0.21) + c["ph"]) ** 2)
            for k in range(int(height)):
                f = k / height
                col = lerp(c["c0"], c["c1"], f)
                a = int(255 * intensity * fade * (1 - f) ** 1.6 * c["a"])
                y = int(cy - k)
                if 0 <= y < H:
                    d.point((x, y), fill=col + (a,))
                    d.point((x + 1, y), fill=col + (a,))
    layer = layer.filter(ImageFilter.GaussianBlur(blur))
    img.alpha_composite(layer)
    img.alpha_composite(layer.filter(ImageFilter.GaussianBlur(4)))


def mountains(img, base, color, snow, seed, peaks=7, alpha=255):
    rnd = random.Random(seed)
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    pts = [(0, H)]
    xs = sorted([0, W] + [rnd.randint(0, W) for _ in range(peaks * 2)])
    tops = []
    for i, x in enumerate(xs):
        y = base - (rnd.randint(60, 220) if i % 2 == 0 else rnd.randint(0, 60))
        pts.append((x, y))
        tops.append((x, y))
    pts.append((W, H))
    d.polygon(pts, fill=color + (alpha,))
    if snow:
        for i in range(1, len(tops) - 1):
            x, y = tops[i]
            if y < base - 90:
                lx, ly = tops[i - 1]
                rx, ry = tops[i + 1]
                s = 0.22
                d.polygon(
                    [(x, y), (x + (lx - x) * s, y + (ly - y) * s), (x + (rx - x) * s, y + (ry - y) * s)],
                    fill=snow + (alpha,),
                )
    img.alpha_composite(layer)


def cover():
    img = sky((4, 10, 26), (16, 44, 74)).convert("RGBA")
    stars(img, 420, 1)
    aurora(img, [
        dict(x=(-100, 1500), y=520, amp=70, wl=180, ph=0.2, h=420, c0=(61, 220, 151), c1=(142, 124, 195), a=0.9),
        dict(x=(500, 2100), y=430, amp=90, wl=230, ph=1.7, h=360, c0=(46, 196, 182), c1=(120, 90, 200), a=0.75),
        dict(x=(1100, 2000), y=600, amp=50, wl=140, ph=2.5, h=260, c0=(110, 240, 170), c1=(60, 140, 200), a=0.6),
    ])
    mountains(img, 900, (14, 32, 54), (120, 150, 180), seed=3, peaks=6)
    mountains(img, 990, (8, 20, 36), (70, 95, 125), seed=8, peaks=8)
    img.convert("RGB").save(OUT / "cover.jpg", quality=92)


def section():
    img = sky((5, 12, 30), (12, 34, 60)).convert("RGBA")
    stars(img, 300, 5)
    aurora(img, [
        dict(x=(600, 2100), y=560, amp=80, wl=200, ph=0.9, h=440, c0=(61, 220, 151), c1=(142, 124, 195), a=0.85),
        dict(x=(1000, 2000), y=460, amp=60, wl=160, ph=2.1, h=300, c0=(46, 196, 182), c1=(100, 80, 190), a=0.6),
    ])
    mountains(img, 960, (10, 24, 42), (80, 105, 135), seed=11, peaks=7)
    img.convert("RGB").save(OUT / "section.jpg", quality=92)


def content():
    img = sky((9, 22, 42), (11, 29, 51)).convert("RGBA")
    stars(img, 140, 9, max_alpha=150)
    aurora(img, [
        dict(x=(1150, 2100), y=230, amp=40, wl=150, ph=0.4, h=260, c0=(61, 220, 151), c1=(142, 124, 195), a=0.35),
    ], blur=40)
    mountains(img, 1060, (13, 33, 57), None, seed=21, peaks=9, alpha=255)
    img.convert("RGB").save(OUT / "content.jpg", quality=92)


if __name__ == "__main__":
    cover()
    section()
    content()
    print("backgrounds written to", OUT)
