# -*- coding: utf-8 -*-
"""assets/ の背景 PNG（文字なし・透かしなし・1920x1080）を PIL で生成する。
スキル制作側で1回実行して同梱する。Notebook 側で作り直す必要はない。"""
from pathlib import Path
from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent.parent / "assets"
OUT.mkdir(exist_ok=True)
W, H = 1920, 1080


def band_diagonal(name: str, base=(244, 246, 248), tone=(226, 231, 238)):
    im = Image.new("RGB", (W, H), base); d = ImageDraw.Draw(im)
    d.polygon([(W * 0.62, 0), (W, 0), (W, H), (W * 0.48, H)], fill=tone)
    im.save(OUT / name)


def dots(name: str, base=(250, 250, 247), tone=(214, 219, 226)):
    im = Image.new("RGB", (W, H), base); d = ImageDraw.Draw(im)
    for x in range(120, W, 60):
        for y in range(120, H, 60):
            d.ellipse([x - 3, y - 3, x + 3, y + 3], fill=tone)
    im.save(OUT / name)


def corner_arc(name: str, base=(248, 250, 252), tone=(228, 233, 240)):
    im = Image.new("RGB", (W, H), base); d = ImageDraw.Draw(im)
    d.ellipse([W * 0.55, H * 0.35, W * 1.25, H * 1.35], fill=tone)
    im.save(OUT / name)


if __name__ == "__main__":
    band_diagonal("bg_band.png"); dots("bg_dots.png"); corner_arc("bg_arc.png")
    print("生成:", sorted(p.name for p in OUT.glob("*.png")))
