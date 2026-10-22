"""README 用の画像（修正前後の比較・スライド一覧）を作る。"""
import os
import sys
from PIL import Image, ImageDraw, ImageFont

root = sys.argv[1]
ex = os.path.join(root, "examples")
out = os.path.join(root, "images")
os.makedirs(out, exist_ok=True)
F = r"C:\Windows\Fonts\BIZ-UDGothicB.ttc"


def label(img, text, color):
    bar = Image.new("RGB", (img.width, 64), color)
    ImageDraw.Draw(bar).text((24, 32), text, font=ImageFont.truetype(F, 30), fill="white", anchor="lm")
    canvas = Image.new("RGB", (img.width, img.height + 64), "white")
    canvas.paste(bar, (0, 0))
    canvas.paste(img, (0, 64))
    return canvas


def before_after(slide, box, name, caption_before, caption_after):
    b = Image.open(os.path.join(ex, "renders", f"slide-{slide:02d}.png")).crop(box)
    a = Image.open(os.path.join(ex, "renders_fixed", f"slide-{slide:02d}.png")).crop(box)
    b = label(b, caption_before, (178, 34, 34))
    a = label(a, caption_after, (30, 110, 70))
    gap = 24
    im = Image.new("RGB", (b.width, b.height + a.height + gap), "white")
    im.paste(b, (0, 0))
    im.paste(a, (0, b.height + gap))
    im.save(os.path.join(out, name))


before_after(6, (60, 440, 1540, 560), "before_after_kpi.png",
             "修正前：「例｜外処理」「自動｜照合」「860万｜円」", "修正後：文節の切れ目で改行（jp_fix_wrap.py）")
before_after(9, (60, 300, 1540, 400), "before_after_process.png",
             "修正前：「専用｜フォルダ」「API経｜由」「項目の｜み」", "修正後：文節の切れ目で改行（jp_fix_wrap.py）")

# スライド一覧（2列×5行）
files = sorted(f for f in os.listdir(os.path.join(ex, "renders_fixed")) if f.endswith(".png"))
tw, th, pad = 760, 428, 16
grid = Image.new("RGB", (tw * 2 + pad * 3, (th + pad) * ((len(files) + 1) // 2) + pad), (230, 233, 238))
for i, f in enumerate(files):
    im = Image.open(os.path.join(ex, "renders_fixed", f)).resize((tw, th), Image.LANCZOS)
    grid.paste(im, (pad + (i % 2) * (tw + pad), pad + (i // 2) * (th + pad)))
grid.save(os.path.join(out, "gallery.png"))
print("ok")


