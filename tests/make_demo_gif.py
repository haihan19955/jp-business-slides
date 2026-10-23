"""README / note 用のデモ GIF：修正前 → jp_fix_wrap.py 実行 → 修正後"""
import os
import sys
from PIL import Image, ImageDraw, ImageFont

root = sys.argv[1]
ex, out = os.path.join(root, "examples"), os.path.join(root, "images")
FB = r"C:\Windows\Fonts\BIZ-UDGothicB.ttc"
FM = r"C:\Windows\Fonts\BIZ-UDGothicR.ttc"
W = 1000

def slide(folder, n):
    im = Image.open(os.path.join(ex, folder, f"slide-{n:02d}.png")).convert("RGB")
    return im.resize((W, int(im.height * W / im.width)), Image.LANCZOS)

def with_bar(im, text, color):
    bar_h = 56
    c = Image.new("RGB", (im.width, im.height + bar_h), "white")
    d = ImageDraw.Draw(c)
    d.rectangle((0, 0, im.width, bar_h), fill=color)
    d.text((20, bar_h // 2), text, font=ImageFont.truetype(FB, 26), fill="white", anchor="lm")
    c.paste(im, (0, bar_h))
    return c

def terminal(size, lines):
    c = Image.new("RGB", size, (24, 27, 34))
    d = ImageDraw.Draw(c)
    f = ImageFont.truetype(FM, 22)
    y = 40
    for color, t in lines:
        d.text((36, y), t, font=f, fill=color)
        y += 38
    return c

frames, durations = [], []
def add(im, ms):
    frames.append(im.convert("P", palette=Image.ADAPTIVE, colors=128))
    durations.append(ms)

for n in (9, 6):
    before = with_bar(slide("renders", n), "修正前：単語の途中で改行されている（泣き別れ）", (178, 34, 34))
    after = with_bar(slide("renders_fixed", n), "修正後：文節の切れ目で改行（自動）", (30, 110, 70))
    add(before, 2200)
    log = [((150, 220, 150), "$ python jp_fix_wrap.py 提案書.pptx -o 提案書_修正版.pptx"),
           ((200, 200, 200), "  PowerPoint で改行位置を取得…"),
           ((240, 200, 120), "  round 1: スライド9 「専用｜フォルダ」→「専用フォルダに」の前で改行" if n == 9 else "  round 1: スライド6 「例｜外処理」→「例外処理と」の前で改行"),
           ((240, 200, 120), "  round 1: スライド9 「API経｜由」→「API経由で」の前で改行" if n == 9 else "  round 1: スライド6 「860万｜円」→「年間約860万円の」の前で改行"),
           ((200, 200, 200), "  PowerPoint で再確認…"),
           ((150, 220, 150), "  修正 13 か所 → jp_lint.py: error 0 件")]
    add(terminal(before.size, log), 2400)
    for k in range(1, 6):
        add(Image.blend(before, after, k / 6), 90)
    add(after, 2800)

frames[0].save(os.path.join(out, "demo.gif"), save_all=True, append_images=frames[1:], duration=durations, loop=0, optimize=True)
print(os.path.getsize(os.path.join(out, "demo.gif")) // 1024, "KB")
