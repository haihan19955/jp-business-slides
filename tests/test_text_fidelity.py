#!/usr/bin/env python3
"""jp_fix_wrap.py が文字を変えないこと、変わったら検出できることを確かめる（--simulate で動く）

  python tests/test_text_fidelity.py
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "skills", "japanese-business-slides", "scripts")
sys.path.insert(0, SCRIPTS)

from pptx import Presentation  # noqa: E402
from jp_fix_wrap import A, deck_text, text_diffs  # noqa: E402

DECKS = ["examples/sample.pptx", "tests/messy.pptx"]


def main():
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    with tempfile.TemporaryDirectory() as tmp:
        for deck in DECKS:
            src = os.path.join(ROOT, deck)
            if not os.path.exists(src):  # tests/*.pptx は make_messy_deck.py で作る（git には入れていない）
                print(f"skip {deck}")
                continue
            out = os.path.join(tmp, "fixed.pptx")
            subprocess.run([sys.executable, os.path.join(SCRIPTS, "jp_fix_wrap.py"), src, "-o", out, "--simulate"],
                           check=True, capture_output=True, env=env)
            diffs = text_diffs(deck_text(src), deck_text(out))
            assert not diffs, f"{deck}: {diffs[:3]}"
            print(f"ok  {deck}: {sum(len(s) for s in deck_text(src))} 段落の文字が同じ")

            # 1字消した資料は検出できること
            prs = Presentation(out)
            t = next(t for s in prs.slides for t in s._element.iter(A + "t") if t.text and len(t.text) > 1)
            t.text = t.text[:-1]
            broken = os.path.join(tmp, "broken.pptx")
            prs.save(broken)
            assert text_diffs(deck_text(src), deck_text(broken)), f"{deck}: 1字の欠落を検出できなかった"
            print(f"ok  {deck}: 1字の欠落を検出できる")


if __name__ == "__main__":
    main()
