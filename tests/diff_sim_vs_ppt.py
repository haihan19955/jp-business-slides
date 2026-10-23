"""1つの pptx について、シミュレーションと PowerPoint 実機の改行の違いを表示する。
python tests/diff_sim_vs_ppt.py deck.pptx layout.json
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "skills", "japanese-business-slides", "scripts"))
from jp_wrap_sim import simulate_deck  # noqa: E402

truth = {(f["slide"], f["shape"]): f for f in json.load(open(sys.argv[2], encoding="utf-8-sig"))}
sim, _ = simulate_deck(sys.argv[1])
n = bad = 0
for f in sim:
    t = truth.get((f["slide"], f["shape"]))
    if t is None:
        print("missing in PowerPoint:", f["slide"], f["shape"])
        continue
    for ps, pt in zip(f["paragraphs"], t["paragraphs"]):
        if not pt["text"]:
            continue
        n += 1
        a = [l.rstrip() for l in ps["lines"]]
        b = [l.rstrip() for l in pt["lines"]]
        if a != b:
            bad += 1
            print(f"slide {f['slide']} [{f['shape']}] kind={f['kind']} w={f['boxW']}\n  PPT: {b}\n  SIM: {a}")
print(f"\n一致 {n - bad}/{n}")
