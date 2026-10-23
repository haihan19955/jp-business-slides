"""折り返しシミュレータを PowerPoint 実機の結果と比べる。
python tests/eval_wrap_sim.py tests/wrap_corpus.pptx tests/wrap_corpus.layout.json
"""
import json
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "skills", "japanese-business-slides", "scripts"))
from jp_wrap_sim import simulate_deck  # noqa: E402

pptx, truth_path = sys.argv[1], sys.argv[2]
truth = {(f["slide"], f["shape"]): f for f in json.load(open(truth_path, encoding="utf-8-sig"))}
sim, notes = simulate_deck(pptx)
ok = tot = line_ok = line_tot = 0
by = Counter(); byok = Counter()
fails = []
for f in sim:
    t = truth.get((f["slide"], f["shape"]))
    if not t or "Slide Number" in f["shape"]:
        continue
    for ps, pt in zip(f["paragraphs"], t["paragraphs"]):
        if not pt["text"]:
            continue
        tot += 1
        key = (ps["font"], ps["size"])
        by[ps["font"]] += 1
        strip = lambda ls: [l.rstrip(" ") for l in ls]
        same = strip(ps["lines"]) == strip(pt["lines"])
        ok += same
        byok[ps["font"]] += same
        # 改行位置（文字オフセット）の一致率
        def offs(ls):
            o, out = 0, set()
            for l in ls[:-1]:
                o += len(l); out.add(o)
            return out
        a, b = offs(ps["lines"]), offs(pt["lines"])
        line_tot += len(b) or 1
        line_ok += len(a & b) if b else (1 if not a else 0)
        if not same and len(fails) < 12:
            fails.append((ps["font"], ps["size"], f["boxW"], pt["lines"], ps["lines"]))
print(f"段落の完全一致: {ok}/{tot} = {ok / tot:.1%}")
print(f"改行位置の一致: {line_ok}/{line_tot} = {line_ok / line_tot:.1%}")
for k in by:
    print(f"  {k}: {byok[k]}/{by[k]} = {byok[k] / by[k]:.1%}")
for n in notes:
    print("note:", n)
for fnt, sz, w, t, s in fails:
    print(f"\n[{fnt} {sz}pt w={w}]\n  PPT: {t}\n  SIM: {s}")
