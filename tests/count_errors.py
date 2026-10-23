"""jp_lint のチェック結果を種類ごとに数える。python tests/count_errors.py deck.pptx layout.json"""
import json
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "skills", "japanese-business-slides", "scripts"))
from jp_lint import Report, check_deck, check_layout  # noqa: E402
from pptx import Presentation  # noqa: E402

rep = Report()
check_deck(Presentation(sys.argv[1]), rep)
with open(sys.argv[2], encoding="utf-8-sig") as fh:
    check_layout(json.load(fh), rep)
print("errors", rep.errors, dict(Counter(i["code"] for i in rep.items if i["level"] == "error")))
