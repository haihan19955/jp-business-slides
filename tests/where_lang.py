"""en-US が残っている要素の種類と、その中のテキストを表示する（調査用）。"""
import re
import sys
import zipfile
from collections import Counter

z = zipfile.ZipFile(sys.argv[1])
c = Counter()
for name in sorted(z.namelist()):
    if not re.match(r"ppt/slides/slide\d+\.xml$", name):
        continue
    x = z.read(name).decode("utf-8")
    for m in re.finditer(r'<a:(rPr|endParaRPr|defRPr)[^>]*lang="en-US"[^>]*/?>(?:.*?<a:t>([^<]*)</a:t>)?', x):
        tag, text = m.group(1), (m.group(2) or "")
        in_tbl = "<a:tbl>" in x[: m.start()] and x.rfind("<a:tbl>", 0, m.start()) > x.rfind("</a:tbl>", 0, m.start())
        c[(tag, "table" if in_tbl else "shape", bool(re.search(r"[぀-鿿]", text)))] += 1
for k, v in c.most_common():
    print(k, v)
