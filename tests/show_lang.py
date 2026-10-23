"""pptx の全スライドで使われている lang 属性を数える（調査用）。"""
import re
import sys
import zipfile
from collections import Counter

for path in sys.argv[1:]:
    z = zipfile.ZipFile(path)
    c = Counter()
    for name in z.namelist():
        if re.match(r"ppt/slides/slide\d+\.xml$", name):
            c.update(re.findall(r'lang="([^"]+)"', z.read(name).decode("utf-8")))
    print(path, dict(c))
