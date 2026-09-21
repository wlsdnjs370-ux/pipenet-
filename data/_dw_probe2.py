# -*- coding: utf-8 -*-
import collections
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT)]

from pypdf import PdfReader  # noqa: E402

OP = re.compile(
    r"(?P<grey>[\d.]+) G\b|(?P<rgb>(?:[\d.]+ ){3})RG\b|(?P<stack>(?<![\w.])[qQ](?![\w.]))"
    r"|(?P<x>[\d.\-]+) (?P<y>[\d.\-]+) [mlc]\b")

data = PdfReader(str(ROOT / "data" / "_dw_probe.pdf")).pages[0].get_contents().get_data().decode("latin-1")
out, colour, started, cur = [], None, None, []
for m in OP.finditer(data):
    if m["x"] is None:
        if m["grey"]:
            colour = float(m["grey"])
        elif m["rgb"]:
            colour = tuple(round(float(v), 4) for v in m["rgb"].split())
        continue
    p = (float(m["x"]), float(m["y"]))
    if data[m.end() - 1] == "m":
        if cur:
            out.append((started, cur))
        cur, started = [p], colour
    else:
        cur.append(p)
if cur:
    out.append((started, cur))

c = collections.Counter((str(k), len(v)) for k, v in out)
for k, v in c.most_common(12):
    print(v, k)
print("---- 스트림 조각 ----")
i = data.rfind("/M0 Do")
print(repr(data[i:i + 1200]))
