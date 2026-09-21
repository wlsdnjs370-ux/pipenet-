# -*- coding: utf-8 -*-
"""무채색 칸이 PDF 에 어떤 연산자로 적히는지 본다."""
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from pypdf import PdfReader  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402
from core.d_iso_renderer import render_iso  # noqa: E402
from core.d_result_binder import bind_results  # noqa: E402

base = ROOT / "routes" / "제출용[최종]"
b = bind_results(load_display_model(base / "2. Pipenet_hand.sdf"), base / "2. Pipenet_hand.xml")
out = ROOT / "data" / "_dcmp" / "grey.pdf"
render_iso(b, out, link_item="Pipe velocity", node_item="Node pressure")

data = PdfReader(str(out)).pages[0].get_contents().get_data().decode("latin-1")
op = re.compile(r"(?P<rgb>(?:[\d.]+\s+){3})rg\b|(?P<g>[\d.]+)\s+g\b"
                r"|(?P<x>[\d.\-]+)\s+(?P<y>[\d.\-]+)\s+m\b")
colour = None
rows = []
for m in op.finditer(data):
    if m["rgb"]:
        colour = tuple(round(float(v), 4) for v in m["rgb"].split())
    elif m["g"]:
        colour = (round(float(m["g"]), 4),) * 3
    elif colour is not None and float(m["y"]) < 131.0:
        rows.append((round(float(m["y"]), 1), round(float(m["x"]), 1), colour))
for y, x, c in sorted(rows, key=lambda r: (-r[0], r[1])):
    print(f"y {y:7.1f}  x {x:7.1f}  {c}")
