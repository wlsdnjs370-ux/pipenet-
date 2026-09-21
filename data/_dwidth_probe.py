# -*- coding: utf-8 -*-
"""PDF 에 적힌 선 굵기 값을 센다 — 검사에서 지시선/배관을 갈라 보려면 필요하다."""
import collections
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
out = ROOT / "data" / "_dcmp" / "w.pdf"
render_iso(b, out, link_item="Pipe velocity", node_item="Node pressure")

data = PdfReader(str(out)).pages[0].get_contents().get_data().decode("latin-1")
print(collections.Counter(round(float(v), 4)
                          for v in re.findall(r"([\d.]+)\s+w\b", data)).most_common())
