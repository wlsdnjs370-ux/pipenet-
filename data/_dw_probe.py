# -*- coding: utf-8 -*-
import collections
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402
from core.d_iso_renderer import render_iso  # noqa: E402
from core.d_result_binder import bind_results  # noqa: E402

base = ROOT / "routes" / "제출용[최종]"
b = bind_results(load_display_model(base / "2. Pipenet_auto.sdf"), base / "2. Pipenet_auto.xml")
out = ROOT / "data" / "_dw_probe.pdf"
render_iso(b, out, link_item="Pipe velocity")

doc = fitz.open(out)
c = collections.Counter()
for d in doc[0].get_drawings():
    c[(round(d["width"] or 0, 4),
       None if d["color"] is None else tuple(round(v, 3) for v in d["color"]),
       "".join(i[0] for i in d["items"])[:6],
       bool(d.get("dashes") and d["dashes"] != "[] 0"))] += 1
for k, v in c.most_common(15):
    print(v, k)
