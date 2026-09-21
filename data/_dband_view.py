# -*- coding: utf-8 -*-
"""판면을 바꾼 뒤 실제로 어떻게 보이는지 한 장 뽑는다."""
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402
from core.d_iso_renderer import render_iso  # noqa: E402
from core.d_result_binder import bind_results  # noqa: E402

base = ROOT / "routes" / "제출용[최종]"
b = bind_results(load_display_model(base / "2. Pipenet_hand.sdf"), base / "2. Pipenet_hand.xml")
out = ROOT / "data" / "_dcmp" / "bands.pdf"
render_iso(b, out, link_item="Pipe volumetric flow", node_item="Node pressure")
doc = fitz.open(out)
doc[0].get_pixmap(dpi=200).save(ROOT / "data" / "_dcmp" / "bands.png")
doc.close()
print("됐다")
