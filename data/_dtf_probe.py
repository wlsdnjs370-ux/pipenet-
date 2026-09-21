# -*- coding: utf-8 -*-
"""우리 PDF 가 글자 크기를 어떤 모양으로 적는지 본다 — 검사에서 읽어내려면 필요하다."""
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
out = ROOT / "data" / "_dcmp" / "tf.pdf"
render_iso(b, out, link_item="Pipe velocity", node_item="Node pressure")

data = PdfReader(str(out)).pages[0].get_contents().get_data().decode("latin-1")
i = data.find("BT")
print(repr(data[i:i + 320]))
print("---- Tf 값 분포 ----")
import collections  # noqa: E402
print(collections.Counter(re.findall(r"/\S+\s+([\d.]+)\s+Tf", data)).most_common())
