# -*- coding: utf-8 -*-
"""우리 모듈 D 출력을 참조와 같은 조건(세로 A4)으로 한 장 뽑는다."""
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from core.d_display_model import load_display_model  # noqa: E402
from core.d_iso_renderer import render_iso  # noqa: E402
from core.d_result_binder import bind_results  # noqa: E402

base = ROOT / "routes" / "제출용[최종]"
bound = bind_results(load_display_model(base / "2. Pipenet_hand.sdf"),
                     base / "2. Pipenet_hand.xml")
out = ROOT / "data" / "_dform" / "ours.pdf"
out.parent.mkdir(parents=True, exist_ok=True)
drawn = render_iso(bound, out, link_item="Pipe velocity", node_item="Node pressure",
                   also_png=out.with_suffix(".png"))
print(f"{out}  방향={drawn.orientation}  배관 {drawn.pipes_drawn} · 라벨 {drawn.labels_drawn}")
