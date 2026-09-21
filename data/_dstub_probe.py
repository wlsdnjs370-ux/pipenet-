# -*- coding: utf-8 -*-
"""수작업본이 실제로 쓴 노즐 스텁 길이 — span 대비 비율로 환산해 본다."""
from __future__ import annotations

import math
import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from core.d_display_model import load_display_model  # noqa: E402

SUB = ROOT / "routes" / "제출용[최종]"

for stem in ("2. Pipenet_hand", "2. Pipenet_auto"):
    model = load_display_model(SUB / f"{stem}.sdf")
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    minx, miny, maxx, maxy = model.bounds()
    span = max(maxx - minx, maxy - miny)
    lengths = []
    for z in model.nozzles:
        a, b = coords.get(z.input_node), coords.get(z.output_node)
        if a is None or b is None or a == b:
            continue
        lengths.append(math.hypot(b[0] - a[0], b[1] - a[1]))
    print(f"[{stem}] span {span:.1f} · 방향 있는 노즐 {len(lengths)}/{len(model.nozzles)}")
    if lengths:
        med = statistics.median(lengths)
        print(f"    길이 최소 {min(lengths):.1f} 중앙 {med:.1f} 최대 {max(lengths):.1f}")
        print(f"    span 대비  최소 {min(lengths)/span:.5f} 중앙 {med/span:.5f}"
              f" 최대 {max(lengths)/span:.5f}")
