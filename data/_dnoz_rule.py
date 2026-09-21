# -*- coding: utf-8 -*-
"""헤드 방향 유도 규칙 후보를 수작업본(PIPENET 원본)의 실제 오프셋과 대조한다.

수작업 SDF 는 @/n 에 진짜 좌표가 있으므로 정답지가 된다. 자동 SDF 는 그 좌표가
입력노드와 같아 방향이 없다 — 유도 규칙이 정답지를 재현하는지 먼저 본다.
"""
from __future__ import annotations

import math
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.d_display_model import load_display_model

SUB = ROOT / "routes" / "제출용[최종]"


def incident_dirs(model, coords, base_label):
    """base 노드에 붙은 관로들이 base 에서 뻗어 나가는 단위벡터."""
    out = []
    for p in model.pipes:
        path = None
        if p.input_node == base_label:
            path = [coords[p.input_node], *p.waypoints, coords[p.output_node]]
        elif p.output_node == base_label:
            path = [coords[p.output_node], *reversed(p.waypoints), coords[p.input_node]]
        if path is None:
            continue
        for nxt in path[1:]:
            dx, dy = nxt[0] - path[0][0], nxt[1] - path[0][1]
            d = math.hypot(dx, dy)
            if d > 1e-9:
                out.append((dx / d, dy / d))
                break
    return out


for stem in ("2. Pipenet_hand", "2. Pipenet_auto"):
    model = load_display_model(SUB / f"{stem}.sdf")
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    xs = [n.x for n in model.nodes]
    ys = [n.y for n in model.nodes]
    span = max(max(xs) - min(xs), max(ys) - min(ys))
    print(f"\n=== {stem} === span {span:.1f}")

    deg = Counter()
    offsets = []
    agree_sum, agree_perp, total_true = 0, 0, 0
    for z in model.nozzles:
        base, tip = coords.get(z.input_node), coords.get(z.output_node)
        dirs = incident_dirs(model, coords, z.input_node)
        deg[len(dirs)] += 1
        if base is None or tip is None:
            continue
        dx, dy = tip[0] - base[0], tip[1] - base[1]
        dist = math.hypot(dx, dy)
        if dist <= 1e-9:
            continue
        total_true += 1
        offsets.append(dist)
        true_dir = (dx / dist, dy / dist)

        sx = sum(d[0] for d in dirs)
        sy = sum(d[1] for d in dirs)
        mag = math.hypot(sx, sy)
        if mag > 1e-6:
            cand = (-sx / mag, -sy / mag)
            if cand[0] * true_dir[0] + cand[1] * true_dir[1] > 0.9:
                agree_sum += 1
        if dirs:
            ux, uy = dirs[0]
            for cand in ((-uy, ux), (uy, -ux)):
                if cand[0] * true_dir[0] + cand[1] * true_dir[1] > 0.9:
                    agree_perp += 1
                    break

    print("  base 노드 차수 분포", dict(sorted(deg.items())))
    if offsets:
        print(f"  실제 오프셋 min {min(offsets):.1f} max {max(offsets):.1f}"
              f" 평균 {sum(offsets)/len(offsets):.1f}"
              f" (span 대비 {sum(offsets)/len(offsets)/span*100:.2f}%)")
    if total_true:
        print(f"  방향 일치 / 전체 {total_true}")
        print(f"    규칙A 입사벡터합의 반대  {agree_sum}")
        print(f"    규칙B 입사관로에 수직(둘 중 하나 맞으면) {agree_perp}")
