# -*- coding: utf-8 -*-
"""PIPENET 원본 SDF 코퍼스 전량에서 노즐 스텁의 방향·길이 규칙을 실측한다.

목적: 자동 SDF 는 @/n 좌표가 입력노드와 겹쳐 방향이 없다. 무엇으로 대신할지
정하려면 원본이 실제로 어떻게 놓는지가 근거여야 한다.
"""
from __future__ import annotations

import math
import statistics
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.d_display_model import load_display_model

LIB = ROOT / "data" / "reference_library"
files = sorted(LIB.rglob("*.sdf"))
print("SDF", len(files))

deg = Counter()
verdict = Counter()
ratios: list[float] = []
degenerate_files = 0
total_nozzles = 0
files_with_nozzles = 0
failed = 0

for path in files:
    try:
        model = load_display_model(path)
    except Exception:
        failed += 1
        continue
    if not model.nozzles:
        continue
    files_with_nozzles += 1
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    xs = [n.x for n in model.nodes]
    ys = [n.y for n in model.nodes]
    if not xs:
        continue
    span = max(max(xs) - min(xs), max(ys) - min(ys)) or 1.0

    # base 노드에서 뻗어 나가는 첫 방향들
    incident: dict[str, list[tuple[float, float]]] = {}
    for p in model.pipes:
        for a, b, way in ((p.input_node, p.output_node, list(p.waypoints)),
                          (p.output_node, p.input_node, list(reversed(p.waypoints)))):
            if a not in coords or b not in coords:
                continue
            origin = coords[a]
            for nxt in [*way, coords[b]]:
                dx, dy = nxt[0] - origin[0], nxt[1] - origin[1]
                d = math.hypot(dx, dy)
                if d > 1e-9:
                    incident.setdefault(a, []).append((dx / d, dy / d))
                    break

    file_degenerate = True
    for z in model.nozzles:
        total_nozzles += 1
        base, tip = coords.get(z.input_node), coords.get(z.output_node)
        dirs = incident.get(z.input_node, [])
        deg[len(dirs)] += 1
        if base is None or tip is None:
            verdict["좌표없음"] += 1
            continue
        dx, dy = tip[0] - base[0], tip[1] - base[1]
        dist = math.hypot(dx, dy)
        if dist <= 1e-9:
            verdict["겹침(방향없음)"] += 1
            continue
        file_degenerate = False
        ratios.append(dist / span)
        if not dirs:
            verdict["입사관로없음"] += 1
            continue
        ux, uy = dirs[0]
        tx, ty = dx / dist, dy / dist
        dot = -(ux * tx + uy * ty)          # 입사방향의 반대와 얼마나 같은가
        cross = abs(-ux * ty + uy * tx)     # 수직 성분
        if dot > 0.94:
            verdict["연장선(관로 반대방향)"] += 1
        elif cross > 0.94:
            verdict["수직"] += 1
        elif dot < -0.94:
            verdict["관로쪽으로 되꺾임"] += 1
        else:
            verdict["기타 각도"] += 1
    if file_degenerate:
        degenerate_files += 1

print("노즐 있는 파일", files_with_nozzles, "/ 파싱실패", failed)
print("노즐 총", total_nozzles)
print("base 노드 차수", dict(sorted(deg.items())))
print("방향 판정", dict(verdict.most_common()))
if ratios:
    print(f"스텁 길이 / 도면 span : 중앙값 {statistics.median(ratios)*100:.3f}%"
          f" 평균 {statistics.mean(ratios)*100:.3f}%"
          f" 5~95% {sorted(ratios)[len(ratios)//20]*100:.3f}~"
          f"{sorted(ratios)[len(ratios)*19//20]*100:.3f}%")
print("전량 겹침인 파일", degenerate_files)
