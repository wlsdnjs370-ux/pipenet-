# -*- coding: utf-8 -*-
"""헤드(노즐) 노드가 도면에 어떻게 놓이는지 실측 — 고치기 전 진단."""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.d_display_model import load_display_model
from core.d_result_binder import bind_results

SUB = ROOT / "routes" / "제출용[최종]"

for stem in ("2. Pipenet_auto", "2. Pipenet_hand"):
    sdf = SUB / f"{stem}.sdf"
    if not sdf.exists():
        print(stem, "— 없음")
        continue
    model = load_display_model(sdf)
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    virtual = {n.label for n in model.nodes if n.virtual}
    print(f"\n=== {stem} ===")
    print("노드", len(model.nodes), "(가상", len(virtual), ") 관로", len(model.pipes),
          "노즐", len(model.nozzles))

    same, missing_out, missing_in, ok = 0, [], [], 0
    kinds: Counter[str] = Counter()
    for z in model.nozzles:
        a = coords.get(z.input_node)
        b = coords.get(z.output_node)
        if a is None:
            missing_in.append(z.label)
            continue
        if b is None:
            missing_out.append(z.label)
            kinds["출력노드 좌표없음"] += 1
            continue
        if a == b:
            same += 1
            kinds["입력=출력 같은 자리"] += 1
        else:
            ok += 1
            kinds["정상"] += 1
    print("노즐 배치:", dict(kinds))
    print("  입력노드 좌표없음", len(missing_in), missing_in[:5])
    print("  출력노드 좌표없음", len(missing_out), missing_out[:5])

    # 출력노드가 가상노드인가?
    out_virtual = sum(1 for z in model.nozzles if z.output_node in virtual)
    in_virtual = sum(1 for z in model.nozzles if z.input_node in virtual)
    print("  출력노드가 가상", out_virtual, "/ 입력노드가 가상", in_virtual)

    # 노즐 출력노드가 real_nodes 에 들어가 노드 라벨/값 대상이 되는가
    real = {n.label for n in model.real_nodes}
    print("  출력노드가 real_nodes 에 포함", sum(1 for z in model.nozzles if z.output_node in real))
    print("  입력노드가 real_nodes 에 포함", sum(1 for z in model.nozzles if z.input_node in real))

    for z in model.nozzles[:6]:
        print(f"   {z.label:<12} in={z.input_node}{coords.get(z.input_node)}"
              f" out={z.output_node}{coords.get(z.output_node)}")

    xml = SUB / f"{stem}.xml"
    if xml.exists():
        bound = bind_results(model, xml)
        print("  결과결합 노즐", len(bound.nozzles), "노드", len(bound.nodes))
        print("  xml_only_nodes", len(bound.report.xml_only_nodes),
              bound.report.xml_only_nodes[:5])
