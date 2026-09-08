# -*- coding: utf-8 -*-
"""[계통도 꼬임] 자동 레이어 «섞기» 가 경로에 무엇을 하는지 잰다.

사용자 지적: 「계통도 경로 추출이 그 레이어 안에서 색깔별로 최소 길이가
나와야 되는데, 뭔가 경로가 꼬여서 추적이 된다」.

앞선 측정: 자동 필터가 HSP·LSP·감압밸브를 **한 그래프에 섞는다**. 여기서는
①섞은 것 ②레이어 하나씩 을 나란히 세워, 절점·조각·추측다리가 어떻게
달라지는지 숫자로 남긴다. 기계실도 같은 잣대로 잰다.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_D = os.path.join(_ROOT, "routes", "제출용[최종]")


def measure(entities, layers):
    """★`components_after_bridge` 를 «다리 전» 으로 읽지 않는다.

    force_connect=False 라도 엔진은 허용오차 다리(200mm~10m)를 먼저 놓는다.
    끊긴 정도를 보려면 `components_before_bridge` 를 봐야 한다 — 한 번
    잘못 읽어 「조각 1개」라고 적을 뻔했다.
    """
    from remote30_prototype import build_system_graph
    g, _el, st = build_system_graph(entities, layer_filter=layers,
                                    force_connect=False)
    return (len(g), st.get("components_before_bridge"),
            st.get("bridges_applied"), st.get("components_after_bridge"),
            st.get("layer_filter_used"))


def main() -> int:
    from routes.module_f.subdrawing import graph_payload, parse_subdrawing
    for label, fn in (("계통도", "1. 입력도면 대명동 단위세대 계통도.dxf"),
                      ("기계실", "1. 입력도면 대명동 단위세대 기계실.dxf")):
        path = os.path.join(_D, fn)
        if not os.path.isfile(path):
            print(f"표본 없음: {path}")
            continue
        entities, _parsed = parse_subdrawing(path)
        pay = graph_payload(entities)
        auto = pay.get("auto_layers") or []
        print(f"\n=== {label} · entity {len(entities)}")
        print(f"  자동 배관 레이어: {auto}")
        n, cb, br, ca, used = measure(entities, None)
        print(f"  [섞음]  절점 {n} · 조각 {cb}→{ca} (다리 {br}) · 쓴 것 {used}")
        for nm in auto:
            n1, cb1, br1, ca1, _u = measure(entities, {nm})
            print(f"  [{nm:>8}] 절점 {n1} · 조각 {cb1}→{ca1} (다리 {br1})")
        p1 = graph_payload(entities, layer_filter={auto[0]} if auto else None)
        print(f"  섞음 payload: 절점 {len(pay['nodes'])} · 추측다리 {pay['forced']}"
              f" · 조각 {pay['components']}")
        print(f"  {auto[0] if auto else '-'} payload: 절점 {len(p1['nodes'])}"
              f" · 추측다리 {p1['forced']} · 조각 {p1['components']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
