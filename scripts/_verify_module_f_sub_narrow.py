# -*- coding: utf-8 -*-
"""[계통도 꼬임] 좁히기가 **실도면에서** 무엇을 바꾸는지 잰다.

`scripts/_probe_sub_bridge_path.py` 가 병을 잡았다면(경로가 LSP→HSP→LSP 로
계통을 넘나든다), 여기는 약이 듣는지를 본다: 사람이 찍는 두 점을 그대로 주고
①좁히기 전 ②좁힌 뒤 의 «넘나듦 횟수·연장» 을 나란히 적는다.
"""
from __future__ import annotations

import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_D = os.path.join(_ROOT, "routes", "제출용[최종]")
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(name)
    return ok


def hops(entities, layers, a, b):
    from remote30_prototype import (_auto_pipe_layer_filter,
                                    _nearest_graph_node, _shortest_path,
                                    build_system_graph)
    g, el, st = build_system_graph(entities, layer_filter=layers,
                                   force_connect=True)
    fk = set()
    for (ea, eb) in (st.get("forced_bridge_edges") or ()):
        ka, kb = (int(ea[0]), int(ea[1])), (int(eb[0]), int(eb[1]))
        fk.add((min(ka, kb), max(ka, kb)))
    own = {}
    for nm in sorted(_auto_pipe_layer_filter(entities)):
        _g1, el1, _s = build_system_graph(entities, layer_filter={nm},
                                          force_connect=False)
        own[nm] = set(el1)
    p = _shortest_path(g, el, _nearest_graph_node(g, a),
                       _nearest_graph_node(g, b), penalty_keys=fk)
    if not p:
        return None, None, []
    tot, seq = 0.0, []
    for u, v in zip(p, p[1:]):
        k = (min(u, v), max(u, v))
        ln = el.get(k)
        tot += ln if ln is not None else math.hypot(u[0] - v[0], u[1] - v[1])
        hit = [nm for nm, ks in own.items() if k in ks]
        s = hit[0] if len(hit) == 1 else "?"
        if not seq or seq[-1] != s:
            seq.append(s)
    real = [s for s in seq if s != "?"]
    runs = [real[0]] if real else []
    for s in real[1:]:
        if s != runs[-1]:
            runs.append(s)
    return tot / 1000.0, len(runs) - 1, runs


def main() -> int:
    from routes.module_f.subdrawing import parse_subdrawing, pick_system_layer
    for label, fn in (("계통도", "1. 입력도면 대명동 단위세대 계통도.dxf"),
                      ("기계실", "1. 입력도면 대명동 단위세대 기계실.dxf")):
        path = os.path.join(_D, fn)
        if not os.path.isfile(path):
            print(f"표본 없음: {path}")
            return 0
        ents, _ = parse_subdrawing(path)
        # 사람은 «배관 근처» 를 찍는다 — 도면 한복판을 찍는 것이 아니다.
        # 한 계통(첫 자동 레이어)의 아래끝·위끝을 골라 그 자리를 흉내낸다.
        from remote30_prototype import (_auto_pipe_layer_filter,
                                        build_system_graph)
        auto = sorted(_auto_pipe_layer_filter(ents))
        g0, _e0, _s0 = build_system_graph(ents, layer_filter={auto[0]},
                                          force_connect=True)
        ns = sorted(g0, key=lambda n: n[1])
        # 정확히 절점 위가 아니라 살짝 빗나가게 찍는다(사람 손은 안 맞는다).
        a = (ns[0][0] + 120.0, ns[0][1] - 90.0)
        b = (ns[-1][0] - 80.0, ns[-1][1] + 140.0)
        print(f"\n=== {label} · 흉내낸 클릭 = «{auto[0]}» 아래끝·위끝")
        m0, h0, r0 = hops(ents, None, a, b)
        print(f"  [좁히기 전] 연장 {m0:.1f} m · 넘나듦 {h0}회 · {' → '.join(r0)}")
        nm, diag = pick_system_layer(ents, a, b)
        print(f"  고른 계통: {nm} — {diag.get('reason')}")
        for r in diag.get("candidates") or ():
            print(f"      {r['layer']:>12} 절점 {r['nodes']:>4}"
                  f" · 붙은거리 {r.get('snap_mm')}"
                  f" · {'연장 %s m' % r['path_m'] if r.get('path_m') else '×'}")
        if not check(f"{label} — 계통 하나를 골랐다", bool(nm)):
            continue
        m1, h1, r1 = hops(ents, {nm}, a, b)
        print(f"  [좁힌 뒤 ] 연장 {m1:.1f} m · 넘나듦 {h1}회 · {' → '.join(r1)}")
        check(f"{label} — 경로가 계통을 넘나들지 않는다", h1 == 0, f"{h1}회")
        check(f"{label} — 넘나듦이 줄었다", h1 <= h0, f"{h0} → {h1}")

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("좁히기 — 실도면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
