# -*- coding: utf-8 -*-
"""두 .kfp 가 «같은 망» 인가 — 노드 id 부여 순서와 무관한 정준 비교.

전체망 변환은 set 순회가 끼어 실행마다 노드 번호가 다르게 붙는다(크기는
같은데 해시가 다른 이유). 그래서 바이트 비교로는 «최적화 전후 동등» 을 못
판정한다 — 좌표로 이름을 지우고 비교한다.

    노드: (좌표, 타입, K, 고도, …) 의 multiset
    배관: (양끝 좌표 정렬쌍, 규격·길이·등가길이·부속) 의 multiset

실행: python data/_compare_kfp_canonical.py A.kfp B.kfp
"""
from __future__ import annotations

import json
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def _r(v, nd=4):
    try:
        return round(float(v), nd)
    except (TypeError, ValueError):
        return v


def canon(path):
    k = json.load(open(path, encoding="utf-8"))
    nodes = k.get("nodes_meta_runtime") or {}
    pipes = k.get("pipe_data") or {}

    def nkey(nid):
        n = nodes.get(nid) or {}
        c = n.get("coords") or [0, 0, 0]
        return tuple(_r(v) for v in c)

    node_sig = Counter()
    for nid, n in nodes.items():
        c = tuple(_r(v) for v in (n.get("coords") or [0, 0, 0]))
        node_sig[(c, str(n.get("type_id") or ""), str(n.get("type") or ""),
                  _r(n.get("elevation_m")), _r(n.get("k_factor_si")),
                  _r(n.get("required_pressure_bar")),
                  bool(n.get("is_active")),
                  str(n.get("head_spec_name") or ""))] += 1

    pipe_sig = Counter()
    for pid, p in pipes.items():
        a, b = nkey(p.get("start")), nkey(p.get("end"))
        ends = tuple(sorted((a, b)))
        pipe_sig[(ends, str(p.get("type") or ""), _r(p.get("diameter")),
                  _r(p.get("nominal_mm")), _r(p.get("length_m")),
                  _r(p.get("equivalent_length")), _r(p.get("C")),
                  _r(p.get("roughness_mm")),
                  tuple(sorted(map(str, p.get("fittings") or ()))))] += 1
    return node_sig, pipe_sig


def diff(name, a: Counter, b: Counter) -> int:
    only_a = a - b
    only_b = b - a
    n = sum(only_a.values()) + sum(only_b.values())
    if n:
        print(f"[{name}] 다름 — A에만 {sum(only_a.values())} · "
              f"B에만 {sum(only_b.values())}")
        for sig, cnt in list(only_a.items())[:3]:
            print(f"  A에만 ×{cnt}: {sig}")
        for sig, cnt in list(only_b.items())[:3]:
            print(f"  B에만 ×{cnt}: {sig}")
    else:
        print(f"[{name}] 동일 · {sum(a.values()):,}개")
    return n


def main():
    pa, pb = sys.argv[1], sys.argv[2]
    na, ea = canon(pa)
    nb, eb = canon(pb)
    bad = diff("노드", na, nb) + diff("배관", ea, eb)
    print("→ " + ("정준 동일 — 같은 망이다" if bad == 0
                  else f"차이 {bad}건 — 같은 망이 아니다"))
    raise SystemExit(0 if bad == 0 else 1)


if __name__ == "__main__":
    main()
