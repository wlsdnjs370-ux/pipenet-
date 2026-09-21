# -*- coding: utf-8 -*-
"""FX 신축배관 스텁이 헤드 노드에 직렬로 끼는 결함의 크기를 실측한다.

_materialize_fx_pipes 는 부모 파이프 P 의 끝점만 새 노드 F 로 옮기고, 같은 헤드
노드 H 에 붙어 있던 나머지 배관은 그대로 둔다. 헤드가 가지관 중간 분기점이면
하류 전체 유량이 20A FX 스텁을 통과한다 — 사용자가 본 23 m/s 의 직접 원인.

여기서는 사용자 제출 KFP 를 (a) 그대로, (b) FX 스텁을 진짜 스퍼로 교정한 뒤
같은 경계조건(고가수조 수두 고정)으로 풀어 최대유속을 비교한다.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
for p in (BASE, BASE / "core"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))
sys.stdout.reconfigure(encoding="utf-8")

from hydraulic_solver import (  # noqa: E402
    BAR_PER_M_WATER,
    NOZZLE_K_FACTOR,
    accumulate_flows,
    build_topology,
    classify_pipe_roles,
    hazen_williams_drop_bar,
    inner_diameter_mm,
    velocity_mps,
)

KFP = BASE / "static" / "combined_89a5a27dbb53_iso_유속초과.kfp"
raw = json.loads(KFP.read_text(encoding="utf-8-sig"))
meta = raw.get("nodes_meta_runtime") or raw.get("nodes_meta") or {}
pdata = raw["pipe_data"]

heads = {n for n, m in meta.items() if m.get("k_factor_si")}
src = next(n for n, m in meta.items() if (m.get("type_id") or "").lower() == "wt")
tank_bar = float(meta[src].get("water_level") or 0.0) * BAR_PER_M_WATER


def build(fix_spur: bool):
    """KFP → 솔버 dict. fix_spur=True 면 헤드 노드를 진짜 말단으로 되돌린다."""
    nodes = [{"label": n, "io_node": "Input" if n == src else "No",
              "elevation": float(m.get("elevation_m") or 0.0)}
             for n, m in meta.items()]
    pipes = [{"label": pid, "in": p["start"], "out": p["end"],
              "dia": float(p.get("nominal_mm") or 0), "length": float(p.get("length_m") or 0),
              "c": float(p.get("C") or 120.0),
              "eq": float(p.get("equivalent_length") or 0.0)}
             for pid, p in pdata.items()]
    if fix_spur:
        # 헤드 H 에 붙은 배관 중 FX 스텁(20A)만 남기고, 나머지는 스텁의 반대편
        # 노드 F 로 옮긴다 → H 는 노즐만 달린 진짜 말단이 된다.
        stub_far = {}
        for p in pipes:
            if int(p["dia"]) != 20:
                continue
            h, f = (p["out"], p["in"]) if p["out"] in heads else (p["in"], p["out"])
            stub_far[h] = f
        for p in pipes:
            if int(p["dia"]) == 20:
                continue
            for end in ("in", "out"):
                if p[end] in stub_far:
                    p[end] = stub_far[p[end]]
    nozzles = [{"in": n} for n in heads]
    return nodes, pipes, nozzles


def solve_fixed_source(topo, pipes, bores, damping=0.4, iters=600):
    """소스 압력을 못 박고 q=K√P 고정점 반복 — K-solver/EPANET 이 푸는 조건."""
    hq = {h: NOZZLE_K_FACTOR * math.sqrt(tank_bar) * c for h, c in topo.head_count.items()}
    asc = list(reversed(topo.order))
    flows, press = [0.0] * len(pipes), {}
    for _ in range(iters):
        flows = accumulate_flows(topo, hq)
        press = {topo.source: tank_bar}
        for i in asc:
            u, d = topo.up[i], topo.down[i]
            if u not in press:
                continue
            drop = hazen_williams_drop_bar(
                flows[i], pipes[i]["length"] + pipes[i]["eq"], pipes[i]["c"], bores[i])
            lift = (topo.elevation.get(d, 0.0) - topo.elevation.get(u, 0.0)) * BAR_PER_M_WATER
            press[d] = press[u] - drop - lift
        delta = 0.0
        for h in topo.head_count:
            if h not in press:
                continue
            tgt = NOZZLE_K_FACTOR * math.sqrt(max(0.0, press[h])) * topo.head_count[h]
            new = hq[h] + damping * (tgt - hq[h])
            delta = max(delta, abs(new - hq[h]))
            hq[h] = new
        if delta < 0.05:
            break
    return flows, hq, press


for tag, fix in (("① 현재 출력물 그대로", False), ("② FX 스텁을 진짜 스퍼로 교정", True)):
    nodes, pipes, nozzles = build(fix)
    topo = build_topology(nodes, pipes, nozzles)
    roles = classify_pipe_roles(topo, pipes)
    limits = [6.0 if r == "branch" else 10.0 for r in roles]
    bores = [inner_diameter_mm(p["dia"]) for p in pipes]
    flows, hq, press = solve_fixed_source(topo, pipes, bores)
    vs = [velocity_mps(flows[i], bores[i]) for i in range(len(pipes))]
    over = [i for i in range(len(pipes)) if vs[i] > limits[i] + 1e-9]
    stub = [i for i in range(len(pipes)) if int(pipes[i]["dia"]) == 20]
    print(f"\n── {tag}")
    print(f"   헤드유량 {min(hq.values()):.0f} ~ {max(hq.values()):.0f} L/min · "
          f"총 {sum(hq.values()):.0f} L/min")
    print(f"   최대유속 {max(vs):.2f} m/s · 상한초과 {len(over)}/{len(pipes)}")
    print(f"   20A FX 스텁 통과유량 {min(flows[i] for i in stub):.0f} ~ "
          f"{max(flows[i] for i in stub):.0f} L/min · 유속 최대 {max(vs[i] for i in stub):.2f} m/s")
    for i in sorted(over, key=lambda j: -vs[j])[:5]:
        print(f"     {pipes[i]['label']:>5} {pipes[i]['dia']:>4.0f}A {roles[i]:<6} "
              f"Q={flows[i]:7.0f} v={vs[i]:6.2f} (상한 {limits[i]:.0f})")
