# -*- coding: utf-8 -*-
"""사용자가 올린 유속초과 KFP 를 우리 솔버로 재해석해 원인을 가른다.

질문: 다운로드한 KFP 를 K-Fire_Solver 로 돌리니 유속 23 m/s — 우리 수렴이 거짓인가,
아니면 가압 조건(고가수조/펌프)이 달라 유량이 커진 것인가.

두 경계조건으로 같은 망을 푼다:
  (A) 설계기준 역산 — 최원단 헤드 방수압 0.1 MPa 고정 (우리 사이징이 쓰는 가정)
  (B) 수원 압력 고정 — KFP 의 wt 노드 수두(고가수조 48.3 m + 10.33 m)를 그대로 인가
      (K-Fire_Solver/EPANET 이 실제로 푸는 조건)
"""
from __future__ import annotations

import json
import math
import sys
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
for p in (BASE, BASE / "core", BASE / "pipenet_converter" / "src"):
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
    solve_network,
    velocity_mps,
)

KFP = BASE / "static" / "combined_89a5a27dbb53_iso_유속초과.kfp"
raw = json.loads(KFP.read_text(encoding="utf-8-sig"))
meta = raw.get("nodes_meta_runtime") or raw.get("nodes_meta") or {}
pd = raw["pipe_data"]

nodes, pipes, nozzles = [], [], []
src = None
for nid, m in meta.items():
    t = (m.get("type_id") or "").lower()
    if t == "wt":
        src = (nid, float(m.get("elevation_m") or 0.0),
               float(m.get("water_level") or 0.0),
               float(m.get("required_pressure_bar") or 0.0))
    nodes.append({"label": nid, "io_node": "Input" if t == "wt" else "No",
                  "elevation": float(m.get("elevation_m") or 0.0)})
    if m.get("k_factor_si"):
        nozzles.append({"in": nid})
for pid, p in pd.items():
    pipes.append({"label": pid, "in": p["start"], "out": p["end"],
                  "dia": float(p.get("nominal_mm") or 0),
                  "length": float(p.get("length_m") or 0.0),
                  "c": float(p.get("C") or 120.0),
                  "eq": float(p.get("equivalent_length") or 0.0)})

print(f"노드 {len(nodes)} · 배관 {len(pipes)} · 헤드 {len(nozzles)}")
print(f"수원 {src[0]}: 표고 {src[1]} m · water_level {src[2]} m · "
      f"required_p {src[3]} bar")
elevs = [n["elevation"] for n in nodes]
print(f"표고 범위 {min(elevs):.1f} ~ {max(elevs):.1f} m")
head_elev = [float(meta[n['in']].get('elevation_m') or 0) for n in nozzles]
print(f"헤드 표고 {min(head_elev):.1f} ~ {max(head_elev):.1f} m")

topo = build_topology(nodes, pipes, nozzles)
roles = classify_pipe_roles(topo, pipes)
limits = [6.0 if r == "branch" else 10.0 for r in roles]
eqlen = {p["label"]: p["eq"] for p in pipes}
bores = [inner_diameter_mm(p["dia"]) for p in pipes]


def report(tag, flows, head_q):
    vs = [velocity_mps(flows[i], bores[i]) for i in range(len(pipes))]
    over = [(i, vs[i]) for i in range(len(pipes)) if vs[i] > limits[i] + 1e-9]
    print(f"\n── {tag}")
    print(f"   헤드유량 {min(head_q):.0f} ~ {max(head_q):.0f} L/min "
          f"(설계 80 대비 x{max(head_q)/80:.2f}) · 총 {sum(head_q):.0f} L/min")
    print(f"   최대유속 {max(vs):.2f} m/s · 상한초과 {len(over)}/{len(pipes)} 구간")
    for i, v in sorted(over, key=lambda t: -t[1])[:8]:
        print(f"     {pipes[i]['label']:>5} {pipes[i]['dia']:>4.0f}A "
              f"({bores[i]:.1f}mm) {roles[i]:<6} Q={flows[i]:8.0f} L/min "
              f"v={v:6.2f} (상한 {limits[i]:.0f})")
    return vs


# (A) 설계기준 역산 — 우리 사이징이 쓰는 가정
solA = solve_network(pipes, topo, equiv_length_m=eqlen)
qA = [solA.head_flow_lpm[h] for h in topo.head_count]
print(f"\n[A] 소스압 {solA.source_pressure_bar:.2f} bar "
      f"(수렴 {solA.converged}, {solA.iterations}회)")
vA = report("A) 최원단 헤드 0.1 MPa 고정 — 우리 역산표 기준", solA.pipe_flow_lpm, qA)


# (B) 수원 압력 고정 — KFP 가 K-solver 에 실제로 주는 조건
def solve_fixed_source(source_bar, damping=0.4, iters=400):
    """소스 압력을 못 박고 q=K√P 고정점 반복. (A) 와 달리 헤드압이 결과로 나온다."""
    hq = {h: NOZZLE_K_FACTOR * math.sqrt(source_bar) * c
          for h, c in topo.head_count.items()}
    asc = list(reversed(topo.order))
    flows = [0.0] * len(pipes)
    press: dict[str, float] = {}
    for _ in range(iters):
        flows = accumulate_flows(topo, hq)
        press = {topo.source: source_bar}
        for i in asc:
            u, d = topo.up[i], topo.down[i]
            if u not in press:
                continue
            drop = hazen_williams_drop_bar(flows[i], pipes[i]["length"] + pipes[i]["eq"],
                                           pipes[i]["c"], bores[i])
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


# 고가수조 총수두 = 표고차 + 수위. 소스 노드 기준 게이지압으로 환산.
tank_bar = src[2] * BAR_PER_M_WATER
print(f"\n[B] 수원 게이지압 {tank_bar:.3f} bar (수위 {src[2]} m) — "
      f"헤드까지 표고강하 {src[1] - min(head_elev):.1f} m 가 추가 가압")
flowsB, hqB, pressB = solve_fixed_source(tank_bar)
vB = report("B) KFP 수원 수두 그대로 — K-solver 가 푸는 조건", flowsB,
            [hqB[h] for h in topo.head_count])
hp = [pressB[n["in"]] for n in nozzles if n["in"] in pressB]
print(f"   헤드 방수압 {min(hp):.2f} ~ {max(hp):.2f} bar (설계 최소 1.0)")

print("\n── 요약")
print(f"   같은 관경인데 최대유속 A {max(vA):.2f} → B {max(vB):.2f} m/s "
      f"(x{max(vB)/max(vA):.2f})")
print(f"   호칭경 분포: {dict(sorted(Counter(p['dia'] for p in pipes).items()))}")
