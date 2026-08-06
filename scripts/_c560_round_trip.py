# -*- coding: utf-8 -*-
"""지시서 §13 — C560 왕복 검증. 그래프 → SDF → 재파싱 → 대조.

단위 시험은 우리가 만든 dict 를 우리가 읽는 것이라, 필드 이름이 맞아도 **모듈 A
방출기가 그 dict 를 실제로 삼키는지**는 말해 주지 못한다. 여기서는 진짜
`emit_full_sdf` 로 SDF 를 쓰고 `parse_sdf` 로 다시 읽어, 노드 수·배관 수·총
배관장·헤드 수·관경이 살아 돌아오는지 본다.

`kfp_sdf_converter.round_trip_check` 는 KFP 를 기점으로 도는 함수라 여기서 쓸 수
없다(우리 기점은 그래프다). 같은 항목을 같은 방식으로 잰다.
"""
from __future__ import annotations

import math
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "core"))

from core.design.deterministic import emit_graph as EG      # noqa: E402
from core.design.deterministic import pipe_routing as PR    # noqa: E402
from core.design.deterministic import pipe_sizing as PS     # noqa: E402
from core.design.deterministic.zoning import FLEX_SPEC      # noqa: E402
from core.remote30_full_network import (                    # noqa: E402
    CombinedTables, ProjectContext, ZoneSpec, ZoneType, emit_full_sdf)
from kfp_sdf_converter import parse_sdf                     # noqa: E402

STEP = 3200.0
CONSTRAINTS = {"flow_lpm_min": 80.0,
               "velocity_limit_mps": {"branch": 6.0, "other": 10.0},
               "cross_main_min_dn": 40}


def _design():
    """2 층 × 구역 하나 × 가지배관 둘. 상향식·하향식을 섞어 신축배관도 태운다."""
    axes = PR.ZoneAxes(theta=0.0, cross_span_mm=20000.0, branch_span_mm=12800.0)
    riser = PR.Riser(id="RS-C1", core_id="C1", point=(0.0, 0.0), levels=[
        PR.RiserLevel(floor=f, elevation_m=3.5 * i, display_z_m=3.5 * i)
        for i, f in enumerate(("1F", "2F"))])

    heads, conns, zones = [], [], []
    for floor in ("1F", "2F"):
        cross = PR.CrossMain(id=f"CM-{floor}", b_mm=0.0, a_span=(5000.0, 21000.0),
                             path=[[5000.0, 0.0], [21000.0, 0.0]])
        branches = []
        for j, tee_a in enumerate((9000.0, 17000.0), start=1):
            side_ids = []
            for k in range(-3, 4):
                if k == 0:
                    continue
                head_id = f"H-{floor}-{j}-{k}"
                heads.append({"id": head_id, "x": tee_a, "y": k * STEP})
                side_ids.append((k, head_id))
                pendent = (k % 2 == 0)
                conns.append({
                    "head_id": head_id, "room_id": f"R-{floor}",
                    "orientation": "pendent" if pendent else "upright",
                    "kind": "flex" if pendent else "direct",
                    "flex": dict(FLEX_SPEC) if pendent else None})
            branches.append(PR.Branch(
                id=f"BR-{floor}-{j}", cross_id=cross.id, a_mm=tee_a,
                tee=(tee_a, 0.0),
                left=[h for k, h in side_ids if k > 0],
                right=[h for k, h in side_ids if k < 0]))
        main = PR.MainRun(id=f"MN-{floor}", zone_id=f"Z-{floor}",
                          valve=(1000.0, 0.0),
                          path=[[1000.0, 0.0], [5000.0, 0.0]])
        zones.append({"plan": PR.ZonePlan(
            zone_id=f"Z-{floor}", axes=axes, crosses=[cross],
            branches=branches, main=main), "floor": floor})
    return riser, zones, heads, conns


def build() -> EG.GraphTables:
    riser, zones, heads, conns = _design()
    topo = EG.build_topology(riser, zones, heads, conns)
    merged = EG.absorb_bends(topo)

    heads_at = {node_id: 1 for node_id in topo.head_nodes.values()}
    loads, unreached = PS.edge_load(topo.adjacency(), topo.source, heads_at)
    roles = {e: topo.edges[EG._key(*e)].role for e in loads}
    sizes, size_flags = PS.size_pipes(loads, roles, CONSTRAINTS,
                                      lambda dn: dn * 1.0)
    clamps = PS.enforce_monotonic(sizes, topo.source, lambda dn: dn * 1.0)

    print(f"위상: 노드 {len(topo.nodes)} / 간선 {len(topo.edges)} / "
          f"헤드 {len(topo.head_nodes)} / 흡수한 통과점 {merged}")
    print(f"닿지 않은 헤드 {len(unreached)} / 관경 플래그 {len(size_flags)} / "
          f"단조화 조정 {len(clamps)}")
    for flag in topo.flags:
        print(f"  플래그 {flag['code']}: {flag['message']}")
    return EG.emit_tables(topo, sizes, CONSTRAINTS,
                          meta=[("Project", "C560 왕복 검증")])


def main() -> int:
    tables = build()
    for flag in tables.flags:
        print(f"  방출 플래그 {flag['code']}: {flag['message']}")

    combined = CombinedTables(
        nodes=tables.nodes, pipes=tables.pipes, nozzles=tables.nozzles,
        fittings=tables.fittings, equipment=tables.equipment, meta=tables.meta)
    ctx = ProjectContext(zone_spec=ZoneSpec(zone_type=ZoneType.LSP_GRAVITY),
                         project_title="C560 왕복 검증")

    out = Path(tempfile.gettempdir()) / "_c560_round_trip.sdf"
    emit_full_sdf(combined, out, ctx=ctx)
    net = parse_sdf(out)

    # 모듈 A 는 신축배관 등가길이를 실배관 한 토막으로 편다(`_materialize_fx_pipes`).
    # 그래서 배관 수와 총연장은 그만큼 **늘어나는 것이 정상**이다. 우리가 낸 것이
    # 그대로 살아 있는지는 라벨로 맞대어 본다.
    fx_len = sum(float(eq["drawing_len_mm"]) / 1000.0 for eq in tables.equipment)
    len_before = sum(float(p["length"]) for p in tables.pipes)
    len_after = sum(p.length_m for p in net.pipes.values())
    dn_before = {f"P{p['label']}": int(p["dia"]) for p in tables.pipes}
    dn_after = {p.id: p.nominal_mm for p in net.pipes.values()}
    same_dn = sum(1 for k, v in dn_before.items() if dn_after.get(k) == v)
    heads_after = sum(1 for n in net.nodes.values() if n.kind == "head")

    print(f"\nSDF: {out}  ({out.stat().st_size:,} bytes)")
    print(f"  노드   {len(tables.nodes):5d} → {len(net.nodes):5d}"
          f"  (신축배관 노드 +{len(tables.equipment)})")
    print(f"  배관   {len(tables.pipes):5d} → {len(net.pipes):5d}"
          f"  (신축배관 실배관 +{len(tables.equipment)})")
    print(f"  헤드   {len(tables.nozzles):5d} → {heads_after:5d}")
    print(f"  총배관장 {len_before:9.3f} m → {len_after:9.3f} m "
          f"(신축배관 실배관 {fx_len:.3f} m 포함)")
    print(f"  관경 일치 {same_dn}/{len(dn_before)}")
    print(f"  부속류 {len(tables.fittings)} / 신축배관 {len(tables.equipment)}")

    ok = (len(net.pipes) == len(tables.pipes) + len(tables.equipment)
          and len(net.nodes) == len(tables.nodes) + len(tables.equipment)
          and heads_after == len(tables.nozzles)
          and math.isclose(len_before + fx_len, len_after, abs_tol=0.01)
          and same_dn == len(dn_before))
    print("\n왕복 " + ("일치" if ok else "불일치"))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
