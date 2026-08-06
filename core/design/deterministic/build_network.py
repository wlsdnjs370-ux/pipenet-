# -*- coding: utf-8 -*-
"""지시서 §9 — C5 한 바퀴. 라우팅(C510~C540) → 신축배관(C550) → 관경(C570)
→ 위상 그래프(C560).

C51x~C57x 모듈은 서로를 부르지 않는다. 순서와 자료 넘김이 곧 §9 의 계약이라
라우트 본문에 흩어 두면 시험할 수 없어 여기 모은다. Flask 를 모른다 — 세션에서
읽고 쓰는 일은 라우트가 한다.

`inner_dia_mm` 을 주입받는 이유는 `pipe_sizing.size_pipes` 와 같다. 내경은 관
재질에 딸린 값이고 재질은 DXF 에 없다. 모듈 A 의 내경표(`hydraulic_solver`)는
평면 import 로 짜여 `core` 가 sys.path 에 올라야 열리므로, 그 부담은 서버를 이미
지고 있는 라우트가 진다.

§9.2 는 "구역 수만큼 독립 트리를 만들고 C540 에서 각 트리를 입상관으로 묶는다"
고 했다. 코어가 여럿이면 입상관도 여럿이고, 그 입상관들을 잇는 지하 수평 주배관은
평면도에 없다 — 지어내면 그 길이가 그대로 마찰손실이 된다(금지사항 G9). 그래서
코어마다 독립된 망을 내고 그 사실을 플래그로 밝힌다.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

from . import emit_graph as EG
from . import pipe_routing as PR
from . import pipe_sizing as PS

__all__ = ["Network", "NetworkResult", "build_networks"]


class _Limits:
    """`plan_branches` 가 보는 두 값. 없으면 채우지 않고 거절한다(금지사항 G9)."""

    FIELDS = ("branch_heads_per_side_max", "head_spacing_square_m")

    def __init__(self, data: dict):
        self.missing = [k for k in self.FIELDS if data.get(k) is None]
        for key in self.FIELDS:
            setattr(self, key, float(data.get(key) or 0.0))


@dataclass
class Network:
    """입상관 하나가 먹이는 망. 급수원이 하나이므로 SDF 한 벌이 여기서 나온다."""

    riser: PR.Riser
    zones: list[dict] = field(default_factory=list)
    topology: Any = None
    tables: Any = None
    metrics: dict = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "riser": self.riser.to_dict(),
            "zones": [z["plan"].to_dict() for z in self.zones],
            "graph": self.tables.to_dict() if self.tables else None,
            "metrics": self.metrics,
        }


@dataclass
class NetworkResult:
    networks: list[Network] = field(default_factory=list)
    flags: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"networks": [n.to_dict() for n in self.networks],
                "flags": self.flags}


def _routing_rooms(rooms, layouts: dict) -> dict[str, dict]:
    """라우팅이 보는 실 = C1 실 폴리곤 + C4 헤드 배치.

    `metrics.axis_deg` 를 실어야 C510 이 실 격자 각도를 후보로 쓸 수 있다. 빼면
    헤드 뭉치의 OBB 각도로 되돌아가 가지배관이 헤드 열을 비껴간다.
    """
    out = {}
    for room in rooms:
        layout = layouts.get(room.id) or {}
        out[room.id] = {
            "id": room.id, "floor": room.floor, "polygon": room.polygon,
            "head_exempt": room.head_exempt,
            "heads": layout.get("heads") or [],
            "metrics": layout.get("metrics") or {},
        }
    return out


def _riser_levels(floors, elevations: dict) -> tuple[list, list[str]]:
    """층 라벨 → 입상관 층 노드. 표고를 못 구한 층은 빼고 돌려준다.

    [문서정합 §9.2 C540] 명세는 표시 좌표와 수리 표고를 "분리 보유" 하라고만 했다.
    순방향 설계에는 눌러 그린 도면이 없어 둘이 같다 — 다르게 만들 근거가 없으므로
    같은 값을 넣되 필드는 나눠 둔다. 나중에 계통도를 눌러 그릴 때 여기만 고친다.
    """
    levels, missing = [], []
    for floor in sorted(floors):
        elev = elevations.get(floor)
        if elev is None:
            missing.append(floor)
            continue
        levels.append(PR.RiserLevel(floor=floor, elevation_m=elev,
                                    display_z_m=elev))
    return levels, missing


def _route_zone(zone: dict, valve: dict, rooms: dict, draft, limits,
                *, cores: list, min_dn: int, grid_mm=None) -> PR.ZonePlan:
    """C510~C540 — 한 구역의 골격. 실패는 `plan.flags` 로 남는다."""
    zone_rooms = [rooms[rid] for rid in zone["rooms"] if rid in rooms]
    floor_rooms = [r for r in rooms.values() if r["floor"] == zone["floor"]]
    point = (float(valve["point"][0]), float(valve["point"][1]))
    # 제 코어는 막지 않는다 — 유수검지장치가 입상관 발치에 서므로 그 코어를 절대
    # 장애물로 두면 주배관이 제 급수원에서 한 걸음도 나오지 못한다. 옆 코어는 막힌다.

    plan = PR.plan_branches(zone["id"], zone_rooms, point, limits)
    # 벽·코어는 구역이 아니라 층 전체가 준다 — 옆 구역의 코어도 통과 불가다.
    field_ = PR.RouteField(plan.axes, floor_rooms, cores=cores,
                           doors=draft.virtual_edges, grid_mm=grid_mm)
    PR.route_cross_mains(plan, field_, min_dn=min_dn)
    PR.route_main(plan, field_, point, min_dn=min_dn)
    return plan


def _size(topo, constraints: dict,
          inner_dia_mm: Callable[[int], float]) -> tuple[dict, list[dict]]:
    """C570 — 담당 헤드 수 → 별표1 최소 → 유속 승급 → 단조화."""
    heads_at = {node_id: 1 for node_id in topo.head_nodes.values()}
    loads, unreached = PS.edge_load(topo.adjacency(), topo.source, heads_at)
    roles = {edge: topo.edges[EG._key(*edge)].role for edge in loads}
    sizes, flags = PS.size_pipes(loads, roles, constraints, inner_dia_mm)
    flags += PS.enforce_monotonic(sizes, topo.source, inner_dia_mm)
    if unreached:
        flags.append({
            "code": "HEAD_UNREACHED", "heads": sorted(unreached),
            "message": f"급수원에서 배관으로 닿지 않는 헤드가 {len(unreached)}개 "
                       f"있습니다 — 그만큼 미방호입니다.",
        })
    return sizes, flags


def build_networks(draft, design: dict, constraints: dict, *,
                   inner_dia_mm: Callable[[int], float],
                   meta=(), grid_mm=None) -> NetworkResult:
    """C3·C4 산출물 → 코어별 배관망. 근거가 없으면 채우지 않고 플래그를 남긴다."""
    result = NetworkResult()
    limits = _Limits(constraints)
    if limits.missing:
        result.flags.append({
            "code": "CONSTRAINTS_INCOMPLETE", "fields": limits.missing,
            "message": "기준에 가지배관 값이 없습니다: " + ", ".join(limits.missing),
        })
        return result

    valve_by_id = {v["id"]: v for v in (design.get("valves") or [])}
    layouts = {lay["room_id"]: lay for lay in (design.get("heads") or [])}
    rooms = _routing_rooms(draft.rooms, layouts)
    cores = [c.to_dict() for c in draft.cores]
    core_by_id = {c["id"]: c for c in cores}
    elevations, _unknown = draft.floor_elevations()
    min_dn = int(constraints.get("cross_main_min_dn") or 0)

    by_core: dict[str, list[dict]] = {}
    for zone in (design.get("zones") or []):
        valve = valve_by_id.get(zone.get("valve_id"))
        if valve is None:
            result.flags.append({
                "code": "ZONE_VALVE_MISSING", "zone_id": zone["id"],
                "message": f"{zone['id']} 의 유수검지장치를 찾을 수 없습니다.",
            })
            continue
        by_core.setdefault(str(valve.get("core_id") or ""), []).append(zone)

    if len(by_core) > 1:
        result.flags.append({
            "code": "MULTIPLE_RISERS", "cores": sorted(by_core),
            "message": f"입상관이 {len(by_core)}개입니다. 이들을 잇는 지하 수평 "
                       "주배관은 평면도에 없어 지어내지 않았습니다 — 코어마다 "
                       "독립된 망으로 냅니다.",
        })

    for core_id, zones in sorted(by_core.items()):
        core = core_by_id.get(core_id)
        if core is None:
            result.flags.append({
                "code": "CORE_MISSING", "core_id": core_id,
                "zones": [z["id"] for z in zones],
                "message": f"코어 {core_id} 가 없어 입상관을 세울 수 없습니다.",
            })
            continue
        network = _build_one(core, zones, valve_by_id, rooms,
                             draft, constraints, limits, cores=cores,
                             elevations=elevations, min_dn=min_dn,
                             inner_dia_mm=inner_dia_mm, meta=meta,
                             grid_mm=grid_mm)
        if network is not None:
            result.networks.append(network)
    return result


def _build_one(core: dict, zones: list[dict], valve_by_id: dict, rooms: dict,
               draft, constraints: dict, limits, *,
               cores: list, elevations: dict, min_dn: int,
               inner_dia_mm, meta, grid_mm) -> Network | None:
    levels, no_elev = _riser_levels({z["floor"] for z in zones}, elevations)
    riser = PR.plan_riser(core, levels)
    network = Network(riser=riser)

    flags: list[dict] = []
    if no_elev:
        flags.append({
            "code": "FLOOR_ELEVATION_MISSING", "floors": sorted(no_elev),
            "message": "층고를 못 구한 층이 있습니다: " + ", ".join(sorted(no_elev))
                       + ". 입상관 길이는 그대로 낙차 압력이 되므로 지어내지 "
                         "않았습니다 — 그 층의 구역은 빠집니다.",
        })

    head_ids: set[str] = set()
    for zone in sorted(zones, key=lambda z: z["id"]):
        valve = valve_by_id[zone["valve_id"]]
        try:
            plan = _route_zone(zone, valve, rooms, draft, limits,
                               cores=[c for c in cores if c["id"] != core["id"]],
                               min_dn=min_dn, grid_mm=grid_mm)
        except ValueError as exc:
            # `zone_axes` 는 헤드 없는 구역을 거절한다. 그것은 실패가 아니라
            # 놓을 것이 없다는 뜻이므로 망을 통째로 죽이지 않는다.
            flags.append({"code": "ZONE_NO_HEADS", "zone_id": zone["id"],
                          "message": f"{zone['id']}: {exc}"})
            continue
        flags.extend(plan.flags)
        network.zones.append({"plan": plan, "floor": zone["floor"],
                              "zone_id": zone["id"]})
        head_ids.update(plan.head_ab)

    heads = [h for lay in rooms.values() for h in lay["heads"]
             if h["id"] in head_ids]
    conns, conn_flags = PS.head_connections(heads, draft.rooms, constraints)
    flags.extend(conn_flags)

    topo = EG.build_topology(riser, network.zones, heads, conns)
    merged = EG.absorb_bends(topo)
    network.topology = topo

    sizes: dict = {}
    if topo.source is not None:
        sizes, size_flags = _size(topo, constraints, inner_dia_mm)
        flags.extend(size_flags)
        for message in PR.check_tournament(
                topo.adjacency(), topo.source,
                roots=[nid for nid, n in topo.nodes.items() if n.kind == "branch"]):
            flags.append({"code": "TOURNAMENT_LAYOUT", "message": message})

    tables = EG.emit_tables(topo, sizes, constraints, meta=list(meta))
    tables.flags = flags + tables.flags
    network.tables = tables
    network.metrics = {
        "zones": len(network.zones), "heads": len(heads),
        "nodes": len(tables.nodes), "pipes": len(tables.pipes),
        "fittings": len(tables.fittings), "flex": len(tables.equipment),
        "absorbed_bends": merged,
        "length_m": round(sum(float(p["length"]) for p in tables.pipes), 2),
    }
    return network
