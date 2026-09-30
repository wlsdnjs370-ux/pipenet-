"""Select every defined head in an explicit CAD-mm region, without a quota.

Regions select head centres. Cyclic area geometry and supply connectors are
scoped separately by ``area_network``; supply routes may leave the regions.
Coincident CAD symbols follow the existing flow model's physical-head aliases.
Disconnected symbols are reported, never replaced by heads outside the region.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from .flow import FlowTree
from .regions import normalize_zones, zone_contains


@dataclass(frozen=True)
class AreaHeads:
    """Source disk identities, physical representatives and unresolved symbols."""

    zones: list
    candidates: tuple[int, ...]
    heads: tuple[int, ...]
    disconnected: tuple[int, ...]
    aliases: tuple[tuple[int, int], ...]


def area_candidates(disks: Sequence[Sequence[float]], zones: list) -> tuple[int, ...]:
    """Return the inclusive union; an absent or empty area is an error."""
    regions = normalize_zones(zones)
    if not regions:
        raise ValueError("먼저 영역을 지정하세요. 지정한 영역 안의 모든 헤드를 사용합니다.")
    result = tuple(i for i, p in enumerate(disks) if any(zone_contains(z, p) for z in regions))
    if not result:
        raise ValueError("지정한 영역 안에 정의된 헤드가 없습니다. 영역과 객체 정의를 확인하세요.")
    return result


def select_area_heads(disks: Sequence[Sequence[float]], zones: list, flow: FlowTree) -> AreaHeads:
    """Select all physical heads, preserving identity and explicit failure lists."""
    regions = normalize_zones(zones)
    candidates = area_candidates(disks, regions)
    representatives: dict[int, int] = {}
    disconnected, aliases = [], []
    for hi in candidates:
        if hi not in flow.head_node:
            disconnected.append(hi)
            continue
        identity = flow.representatives[hi]
        representative = representatives.setdefault(identity, hi)
        if representative != hi:
            aliases.append((hi, representative))
    return AreaHeads(regions, candidates, tuple(representatives.values()),
                     tuple(disconnected), tuple(aliases))


def area_corridor(selection: AreaHeads, disks: Sequence[Sequence[float]], flow: FlowTree) -> dict:
    """Adapt an all-head result to the established extraction/table contract.

    Distances are diagnostics only, never a ranking or head-count limit.
    Tree capacity remains the full-network load used by the existing bore rules.
    Cycle callers replace edges with the region-local graph and supply paths.
    """
    if selection.disconnected or not selection.heads:
        raise ValueError("영역 안에 급수원과 연결되지 않은 헤드가 있습니다.")
    heads = list(selection.heads)
    distances = {hi: flow.distance_mm[flow.head_node[hi]] for hi in heads}
    farthest = max(heads, key=lambda hi: (distances[hi], -hi))
    path = flow.path(flow.head_node[farthest])
    loads = flow.selected_loads(heads)
    xs, ys = [disks[h][0] for h in heads], [disks[h][1] for h in heads]
    width, height = (max(xs)-min(xs))/1000, (max(ys)-min(ys))/1000
    aliases_by_head: dict[int, int] = {}
    for _, hi in selection.aliases:
        aliases_by_head[hi] = aliases_by_head.get(hi, 1) + 1
    return {
        "selection_mode": "area_all", "selection_basis": "all_heads_in_regions",
        "heads": heads, "area_candidates": list(selection.candidates),
        "aliases": [list(a) for a in selection.aliases], "zones": selection.zones,
        "candidates": len(selection.candidates), "reachable": len(heads), "unreachable": 0,
        "edges": set(loads), "loads": loads, "physical_loads": dict(flow.loads),
        "nodes": {n for e in loads for n in e} | set(flow.roots),
        "flow_revision": flow.revision, "flow_report": flow.report(),
        "worst_head": farthest, "worst_path": path, "dists": distances,
        "worst_path_m": round(distances[farthest]/1000, 2),
        "far_m": round(max(distances.values())/1000, 2),
        "near_m": round(min(distances.values())/1000, 2),
        "span_m": round(max(width, height), 2), "area_w_m": round(width, 2),
        "area_h_m": round(height, 2), "area_m2": round(width*height, 1),
        "total_m": round(sum(flow.lengths_mm[e] for e in loads)/1000, 2),
        "max_load": max(loads.values(), default=0), "merged": len(selection.aliases),
        "merged_xy": [[*disks[h][:2], count] for h, count in aliases_by_head.items()],
        "hydraulically_verified": False,
    }
