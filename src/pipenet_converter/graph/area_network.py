"""Region-local cyclic graph plus explicit supply connectors, in CAD mm.

Unlike an operating-head scenario on the full network, an area extraction must
not retain every possible out-of-area detour. Keep source edges intersecting
the region, then connect each local component by one canonical supply path.
Existing node/edge identities and declared lengths are never clipped or moved.
Boundary-crossing source segments are therefore retained whole and reported.
This is a user-scoped extraction, not a hydraulically equivalent full network.
"""
from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass
import math

import networkx as nx

from ..dxf.work_region import WorkRegion
from .flow import Edge, edge_key
from .network import FlowNetwork
from .regions import normalize_zones

AREA_SCOPE = "region_and_supply_v1"


@dataclass(frozen=True)
class AreaNetwork:
    """Auditable source-edge selection; no invented boundary junctions."""

    edges: frozenset[Edge]
    local_edges: frozenset[Edge]
    supply_edges: frozenset[Edge]
    boundary_edges: frozenset[Edge]
    entries: tuple[int, ...]

    def report(self, network: FlowNetwork) -> dict:
        """Return counts and identities needed to inspect the extraction scope."""
        nodes = {n for e in self.edges for n in e} | set(network.reference.roots)
        return {
            "policy": AREA_SCOPE,
            "local_edges": len(self.local_edges),
            "supply_edges": len(self.supply_edges),
            "boundary_edges": [list(e) for e in sorted(self.boundary_edges)],
            "entry_nodes": list(self.entries),
            "excluded_outside_edges": len(network.edges - self.edges),
            "edges": len(self.edges),
            "cycle_rank": len(self.edges) - len(nodes) + 1,
            "length_m": round(sum(network.reference.lengths_mm[e] for e in self.edges) / 1000, 3),
            "hydraulically_equivalent_to_full_network": False,
        }


def select_area_network(network: FlowNetwork, points: Sequence[Sequence[float]],
                        heads: Iterable[int], zones: list) -> AreaNetwork:
    """Preserve local loops, adding only one real supply path per component.

    Canonical path priority is the same as FlowTree: fewer inferred links,
    then source length, hops and stable node ID. A crossing is not a junction.
    Every supplied local component is kept, including local return branches;
    unconnected source geometry is not silently attached by proximity.
    """
    regions = [WorkRegion(z) for z in normalize_zones(zones)]
    if not regions:
        raise ValueError("영역을 지정하고 배관망을 다시 추출하세요.")
    ref = network.reference
    selected = set(heads)
    if not selected or not selected <= ref.head_node.keys():
        raise ValueError("영역 안 헤드가 현재 급수원에 연결되지 않았습니다.")
    local, boundary = set(), set()
    for edge in sorted(network.edges):
        a, b = (points[n][:2] for n in edge)
        # Zero-length aliases are real topology and must survive if inside.
        if math.dist(a, b) < 1e-7:
            if any(region.contains(a) for region in regions):
                local.add(edge)
        elif any(region.clip(a, b) for region in regions):
            local.add(edge)
            if not any(region.whole_segment(a, b) for region in regions):
                boundary.add(edge)

    graph = nx.Graph()
    graph.add_edges_from(sorted(local))
    graph.add_nodes_from(ref.head_node[ref.representatives[h]] for h in selected)
    root = ref.roots[0]
    # Compute the canonical accumulated inferred-link cost only once.
    inferred_cost = {root: 0}
    for node in sorted(ref.parent, key=lambda n: (ref.hop[n], n)):
        parent = ref.parent[node]
        inferred_cost[node] = inferred_cost[parent] + int(edge_key(node, parent) in network.inferred)

    def priority(node: int) -> tuple[int, float, int, int]:
        return inferred_cost[node], ref.distance_mm[node], ref.hop[node], node

    entries, connectors = [], set()
    for component in sorted(nx.connected_components(graph), key=lambda c: min(c)):
        entry = min(component, key=priority)
        entries.append(entry)
        path = ref.path(entry)
        connectors.update(edge_key(a, b) for a, b in zip(path, path[1:]))
    edges = frozenset(local | connectors)
    network.resolve_selection(selected, source_edges=edges)
    return AreaNetwork(edges, frozenset(local), frozenset(connectors - local),
                       frozenset(boundary), tuple(entries))
