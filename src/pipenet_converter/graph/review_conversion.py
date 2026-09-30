"""Cycle-preserving calculation-input graph. Not a hydraulic solver.

CAD XY is retained; declared source lengths override geometric lengths. Head
connections are local grafts, never a spanning-tree expansion. An explicit datum
and configured, arc-evidenced rises define review elevations, not surveyed Z.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import math
from typing import Mapping, Sequence

import networkx as nx

from .flow import edge_key
from .network import FlowNetwork


@dataclass(frozen=True)
class ReviewNode:
    """Physical position: XY in CAD mm and elevation in metres."""
    id: str
    x_mm: float
    y_mm: float
    z_m: float
    role: str = 'base'
    source_node: int | None = None
    head_index: int | None = None
    head_kind: str | None = None


@dataclass(frozen=True)
class ReviewPipe:
    """An undirected physical link; output orientation is a solver reference."""
    id: str
    start: str
    end: str
    length_m: float
    source_edge: tuple[int, int] | None = None


@dataclass
class ReviewNetwork:
    """Inspectable intermediate model; no fabricated solved flow or head load."""
    nodes: dict[str, ReviewNode] = field(default_factory=dict)
    pipes: dict[str, ReviewPipe] = field(default_factory=dict)
    cycle_rank: int = 0
    source_revision: str = ''
    source_aliases: dict[int,int] = field(default_factory=dict)
    arc_ports: dict[str,list] = field(default_factory=dict)
    arc_junctions: dict[str,dict] = field(default_factory=dict)
    arc_report: dict = field(default_factory=dict)

    def validate(self) -> None:
        """Reject topology loss, invalid dimensions or a missing supply."""
        graph = nx.MultiGraph()
        graph.add_nodes_from(self.nodes)
        for pipe in self.pipes.values():
            if pipe.start not in self.nodes or pipe.end not in self.nodes or pipe.start == pipe.end:
                raise ValueError(f"배관 {pipe.id}의 절점 참조가 올바르지 않습니다.")
            if not math.isfinite(pipe.length_m) or pipe.length_m <= 0:
                raise ValueError(f"배관 {pipe.id}의 길이가 올바르지 않습니다.")
            rise = self.nodes[pipe.end].z_m - self.nodes[pipe.start].z_m
            if abs(rise) > pipe.length_m + 1e-7:
                raise ValueError(f"배관 {pipe.id}의 표고차가 길이보다 큽니다.")
            graph.add_edge(pipe.start, pipe.end)
        if not graph or not nx.is_connected(graph):
            raise ValueError("변환망이 급수원에 연결된 한 덩이가 아닙니다.")
        if sum(n.role == 'pump' for n in self.nodes.values()) != 1:
            raise ValueError("변환망에는 선택한 급수원 하나가 필요합니다.")
        if graph.number_of_edges()-graph.number_of_nodes()+1 != self.cycle_rank:
            raise ValueError("변환 과정에서 루프 연결 수가 바뀌었습니다.")
        for n in self.nodes.values():
            if not all(math.isfinite(v) for v in (n.x_mm,n.y_mm,n.z_m)):
                raise ValueError("변환망 좌표가 유한한 값이 아닙니다.")
            if n.role == 'head' and graph.degree(n.id) != 1:
                raise ValueError("헤드 노즐이 배관의 통과 절점에 겹쳐 있습니다.")


def build_review_network(network: FlowNetwork, points: Sequence[Sequence[float]],
                         heads: Sequence[int], kinds: Mapping[int, str], *,
                         datum_m: float, upright_m: float, pendant_rise_m: float,
                         pendant_drop_m: float, combo_rise_m: float,
                         combo_up_m: float, combo_drop_m: float,
                         combo_arm_m: float = 0.5,
                         arcs: Sequence[Mapping] = (), branch_rise_m: float = 0.5,
                         drawn_edges: Sequence[tuple[int,int]] = (),
                         protected_nodes: Sequence[int] = (),
                         source_edges: Sequence[tuple[int, int]] | None = None) -> ReviewNetwork:
    """Retain every selected-path edge and graft configured head stems in 3D."""
    values = (datum_m,upright_m,pendant_rise_m,pendant_drop_m,combo_rise_m,combo_up_m,combo_drop_m,combo_arm_m)
    if not all(math.isfinite(v) for v in values) or any(v < 0 for v in values[1:]):
        raise ValueError("표고·헤드 접속관 치수는 유한한 값이어야 하며 길이는 음수일 수 없습니다.")
    ref = network.reference
    selected = sorted({ref.representatives.get(int(h), int(h)) for h in heads})
    if not selected or any(h not in ref.head_node for h in selected):
        raise ValueError("선정 헤드가 급수원에 연결되지 않았습니다. 다시 선정하세요.")
    edges = network.resolve_selection(selected, source_edges=source_edges)
    used = {n for e in edges for n in e}
    graph = ReviewNetwork(cycle_rank=len(edges)-len(used)+1, source_revision=network.revision)
    root = ref.roots[0]
    aliases = nx.utils.UnionFind(used)
    for a,b in edges:
        if ref.lengths_mm[edge_key(a,b)] <= 1e-6 and math.dist(points[a][:2],points[b][:2])<=1e-6:
            aliases.union(a,b)
    for members in aliases.to_sets():
        representative=root if root in members else min(members)
        graph.source_aliases.update({n:representative for n in members})
    for n in sorted(set(graph.source_aliases.values())):
        graph.nodes[f'N{n}'] = ReviewNode(f'N{n}', float(points[n][0]), float(points[n][1]),
                                         datum_m, 'pump' if n == root else 'base', n)
    for a,b in sorted(edges):
        source=edge_key(a,b)
        a,b=graph.source_aliases[a],graph.source_aliases[b]
        if a==b:
            if ref.lengths_mm[source]>1e-6:
                raise ValueError('동일 절점에 돌아오는 유효 길이 배관을 별도 검토해야 합니다.')
            continue
        a,b = sorted((a,b), key=lambda n:(ref.distance_mm.get(n,math.inf),n))
        pid=f'P{len(graph.pipes)+1}'
        graph.pipes[pid] = ReviewPipe(pid, f'N{a}', f'N{b}',ref.lengths_mm[source] / 1000.0,source)
    from .review_elevation import apply_arc_elevations
    # Even an unselected head/source must never be mistaken for symbol ink.
    protected=set(protected_nodes)|set(ref.head_node.values())|set(ref.roots)
    apply_arc_elevations(graph,points,list(ref.lengths_mm),arcs,branch_rise_m,
                        drawn_edges=drawn_edges,protected_nodes=protected)

    def stem(start, z, h, kind, role, offset=(0.0,0.0)):
        parent = graph.nodes[start]
        # A zero rise is an omitted intermediate segment, not a phantom pipe.
        if role != 'head' and abs(z-parent.z_m) < 1e-10 and offset == (0.0,0.0):
            return start
        if role == 'head' and abs(z-parent.z_m) < 1e-10 and offset == (0.0,0.0):
            raise ValueError('헤드 접속관 길이는 0보다 커야 합니다. 헤드 치수를 확인하세요.')
        nid = f'H{h}_{len(graph.nodes)}'
        graph.nodes[nid] = ReviewNode(nid,parent.x_mm+offset[0]*1000,parent.y_mm+offset[1]*1000,z,role,None,h,kind)
        pid = f'P{len(graph.pipes)+1}'
        while pid in graph.pipes:
            pid = f'P{int(pid[1:])+1}'
        graph.pipes[pid] = ReviewPipe(pid,start,nid,math.sqrt((z-parent.z_m)**2+offset[0]**2+offset[1]**2))
        return nid
    for h in selected:
        node = f'N{graph.source_aliases[ref.head_node[h]]}'
        local_datum = graph.nodes[node].z_m
        kind = kinds.get(h)
        if kind == '상향식':
            stem(node,local_datum+upright_m,h,kind,'head')
        elif kind in ('하향식','상하향식'):
            rise = pendant_rise_m if kind == '하향식' else combo_rise_m
            top = stem(node,local_datum+rise,h,kind,'base')
            if kind == '상하향식':
                stem(top,local_datum+rise+combo_up_m,h,'상향식','head')
                # A reference direction along a real connected pipe, not an
                # assumed hydraulic upstream direction. Keep all loop ports.
                anchor = ref.head_node[h]
                neighbors = sorted(b if a==anchor else a for a,b in edges if anchor in (a,b))
                neighbor = next((n for n in neighbors if math.dist(points[n][:2],points[anchor][:2])>1e-7),None)
                if neighbor is None:
                    raise ValueError('상하향 헤드 수평 팔의 기준 방향을 찾지 못했습니다.')
                dx,dy = (points[anchor][i]-points[neighbor][i] for i in range(2))
                scale = combo_arm_m/math.hypot(dx,dy)
                top = stem(top,local_datum+rise,h,kind,'base',(dx*scale,dy*scale))
            stem(top,local_datum+rise-(pendant_drop_m if kind == '하향식' else combo_drop_m),h,'하향식','head')
        else:
            raise ValueError(f"헤드 {h}의 접속 형식 {kind!r}은 확정되지 않았습니다.")
    graph.validate()
    return graph
