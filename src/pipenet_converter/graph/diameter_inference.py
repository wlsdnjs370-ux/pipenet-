"""Conservative annotation ownership on a supplied, full source pipe graph.

Coordinates, distances and nominal diameters are mm. This module neither sizes
pipes hydraulically nor changes geometry. Connected straight runs are candidates
for propagation, never a license to copy a size onto parallel/disconnected runs.
"""
from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import asdict, dataclass, field
import math

from .flow import Edge, FlowTree, edge_key


@dataclass(frozen=True)
class DiameterAnnotation:
    """A mapped drawing annotation; rotation None means genuinely unavailable."""

    identity: str
    x: float
    y: float
    nominal_mm: int
    rotation_deg: float | None = None
    height_mm: float = 0.0
    layer: str = ""
    raw_text: str = ""
    entity_handle: str = ""
    native_anchor_mm: tuple[float, float] | None = None
    leader_xy_mm: tuple[float, float] | None = None
    leader_ambiguous: bool = False
    source_ambiguous: bool = False

    def __post_init__(self) -> None:
        values = [self.x, self.y, self.nominal_mm, self.height_mm]
        if self.rotation_deg is not None:
            values.append(self.rotation_deg)
        for point in (self.native_anchor_mm, self.leader_xy_mm):
            if point is not None:
                if len(point) != 2:
                    raise ValueError('관경 문자 기준점은 XY 좌표여야 합니다.')
                values.extend(point)
        if not all(math.isfinite(v) for v in values) or self.nominal_mm <= 0 or self.height_mm < 0:
            raise ValueError("관경 문자 좌표·치수·방향이 올바르지 않습니다.")


@dataclass(frozen=True)
class InferenceConfig:
    """Drawing-independent tolerances, overridden by a project config file."""

    search_radius_mm: float = 1500.0
    parallel_tolerance_deg: float = 20.0
    continuation_tolerance_deg: float = 10.0
    ownership_margin_mm: float = 100.0
    max_propagation_mm: float = 12000.0
    leader_snap_mm: float = 25.0

    def __post_init__(self) -> None:
        for key, value in asdict(self).items():
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"관경 추론 설정 {key}는 양의 유한값이어야 합니다.")
        if max(self.parallel_tolerance_deg, self.continuation_tolerance_deg) >= 45:
            raise ValueError("관경 추론 방향 허용각은 45도 미만이어야 합니다.")


@dataclass
class AnnotationMatchAudit:
    """Inspectable ownership decision for one original annotation identity."""

    annotation_id: str
    status: str
    basis: str
    candidate_edges: list[Edge]
    owner_edge: Edge | None = None


@dataclass
class DiameterContext:
    """Per-original-edge evidence plus the full, unpruned water-flow basis."""

    flow: FlowTree
    decisions: dict[Edge, dict]
    summary: dict = field(default_factory=dict)
    annotation_audit: list[AnnotationMatchAudit] = field(default_factory=list, kw_only=True)

    def for_path(self, reference: Sequence[int]) -> dict:
        """Resolve an expanded pipe's original path without inventing a reducer."""
        path = self.flow.between(int(reference[0]), int(reference[1]))
        rows = [self.decisions.get(e, {}) for e in path]
        if not path or not set(path) <= self.flow.edges:
            return {"method": "위상 대응 없음", "reason": "no_source_path",
                    "review_reasons": ["원본 물흐름 경로를 대조하지 못했습니다."]}
        values = {r['text_mm'] for r in rows if r.get('text_mm') is not None}
        warnings = sorted({w for r in rows for w in r.get('review_reasons', [])})
        if len(values) > 1:
            return {"method": "관경 변화 구간 충돌", "reason": "diameter_boundary",
                    "candidate_mm": sorted(values), "block_export": True,
                    "review_reasons": warnings + ["한 계산 배관에 서로 다른 관경 표기가 있습니다. 배관 분할 또는 관경 확인이 필요합니다."]}
        if any(r.get('block_export') for r in rows):
            return {"method": "표기 소속 충돌", "reason": "ambiguous_annotation",
                    "candidate_mm": sorted({d for r in rows for d in r.get('candidate_mm', [])}),
                    "block_export": True, "review_reasons": warnings}
        chosen = next((r for r in rows if r.get('text_mm') is not None), rows[0])
        known_loads = [self.flow.loads[e] for e in path if self.flow.loads[e] is not None]
        out = dict(chosen, review_reasons=warnings, full_head_count=max(known_loads) if known_loads else None)
        rejected = [r['excluded_nearest'] for r in rows if r.get('excluded_nearest')]
        if rejected:
            out['excluded_nearest'] = min(rejected,key=lambda r:r['distance_mm'])
        if any(r.get('inferred') for r in rows):
            out.update(inferred=True,method="동일 관로 전파 · 추론")
        if values and any(r.get('text_mm') is None for r in rows):
            out['partial_text_mm'] = sorted(values)
            for key in ('text_mm','text_xy_mm','text_rotation_deg','text_layer','text_raw','text_entity_handle',
                        'distance_mm','annotation_id','owner_edge','distance_basis'):
                out.pop(key,None)
            out.update(method="원본 경로 일부 근거 없음", reason="partial_path",
                       review_reasons=warnings + ["병합 배관의 일부 구간만 표기에 대응합니다. 전체 관경을 확인하세요."])
        return out


def _angle(a: float, b: float) -> float:
    return abs((a - b + 90) % 180 - 90)


def _distance(t: DiameterAnnotation, a: Sequence[float], b: Sequence[float]) -> float:
    dx, dy = b[0]-a[0], b[1]-a[1]
    u = max(0.0, min(1.0, ((t.x-a[0])*dx+(t.y-a[1])*dy)/(dx*dx+dy*dy)))
    return math.hypot(t.x-a[0]-u*dx, t.y-a[1]-u*dy)


def infer_diameter_annotations(
    points: Sequence[Sequence[float]], physical_edges: Iterable[Edge], flow: FlowTree,
    annotations: Sequence[DiameterAnnotation], *, barriers: Iterable[int] = (),
    config: InferenceConfig | None = None,
    use_native_anchors: bool = False,
) -> DiameterContext:
    """Assign each text to at most one run; propagate only connected evidence.

Text direction is supporting evidence, not a hard law: sole perpendicular
matches remain explicit review items. Ambiguous ownership is not tie-broken by
node IDs, pipe size, or head count. Full loads are used for role/consistency
checks, not to claim a nominal diameter is hydraulically correct.
"""
    cfg = config or InferenceConfig()
    active_edges = flow.edges
    edges = sorted({edge_key(*e) for e in physical_edges if e[0] != e[1]})
    if any(not (0 <= a < len(points) and 0 <= b < len(points)) for a,b in edges):
        raise ValueError("관경 추론 원본 노드 참조가 올바르지 않습니다.")
    if any(not all(math.isfinite(v) for v in points[n][:2]) for e in edges for n in e):
        raise ValueError("관경 추론 원본 좌표는 유한한 mm 값이어야 합니다.")
    edges = [e for e in edges if math.dist(points[e[0]][:2], points[e[1]][:2]) > 1e-6]
    adj: dict[int, list[Edge]] = defaultdict(list)
    parent = {e:e for e in edges}
    angles = {}
    for a,b in edges:
        adj[a].append((a,b)); adj[b].append((a,b))
        angles[a,b] = math.degrees(math.atan2(points[b][1]-points[a][1], points[b][0]-points[a][0]))

    def find(e: Edge) -> Edge:
        while parent[e] != e:
            parent[e] = parent[parent[e]]; e = parent[e]
        return e

    # An inline head tap does not end the through-run. Only the terminal head
    # arm ends naturally; stopping at every head binding fragments the main.
    stops = set(barriers) | set(flow.roots)
    for node, incident in adj.items():
        if node in stops:
            continue
        compatible: dict[Edge, list[Edge]] = defaultdict(list)
        for i,e in enumerate(incident):
            other = e[0] if e[1] == node else e[1]
            dx,dy = points[other][0]-points[node][0], points[other][1]-points[node][1]
            for f in incident[i+1:]:
                end = f[0] if f[1] == node else f[1]
                ex,ey = points[end][0]-points[node][0], points[end][1]-points[node][1]
                dot = (dx*ex+dy*ey)/(math.hypot(dx,dy)*math.hypot(ex,ey))
                if (e in active_edges) != (f in active_edges):
                    continue
                if dot < -math.cos(math.radians(cfg.continuation_tolerance_deg)):
                    compatible[e].append(f); compatible[f].append(e)
        for e,others in compatible.items():
            if len(others) == 1 and len(compatible[others[0]]) == 1:
                x,y = find(e),find(others[0]); parent[max(x,y)] = min(x,y)
    group = {e:find(e) for e in edges}
    runs: dict[Edge,list[Edge]] = defaultdict(list)
    for e in edges:
        runs[group[e]].append(e)

    # Spatial bins avoid the full drawing-text × pipe product on large plans.
    radius = cfg.search_radius_mm
    bins: dict[tuple[int,int], list[Edge]] = defaultdict(list)
    for e in edges:
        a,b = (points[n] for n in e)
        for x in range(math.floor((min(a[0],b[0])-radius)/radius), math.floor((max(a[0],b[0])+radius)/radius)+1):
            for y in range(math.floor((min(a[1],b[1])-radius)/radius), math.floor((max(a[1],b[1])+radius)/radius)+1):
                bins[x,y].append(e)
    owned: dict[Edge,list[tuple[DiameterAnnotation,Edge,float,list[str]]]] = defaultdict(list)
    ambiguous: dict[Edge,set[int]] = defaultdict(set)
    near: dict[Edge,list[tuple[float,DiameterAnnotation,Edge | None]]] = defaultdict(list)
    ownership_count = 0
    audit = []
    def location(t: DiameterAnnotation) -> tuple[float, float]:
        return (t.leader_xy_mm or t.native_anchor_mm or (t.x,t.y)) if use_native_anchors else (t.x,t.y)

    def distance_basis(t: DiameterAnnotation) -> str:
        if use_native_anchors and t.leader_xy_mm:
            return '명시 연결된 DXF 지시선 끝점'
        if use_native_anchors and t.native_anchor_mm:
            return '원본 문자 범위 중심과 대응 원본 선분'
        return '원문자와 대응 원본 선분'

    for text in annotations:
        candidates = []
        x,y = location(text)
        leader = use_native_anchors and text.leader_xy_mm is not None
        probe = DiameterAnnotation(text.identity,x,y,text.nominal_mm)
        search = cfg.leader_snap_mm if leader else radius
        for e in bins.get((math.floor(x/radius), math.floor(y/radius)), []):
            distance = _distance(probe, points[e[0]], points[e[1]])
            if distance < search:
                angle = _angle(text.rotation_deg,angles[e]) if text.rotation_deg is not None and not leader else None
                candidates.append((distance,e,angle))
        aligned = [c for c in candidates if c[2] is not None and c[2] <= cfg.parallel_tolerance_deg]
        eligible = aligned or candidates
        by_run = {}
        for c in sorted(eligible):
            by_run.setdefault(group[c[1]], c)
        best = sorted(by_run.values())
        owner = None
        margin = min(cfg.ownership_margin_mm, search*.25) if leader else max(cfg.ownership_margin_mm, text.height_mm*.25)
        if best and not (use_native_anchors and (text.leader_ambiguous or text.source_ambiguous)) and (len(best) == 1 or best[1][0]-best[0][0] > margin):
            distance,e,angle = best[0]; owner=group[e]
            warnings = []
            if angle is None and not leader:
                warnings.append("문자 방향 정보가 없어 관로 소속을 확인해야 합니다.")
            elif not aligned and not leader:
                warnings.append("문자 방향과 배관 방향이 다릅니다. 리더선 또는 표기 관행을 확인하세요.")
            owned[owner].append((text,e,distance,warnings)); ownership_count += 1
        elif best:
            for distance,e,_angle_value in best:
                if distance-best[0][0] <= margin:
                    ambiguous[group[e]].add(text.nominal_mm)
        for distance,e,_ in candidates:
            near[e].append((distance,text,owner))
        audit.append(AnnotationMatchAudit(text.identity,
            'matched' if owner is not None else 'ambiguous' if best else 'unmatched',
            distance_basis(text), [c[1] for c in best], best[0][1] if owner is not None else None))

    decisions = {}
    for e in active_edges:
        if e not in group:
            decisions[e] = {"reason":"no_geometry", "method":"위상 기하 없음"}
            continue
        run = group[e]
        ax = math.radians(angles[run]); ux,uy = math.cos(ax), math.sin(ax)
        def position(x: float,y: float) -> float:
            return x*ux+y*uy
        center = sum(position(*points[n][:2]) for n in e)/2
        records = owned[run]
        direct = [r for r in records if r[1] == e]
        chosen = None; inferred = False; warnings = []
        conflict = {r[0].nominal_mm for r in direct}
        if len(conflict) == 1:
            chosen = min(direct,key=lambda r:r[2])
        elif not direct and records:
            left = [r for r in records if position(*location(r[0])) <= center]
            right = [r for r in records if position(*location(r[0])) > center]
            lo = max(left,key=lambda r:position(*location(r[0]))) if left else None
            hi = min(right,key=lambda r:position(*location(r[0]))) if right else None
            if lo and hi and lo[0].nominal_mm != hi[0].nominal_mm:
                conflict = {lo[0].nominal_mm,hi[0].nominal_mm}
            else:
                nearest = min(records,key=lambda r:abs(position(*location(r[0]))-center))
                if abs(position(*location(nearest[0]))-center) <= cfg.max_propagation_mm:
                    chosen = nearest; inferred = True
                    warnings.append("같은 직선 관로의 표기를 전파한 추론값입니다. 적용 범위를 확인하세요.")
        result = {"method":"위상 대조 · 표기 없음", "reason":"no_match",
                  "run":list(run), "full_head_count":flow.loads[e], "review_reasons":warnings}
        if len(conflict) > 1:
            result.update(method="관경 변화 구간 충돌",reason="diameter_boundary",candidate_mm=sorted(conflict),block_export=True)
            warnings.append("서로 다른 관경 사이의 경계가 확정되지 않았습니다. 구간 분할 또는 관경 확인이 필요합니다.")
        elif chosen:
            text,owner_edge,distance,extra = chosen
            warnings.extend(extra)
            result.update(text_mm=text.nominal_mm,text_xy_mm=[text.x,text.y],
                          text_raw=text.raw_text,text_entity_handle=text.entity_handle,
                          text_rotation_deg=text.rotation_deg,text_layer=text.layer,
                          distance_mm=distance,annotation_id=text.identity,owner_edge=list(owner_edge),
                          inferred=inferred,method="동일 관로 전파 · 추론" if inferred else "위상 · 문자 방향 대응",
                          reason="matched",distance_basis=distance_basis(text))
            if use_native_anchors:
                result['matching_point_mm'] = list(location(text))
                result['leader_xy_mm'] = list(text.leader_xy_mm) if text.leader_xy_mm else None
        elif ambiguous[run]:
            warnings.append("가까운 여러 관로 중 문자의 소속을 구분하지 못했습니다.")
            result.update(reason="ambiguous_annotation",candidate_mm=sorted(ambiguous[run]),block_export=True)
        nearest = min(near[e],key=lambda r:r[0],default=None)
        if nearest and nearest[2] is not None and nearest[2] != run:
            result['excluded_nearest'] = {"text_mm":nearest[1].nominal_mm,
                "distance_mm":nearest[0],"reason":"다른 관로의 표기로 분리", "text_xy_mm":[nearest[1].x,nearest[1].y]}
        decisions[e] = result
    # A size inversion is a review signal, not an unconditional enlargement.
    for node,upstream in flow.parent.items():
        before = edge_key(node,upstream)
        if before not in decisions or node in stops:
            continue
        for after in adj[node]:
            if after == before or after not in decisions:
                continue
            a,b = decisions[before],decisions[after]
            if a.get('text_mm') is not None and b.get('text_mm') is not None and a['text_mm'] < b['text_mm']:
                message=f"상류 표기 {a['text_mm']} → 하류 표기 {b['text_mm']} mm: 관경 증가의 설계 근거를 확인하세요."
                for row in (a,b):
                    row['review_reasons'].append(message)
    return DiameterContext(flow,decisions,{"annotations":len(annotations),"owned":ownership_count,
        "runs":len(runs),"flow_revision":flow.revision,"full_heads":len(set(flow.representatives.values())),
        "ambiguous_annotations":sum(a.status=='ambiguous' for a in audit),
        "unmatched_annotations":sum(a.status=='unmatched' for a in audit)}, annotation_audit=audit)


@dataclass
class AnnotationGraph:
    """Undirected annotation basis: no parents, no per-pipe head-count fiction.

    Conversion currently preserves individual source edges; hence a lookup is
    direct, never the arbitrary spanning-tree detour around a cycle.
    """
    edges: frozenset[Edge]
    roots: tuple[int, ...]
    representatives: dict[int, int]
    revision: str
    parent: dict = field(default_factory=dict, init=False)

    loads: dict = field(default_factory=dict, init=False)

    def __post_init__(self) -> None:
        self.loads = dict.fromkeys(self.edges)

    def between(self, a: int, b: int) -> list[Edge]:
        """Only the identified physical edge; no flow direction is implied."""
        edge = edge_key(a,b)
        return [edge] if edge in self.edges else []
