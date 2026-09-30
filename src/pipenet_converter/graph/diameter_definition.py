"""Drawing-first diameter definitions for the opt-in H workflow.

Real plan coordinates and nominal sizes are millimetres. Unknown is ``None``;
the legacy table adapter alone uses 0 as an explicitly blocked draft sentinel.
No hydraulic sizing, geometry edits, OCR or drawing-specific layer rules live here.
"""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from dataclasses import asdict, dataclass
import hashlib
import json
import math
from typing import Any, Iterable, Sequence

from .diameter_inference import (
    AnnotationGraph, DiameterAnnotation, DiameterContext, InferenceConfig,
    infer_diameter_annotations,
)
from .flow import Edge, edge_key
from ..progress import report

POLICY = "drawing_first_v1"


@dataclass(frozen=True)
class DefinitionConfig:
    """Conservative structural correspondence and bounded work limits."""

    straight_tolerance_deg: float = 10.0
    repeat_length_tolerance: float = .15
    repeat_shape_tolerance: float = .08
    max_branch_edges: int = 500

    def __post_init__(self) -> None:
        if not math.isfinite(self.straight_tolerance_deg) or not 0 < self.straight_tolerance_deg < 45:
            raise ValueError("연속 관로 각도 허용치는 0~45도 사이여야 합니다.")
        if not math.isfinite(self.repeat_length_tolerance) or not 0 <= self.repeat_length_tolerance <= .5:
            raise ValueError("반복 구조 길이 허용치는 0~0.5 사이여야 합니다.")
        if not math.isfinite(self.repeat_shape_tolerance) or not 0 <= self.repeat_shape_tolerance <= .25:
            raise ValueError("반복 구간 모양 허용치는 0~0.25 사이여야 합니다.")
        if not isinstance(self.max_branch_edges, int) or self.max_branch_edges < 1:
            raise ValueError("반복 경로 탐색 한도는 양의 정수여야 합니다.")


@dataclass
class DefinitionContext(DiameterContext):
    """Complete drawing evidence, with original traversal retained for expansion."""

    points: Sequence[Sequence[float]] = ()

    def for_path(self, reference: Sequence[int]) -> dict:
        """Keep every contributing original edge and its annotation provenance."""
        path = self.flow.between(int(reference[0]), int(reference[1]))
        rows = [self.decisions.get(e, {}) for e in path]
        if not rows:
            return {"reason": "no_source_path", "block_export": True}
        values = {r.get("text_mm") for r in rows}
        if len(values) != 1 or any(r.get("block_export") for r in rows):
            return {"reason": "partial_or_conflicting_path", "block_export": True,
                    "method": "구간별 관경 확인 필요", "source_edges": [list(e) for e in path],
                    "evidence_chain": [deepcopy(r) for r in rows],
                    "candidate_mm": sorted(v for v in values if v),
                    "review_reasons": ["전개 배관의 원본 구간 관경이 서로 다르거나 미지정입니다."]}
        result = deepcopy(rows[0])
        result["source_edges"] = [list(e) for e in path]
        result["evidence_chain"] = [deepcopy(r) for r in rows]
        return result


def graph_revision(points: Sequence[Sequence[float]], edges: Iterable[Edge],
                   heads: dict[int, Any], roots: Iterable[int], extra: Any = None) -> str:
    """Fingerprint real identities and settings, never BFS display labels."""
    body = [points, sorted(map(list, edges)), sorted(heads.items()), sorted(roots), extra]
    return hashlib.sha256(json.dumps(body, ensure_ascii=False, sort_keys=True,
                                    allow_nan=False, default=str).encode()).hexdigest()


def _adjacency(edges: Iterable[Edge]) -> dict[int, list[Edge]]:
    adj: dict[int, list[Edge]] = defaultdict(list)
    for e in sorted(edges):
        for node in e:
            adj[node].append(e)
    return dict(adj)


def _other(edge: Edge, node: int) -> int:
    return edge[1] if node == edge[0] else edge[0]


def _through(points: Sequence[Sequence[float]], adj: dict[int, list[Edge]],
             stops: set[int], tolerance: float) -> dict[tuple[int, Edge], Edge]:
    """Only unique opposite ports continue through a tee; degree-2 bends continue."""
    result = {}
    for node, edges in adj.items():
        if node in stops:
            continue
        if len(edges) == 2:
            result[node, edges[0]] = edges[1]
            result[node, edges[1]] = edges[0]
            continue
        candidates: dict[Edge, list[Edge]] = defaultdict(list)
        for i, e in enumerate(edges):
            a = points[_other(e, node)]
            u = (a[0]-points[node][0], a[1]-points[node][1])
            for f in edges[i+1:]:
                b = points[_other(f, node)]
                v = (b[0]-points[node][0], b[1]-points[node][1])
                norm = math.hypot(*u)*math.hypot(*v)
                if norm and sum(x*y for x, y in zip(u, v))/norm < -math.cos(math.radians(tolerance)):
                    candidates[e].append(f)
                    candidates[f].append(e)
        for e, options in candidates.items():
            if len(options) == 1 and len(candidates[options[0]]) == 1:
                result[node, e] = options[0]
    return result


def _runs(edges: set[Edge], through: dict) -> list[list[Edge]]:
    pending, result = set(edges), []
    while pending:
        start = min(pending)
        component, todo = [], [start]
        while todo:
            e = todo.pop()
            if e not in pending:
                continue
            pending.remove(e)
            component.append(e)
            todo.extend(through[n, e] for n in e if (n, e) in through)
        result.append(sorted(component))
    return result


def ordered_path(edges: Iterable[Edge], anchor: int | None = None) -> tuple[list[int], list[Edge]]:
    """A user-selected path must be connected, non-branching and explicitly oriented."""
    edges = {edge_key(*e) for e in edges}
    adj = _adjacency(edges)
    ends = sorted(n for n, incident in adj.items() if len(incident) == 1)
    if not edges or len(ends) != 2 or any(len(v) > 2 for v in adj.values()):
        raise ValueError("분기 없는 한 경로를 선택하세요. 추가 가지는 따로 지정할 수 있습니다.")
    if anchor is not None and anchor not in ends:
        raise ValueError("기준점은 선택 경로의 끝점이어야 합니다.")
    nodes, path, node = [ends[0] if anchor is None else anchor], [], ends[0] if anchor is None else anchor
    used = set()
    while True:
        next_edges = [e for e in adj[node] if e not in used]
        if not next_edges:
            break
        e = next_edges[0]
        path.append(e)
        used.add(e)
        node = _other(e, node)
        nodes.append(node)
    if used != edges:
        raise ValueError("서로 끊어진 배관을 하나의 경로로 붙여넣을 수 없습니다.")
    return nodes, path


def _branches(adj: dict[int, list[Edge]], through: dict, stops: set[int],
              limit: int) -> list[tuple[list[int], list[Edge]]]:
    """Find tee-side chains; a second main connection is a different endpoint role."""
    found, seen = [], set()
    for start, incident in sorted(adj.items()):
        if len(incident) < 3 or start in stops:
            continue
        for e in incident:
            if (start, e) in through:
                continue
            nodes, path, current = [start], [], e
            while len(path) < limit:
                path.append(current)
                node = _other(current, nodes[-1])
                nodes.append(node)
                if node in stops or len(adj[node]) != 2 or node in nodes[:-1]:
                    break
                current = next(edge for edge in adj[node] if edge != current)
            identity = tuple(sorted(path))
            if identity not in seen and len(path) < limit and nodes[-1] not in nodes[:-1]:
                seen.add(identity)
                found.append((nodes, path))
    return found


def _slots(nodes: list[int], path: list[Edge], heads: dict[int, Any]) -> list[list[Edge]]:
    """CAD split fragments between head/port landmarks form one correspondence slot."""
    slots, current = [], []
    for i, e in enumerate(path):
        current.append(e)
        if nodes[i+1] in heads or i == len(path)-1:
            slots.append(current)
            current = []
    return slots


def define_diameters(points: Sequence[Sequence[float]], edges: Iterable[Edge], flow: Any,
                     annotations: Sequence[DiameterAnnotation], *,
                     heads: dict[int, Any] | None = None, barriers: Iterable[int] = (),
                     inference: InferenceConfig | None = None,
                     config: DefinitionConfig | None = None) -> DefinitionContext:
    """Read text, trace physical runs, then transfer only consistent representative patterns.

    Ambiguous ownership, competing representatives, and unknown boundaries remain
    unresolved. Representative search covers the supplied full drawing graph,
    not just the selected operating heads. Manual values are applied later.
    """
    cfg, heads = config or DefinitionConfig(), heads or {}
    barriers = tuple(barriers)
    edges = {edge_key(*e) for e in edges if e[0] != e[1]}
    basis = AnnotationGraph(frozenset(edges), tuple(flow.roots), flow.representatives, flow.revision)
    report("diameter", "도면 관경 표기 읽기", reset=True, policy=POLICY,
           annotations=[asdict(a) for a in annotations])
    # Head-to-head slots are real interpretation boundaries. They may have
    # different diameters; one nearby label must not flood the entire branch.
    original = infer_diameter_annotations(points, edges, basis, annotations,
                                          barriers=set(barriers) | set(heads), config=inference,
                                          use_native_anchors=True)
    rows = deepcopy(original.decisions)
    direct = {e: r for e, r in rows.items() if r.get("text_mm") and not r.get("inferred")}
    for e, r in direct.items():
        r.update(definition_source="drawing_direct", policy=POLICY)
        report("diameter", "원본 표기 대응", edge=list(e), xy=[list(points[n][:2]) for n in e], evidence=r)
    adj = _adjacency(edges)
    stops = set(barriers) | set(flow.roots)
    through = _through(points, adj, stops, cfg.straight_tolerance_deg)
    propagation = _through(points, adj, stops | set(heads), cfg.straight_tolerance_deg)
    runs = _runs({e for e in edges if not rows[e].get('block_export')}, propagation)
    for run in runs:
        known = [(e, direct[e]) for e in run if e in direct]
        if not known:
            continue
        values = {r["text_mm"] for _, r in known}
        # Multiple sizes on one run are NOT evidence that the whole run is uniform.
        # Existing local matches are kept; unknown boundaries are left for review.
        if len(values) != 1:
            continue
        for e in run:
            row = rows[e]
            if e in direct or row.get("block_export"):
                continue
            midpoint = tuple((points[e[0]][i]+points[e[1]][i])/2 for i in (0,1))
            origin, source = min(known,key=lambda pair:(
                math.dist(pair[1].get('matching_point_mm') or pair[1]['text_xy_mm'],midpoint),
                tuple(pair[1]['text_xy_mm']),str(pair[1]['annotation_id'])))
            row.update(deepcopy(source), inferred=True, definition_source="drawing_run",
                       method="연속 관로 대표 표기", source_edge=list(origin), policy=POLICY,
                       review_reasons=["주관 연결·꺾임을 따라 표기를 참조했습니다."])
            report("diameter", "연속 관로 범위", edge=list(e), xy=[list(points[n][:2]) for n in e], evidence=row)
    # A repeat has matching ordered head kinds and endpoint roles, not merely
    # a visually similar bounding box or a count of CAD fragments.
    branches = _branches(adj, through, stops, cfg.max_branch_edges)
    from .diameter_repetition import transfer_repeated_slots
    repetition = transfer_repeated_slots(points, edges, branches, heads, stops, rows, cfg, POLICY)
    for e, row in rows.items():
        row.setdefault("policy", POLICY)
        if row.get("text_mm") and "definition_source" not in row:
            row["definition_source"] = "drawing_run" if row.get("inferred") else "drawing_direct"
        row.setdefault("source_edge", list(e))
    summary = dict(original.summary, policy=POLICY, algorithm='segment_reference_v2', **repetition,
                   continuous_runs=len(runs), full_drawing_edges=len(edges),
                   unresolved=sum(not r.get("text_mm") or bool(r.get("block_export")) for r in rows.values()))
    report("diameter", "관경 해석 완료", summary=summary, complete=True)
    return DefinitionContext(flow, rows, summary, points, annotation_audit=original.annotation_audit)


def resolve_pipe_definitions(net: dict, edge_ref: dict, context: DefinitionContext,
                             loads: dict, *, tree: bool, rule: Any,
                             overrides: dict | None = None) -> tuple[dict, dict, dict]:
    """Drawing evidence first, unknown stays unknown, explicit user overrides last."""
    values, evidence, changed = {}, {}, {}
    for pid, pipe in net["pipe_data"].items():
        ref = edge_ref.get(pid)
        rec = context.for_path(ref) if ref is not None else {"reason": "generated_pipe"}
        rec = deepcopy(rec)
        text = rec.get("text_mm") if not rec.get("block_export") else None
        count = net.get("physical_pipe_loads", {}).get(pid, loads.get(pid))
        if count is None and ref is not None:
            path = context.flow.between(*ref)
            known = [context.flow.loads.get(e) for e in path if context.flow.loads.get(e) is not None]
            count = max(known) if known else None
        rule_mm = rule(int(count)) if tree and count is not None else None
        if text is not None:
            dia, src = int(text), "text"
        else:
            dia, src = 0, "unresolved"
        rec.update(policy=POLICY, version=4, source=src, auto_mm=dia or None,
                   rule_mm=rule_mm, head_count=count if tree else None,
                   block_export=dia == 0, review_only=not tree)
        if text and rule_mm and text < rule_mm:
            rec.setdefault("review_reasons", []).append(f"도면 {text}A 유지 · 기존 기준개수 규칙 {rule_mm}A와 차이")
        key = tuple(sorted(ref)) if ref is not None else None
        override = (overrides or {}).get(key)
        if override is None and ref is not None:
            path = context.flow.between(*ref)
            edits = [(overrides or {}).get(e) for e in path]
            if edits and all(edits) and len({v[0] for v in edits}) == 1:
                override = edits[0]
            elif any(edits):
                dia, src = 0, "unresolved"
                rec.update(block_export=True, reason="partial_manual_path",
                           review_reasons=["하나로 전개된 배관의 일부만 수정되었습니다. 구간 전체를 선택하거나 관경 변경점을 분리하세요."])
        if override:
            new = int(override[0])
            rec["manual"] = {"previous_mm": dia or None, "value_mm": new, "note": str(override[1])}
            rec.update(source="user", block_export=False)
            changed[pid] = dict(dia=new,note=str(override[1]),orig_dia=dia,orig_src=src,a=key[0],b=key[1])
            dia, src = new, "user"
        values[pid], evidence[pid] = (dia, src), rec
        pipe.update(nominal_mm=dia, bore_provenance=deepcopy(rec))
    return values, evidence, changed


def pattern_preview(points: Sequence[Sequence[float]], source: list[dict],
                    target_edges: Iterable[Edge], *, anchor: int | None = None,
                    reverse: bool = False) -> list[dict]:
    """Manual paste proposal, not automatic topology equivalence.

    Relative path position proposes a source interval even for a different
    shape/fragment count. Every target row is visible and editable before commit;
    a target crossing a source size boundary is explicitly flagged.
    """
    nodes, path = ordered_path(target_edges, anchor)
    lengths = [math.dist(points[a][:2], points[b][:2]) for a,b in path]
    source = list(reversed(source)) if reverse else source
    total_source = sum(float(r["length_mm"]) for r in source)
    total_target = sum(lengths)
    if total_source <= 0 or total_target <= 0 or any(not r.get("nominal_mm") for r in source):
        raise ValueError("복사 원본의 관경과 길이가 모두 정의되어야 합니다.")
    boundaries, pos = [], 0.
    for r in source:
        end = pos + float(r["length_mm"])/total_source
        boundaries.append((pos,end,r))
        pos = end
    result, pos = [], 0.
    for e, length in zip(path, lengths):
        end = pos + length/total_target
        matching = [r for lo,hi,r in boundaries if lo < end-1e-9 and hi > pos+1e-9]
        candidates = sorted({int(r["nominal_mm"]) for r in matching})
        chosen = next(r for lo,hi,r in boundaries if lo-1e-9 <= (pos+end)/2 <= hi+1e-9)
        result.append(dict(edge=list(e), nominal_mm=int(chosen["nominal_mm"]),
                           candidates=candidates, needs_choice=len(candidates)>1,
                           source=deepcopy(chosen), mapping="manual_relative_path"))
        pos = end
    return result
