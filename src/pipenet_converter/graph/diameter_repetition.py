"""Partial, non-recursive diameter correspondence between physical branch slots.

Coordinates and lengths are millimetres. This module never edits geometry or
uses head counts to invent sizes. All donors are snapshotted before any copy.
"""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
import math
from typing import Any, Mapping, Protocol, Sequence

from .flow import Edge
from ..progress import report


class RepeatConfig(Protocol):
    """Drawing-independent tolerances supplied by the H definition policy."""

    repeat_length_tolerance: float
    repeat_shape_tolerance: float
    straight_tolerance_deg: float


@dataclass(frozen=True)
class BranchProfile:
    """An oriented branch with head-bounded slots, invariant to CAD fragments."""

    nodes: tuple[int, ...]
    slots: tuple[tuple[Edge, ...], ...]
    lengths: tuple[float, ...]
    shapes: tuple[tuple[tuple[float, float], ...], ...]
    kinds: tuple[str, ...]
    roles: tuple[int, int]
    component: int
    axis: tuple[float, float]


def _profile(points: Sequence[Sequence[float]], nodes: list[int], path: list[Edge],
             heads: Mapping[int, Any], degrees: Mapping[int, int], component: int) -> BranchProfile:
    slots, lengths, shapes, current = [], [], [], []
    first = next((i for i in range(1, len(nodes))
                  if math.dist(points[nodes[0]][:2], points[nodes[i]][:2]) > 1e-6), None)
    if first is None:
        raise ValueError('반복 가지의 길이는 0보다 커야 합니다.')
    a, b = points[nodes[0]], points[nodes[first]]
    norm = math.dist(a[:2], b[:2])
    ux, uy = (b[0]-a[0])/norm, (b[1]-a[1])/norm
    start = 0
    for i, edge in enumerate(path):
        current.append(edge)
        if nodes[i+1] not in heads and i != len(path)-1:
            continue
        pts = [points[n][:2] for n in nodes[start:i+2]]
        ls = [math.dist(p, q) for p, q in zip(pts, pts[1:])]
        length = sum(ls)
        if length <= 1e-6:
            raise ValueError('반복 구간의 길이는 0보다 커야 합니다.')
        shape = []
        for fraction in (.25, .5, .75, 1.):
            remaining = length*fraction
            for p, q, size in zip(pts, pts[1:], ls):
                if remaining <= size+1e-8:
                    ratio = min(1., remaining/size) if size else 0.
                    dx = p[0]+ratio*(q[0]-p[0])-pts[0][0]
                    dy = p[1]+ratio*(q[1]-p[1])-pts[0][1]
                    shape.append(((dx*ux+dy*uy)/length, (-dx*uy+dy*ux)/length))
                    break
                remaining -= size
        slots.append(tuple(current)); lengths.append(length); shapes.append(tuple(shape))
        current, start = [], i+1
    return BranchProfile(tuple(nodes), tuple(slots), tuple(lengths), tuple(shapes),
                         tuple(str(heads[n]) for n in nodes[1:] if n in heads),
                         (degrees[nodes[0]], degrees[nodes[-1]]), component, (ux, uy))


def _matches(a: BranchProfile, b: BranchProfile, cfg: RepeatConfig) -> bool:
    if a.component != b.component or a.roles != b.roles or a.kinds != b.kinds or len(a.slots) != len(b.slots):
        return False
    # Two-main (loop) chains have no terminal endpoint to fix their orientation.
    # Require the same physical direction; node numbering cannot pick a reversal.
    if a.roles[1] != 1 and sum(x*y for x, y in zip(a.axis, b.axis)) < math.cos(math.radians(cfg.straight_tolerance_deg)):
        return False
    return (all(abs(x-y) <= max(x, y)*cfg.repeat_length_tolerance+1e-7 for x, y in zip(a.lengths, b.lengths))
            and all(len(x) == len(y) and all(math.dist(p, q) <= cfg.repeat_shape_tolerance for p, q in zip(x, y))
                    for x, y in zip(a.shapes, b.shapes)))


def _slot_source(slot: tuple[Edge, ...], rows: Mapping[Edge, dict]) -> tuple[int, Edge] | None:
    records = [rows[e] for e in slot]
    values = {r.get('text_mm') for r in records}
    if (len(values) != 1 or None in values or any(r.get('block_export') for r in records)
            or any(r.get('definition_source') == 'drawing_repeat' for r in records)):
        return None
    # Every contributing fragment must itself lead to original, not fabricated text.
    if not all(r.get('annotation_id') is not None and r.get('text_xy_mm') for r in records):
        return None
    origin = min(slot, key=lambda e: (tuple(rows[e]['text_xy_mm']), str(rows[e]['annotation_id'])))
    return int(next(iter(values))), origin


def transfer_repeated_slots(points: Sequence[Sequence[float]], edges: set[Edge],
                            branches: list[tuple[list[int], list[Edge]]],
                            heads: Mapping[int, Any], stops: set[int], rows: dict[Edge, dict],
                            config: RepeatConfig, policy: str) -> dict[str, int]:
    """Transfer each supported slot independently; retain conflicts and source IDs.

    A direct target annotation always wins. An unknown slot with conflicting
    donors stays unknown. A blocked slot is never unlocked by a repeat. Donors
    cannot come from another disconnected/valve-separated component or from
    values produced earlier in this pass.
    """
    incident: dict[int, list[int]] = defaultdict(list)
    for a, b in edges:
        incident[a].append(b); incident[b].append(a)
    components: dict[int, int] = {}
    for node in sorted(incident):
        if node in components or node in stops:
            continue
        stack = [node]
        while stack:
            n = stack.pop()
            if n in components or n in stops:
                continue
            components[n] = node
            stack.extend(incident[n])
    profiles = []
    degrees = {n: len(v) for n, v in incident.items()}
    for nodes, path in branches:
        if any(n in stops for n in nodes) or not any(n in heads for n in nodes[1:]):
            continue
        if any(heads[n] is None or str(heads[n]).strip() in ('', 'unknown') for n in nodes if n in heads):
            continue
        profiles.append(_profile(points, nodes, path, heads, degrees, components[nodes[0]]))
        if degrees[nodes[-1]] >= 3:
            profiles.append(_profile(points, nodes[::-1], path[::-1], heads, degrees, components[nodes[0]]))
    frozen = deepcopy(rows)
    donors = {p: tuple(_slot_source(s, frozen) for s in p.slots) for p in profiles}
    groups: dict[tuple, list[BranchProfile]] = defaultdict(list)
    for p in profiles:
        groups[p.component, p.roles, p.kinds, len(p.slots)].append(p)
    proposals: dict[Edge, list[tuple[int, Edge, BranchProfile, BranchProfile, int]]] = defaultdict(list)
    stats = dict(repeated_edges=0, repeat_conflict_edges=0, repeat_partial_donors=0, repeat_profiles=len(profiles))
    stats['repeat_partial_donors'] = sum(any(v) and not all(v) for v in donors.values())
    for target in profiles:
        for donor in groups[target.component, target.roles, target.kinds, len(target.slots)]:
            if set(target.nodes) == set(donor.nodes) or not _matches(target, donor, config):
                continue
            for i, slot in enumerate(target.slots):
                evidence = donors[donor][i]
                if evidence is None:
                    continue
                value, origin = evidence
                for edge in slot:
                    proposals[edge].append((value, origin, donor, target, i))
    for edge, candidates in sorted(proposals.items()):
        if frozen[edge].get('text_mm') or frozen[edge].get('block_export'):
            continue
        values = {c[0] for c in candidates}
        # A target's own label on another CAD fragment of this SAME slot wins.
        target_slots = {target.slots[i] for _, _, _, target, i in candidates}
        direct_values = {frozen[e]['text_mm'] for slot in target_slots for e in slot if frozen[e].get('text_mm')}
        blocked = any(frozen[e].get('block_export') for slot in target_slots for e in slot)
        if blocked or len(values) != 1 or (direct_values and direct_values != values):
            rows[edge].update(reason='representative_conflict', block_export=True,
                              candidate_mm=sorted(values | direct_values),
                              review_reasons=['이 구간의 대표 표기 또는 대상 표기가 충돌합니다. 다른 구간의 대응은 유지합니다.'])
            stats['repeat_conflict_edges'] += 1
            continue
        # Preserve all supporting origins in audit, choose one deterministic native
        # source for the displayed link (not the smallest graph node ID).
        midpoint = tuple((points[edge[0]][i]+points[edge[1]][i])/2 for i in (0,1))
        value, origin, donor, target, index = min(candidates, key=lambda c:
            (math.dist(frozen[c[1]]['text_xy_mm'], midpoint),
             tuple(frozen[c[1]]['text_xy_mm']), str(frozen[c[1]]['annotation_id'])))
        record = deepcopy(frozen[origin])
        record.update(inferred=True, definition_source='drawing_repeat', method='반복 구조 구간별 대응',
                      source_edge=list(origin), representative_nodes=list(donor.nodes), target_nodes=list(target.nodes),
                      representative_slot=index, source_slot_edges=[list(e) for e in donor.slots[index]],
                      target_slot_edges=[list(e) for e in target.slots[index]],
                      supporting_annotation_ids=sorted({str(frozen[c[1]]['annotation_id']) for c in candidates}),
                      review_reasons=['연결·헤드 종류·구간 길이·꺾임을 대조한 원본 표기 참조값입니다.'], policy=policy)
        rows[edge] = record
        stats['repeated_edges'] += 1
        report('diameter', '반복 구조 구간 대응', edge=list(edge), xy=[list(points[n][:2]) for n in edge], evidence=record)
    return stats
