"""Evidence-backed inline head junctions for preserved networks only.

A selected head circle can conceal a run and a third pipe port. Attaching the
nozzle to just one arm loses the third pipe, even after the run gap is repaired.
This module preserves that tee topology without treating arbitrary intersections,
nearby marks, different materials or four-port symbols as confirmed fittings.
It never edits the legacy tree, moves coordinates, or assigns hydraulic losses.
"""
from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
import math

from .cycle_closures import ClosureResult
from .flow import Edge, edge_key

XY = tuple[float, float]


@dataclass(frozen=True)
class HeadPort:
    """A drawn pipe endpoint on the selected head's circumference (mm)."""

    node: int
    xy: XY
    outward: XY
    material: str


@dataclass(frozen=True)
class HeadJunction:
    """Unmodified source evidence, captured before any tree-only pruning."""

    head_xy: XY
    radius_mm: float
    ports: tuple[HeadPort, ...]

    @classmethod
    def from_record(cls, row: dict) -> HeadJunction:
        """Validate cached evidence; malformed records must not create links."""
        result = cls(tuple(row['head_xy']), float(row['radius_mm']), tuple(
            HeadPort(int(p['node']), tuple(p['xy']), tuple(p['outward']), str(p['material']))
            for p in row['ports']))
        coords = [result.head_xy] + [xy for p in result.ports for xy in (p.xy, p.outward)]
        if (not math.isfinite(result.radius_mm) or result.radius_mm <= 0
                or not 3 <= len(result.ports) <= 4
                or len({p.node for p in result.ports}) != len(result.ports)
                or any(len(xy) != 2 or not all(math.isfinite(float(v)) for v in xy)
                       for xy in coords)
                or any(p.node < 0 or not p.material or math.hypot(*p.outward) < .999
                       or math.hypot(*p.outward) > 1.001 for p in result.ports)):
            raise ValueError('헤드 접속부의 원본 배관 근거가 올바르지 않습니다.')
        return result


def collect_head_junctions(
    points: Sequence[Sequence[float]], edges: Iterable[Edge],
    heads: Sequence[Sequence[float]], materials: Mapping[Edge, object], *,
    boundary_tolerance_mm: float = 2.0, angle_tolerance_deg: float = 8.0,
) -> tuple[HeadJunction, ...]:
    """Find radial pipe endpoints on mapped heads; never infer from proximity alone.

    At least three ports and an opposed run are necessary. A port must have one
    exterior drawn arm, terminate on the actual circle boundary, and point away
    from its center. Material conflicts / four ports remain explicit review rows.
    No project-specific layer names or CAD applications are used.
    """
    if (not math.isfinite(boundary_tolerance_mm) or boundary_tolerance_mm <= 0
            or not math.isfinite(angle_tolerance_deg) or not 0 < angle_tolerance_deg <= 8):
        raise ValueError('헤드 접속부의 거리·각도 허용오차가 올바르지 않습니다.')
    adj: dict[int, set[int]] = defaultdict(set)
    grid: dict[tuple[int, int], list[int]] = defaultdict(list)
    for a, b in edges:
        adj[a].add(b)
        adj[b].add(a)
    cell = 500.0
    for n in adj:
        x, y = points[n][:2]
        grid[(math.floor(x / cell), math.floor(y / cell))].append(n)
    cosine = math.cos(math.radians(angle_tolerance_deg))
    out = []
    seen = set()
    for hx, hy, hr in heads:
        if hr <= 0 or (hx, hy, hr) in seen:
            continue
        seen.add((hx, hy, hr))
        ports = []
        reach = hr + boundary_tolerance_mm
        for gx in range(math.floor((hx - reach) / cell), math.floor((hx + reach) / cell) + 1):
            for gy in range(math.floor((hy - reach) / cell), math.floor((hy + reach) / cell) + 1):
                for n in grid.get((gx, gy), ()):
                    px, py = points[n][:2]
                    radial = math.hypot(px - hx, py - hy)
                    if radial <= 0 or abs(radial - hr) > boundary_tolerance_mm:
                        continue
                    ux, uy = (px - hx) / radial, (py - hy) / radial
                    arms = []
                    for other in adj[n]:
                        ox, oy = points[other][:2]
                        dx, dy = ox - px, oy - py
                        length = math.hypot(dx, dy)
                        material = materials.get(edge_key(n, other))
                        if (length > boundary_tolerance_mm and material is not None
                                and (dx * ux + dy * uy) / length >= cosine):
                            arms.append(HeadPort(n, (px, py), (dx / length, dy / length), repr(material)))
                    # Interior continuations or competing exterior arms make
                    # this a non-terminal, not a concealed source port.
                    if len(arms) == len(adj[n]) == 1:
                        ports.extend(arms)
        ports.sort(key=lambda p: p.node)
        if 3 <= len(ports) <= 4 and any(
                a.outward[0] * b.outward[0] + a.outward[1] * b.outward[1] <= -cosine
                for i, a in enumerate(ports) for b in ports[i + 1:]):
            out.append(HeadJunction((hx, hy), hr, tuple(ports)))
    return tuple(out)


def preserve_head_junctions(
    points: Sequence[Sequence[float]], source: ClosureResult,
    candidates: Sequence[HeadJunction], *, head_centers: Iterable[int],
    forbidden_segments: Sequence[Sequence[Sequence[float]]] = (),
    tolerance_mm: float = .2,
) -> ClosureResult:
    """Join verified tee ports at the existing head center in a derived graph.

    Existing run spans are split, not duplicated. Explicit full/partial deletions,
    missing/changed arms and ambiguous centers fail closed. Unlike cycle closure
    repair, a tee can connect an isolated drawn branch: an existing return route
    is not a prerequisite for a physical junction.
    """
    if not math.isfinite(tolerance_mm) or not 0 < tolerance_mm <= 1:
        raise ValueError('헤드 접속부의 좌표 허용오차는 0 초과 1 mm 이하여야 합니다.')
    effective = set(source.edges)
    adj: dict[int, set[int]] = defaultdict(set)
    for a, b in effective:
        adj[a].add(b)
        adj[b].add(a)
    centers: dict[tuple[int, int], list[int]] = defaultdict(list)
    for n in set(head_centers):
        if 0 <= n < len(points):
            x, y = points[n][:2]
            centers[(math.floor(x), math.floor(y))].append(n)
    records = list(source.records)
    cosine = math.cos(math.radians(8))

    def blocked(a: Sequence[float], b: Sequence[float]) -> bool:
        dx, dy = b[0] - a[0], b[1] - a[1]
        length = math.hypot(dx, dy)
        if length <= tolerance_mm:
            return True
        ux, uy = dx / length, dy / length
        for p, q in forbidden_segments:
            ts = [(xy[0] - a[0]) * ux + (xy[1] - a[1]) * uy for xy in (p, q)]
            ds = [abs((xy[0] - a[0]) * uy - (xy[1] - a[1]) * ux) for xy in (p, q)]
            if max(ds) <= tolerance_mm and min(max(ts), length) - max(min(ts), 0) > tolerance_mm:
                return True
        return False

    for candidate in candidates:
        rec = {'head_xy': list(candidate.head_xy), 'evidence': 'drawn_head_tee_ports',
               'ports': [p.node for p in candidate.ports], 'status': 'review'}
        records.append(rec)
        if len(candidate.ports) != 3:
            rec['status'] = 'four_ports_need_review'
            continue
        if len({p.material for p in candidate.ports}) != 1:
            rec['status'] = 'material_conflict'
            continue
        hx, hy = candidate.head_xy
        matches = [n for gx in range(math.floor(hx) - 1, math.floor(hx) + 2)
                   for gy in range(math.floor(hy) - 1, math.floor(hy) + 2)
                   for n in centers.get((gx, gy), ())
                   if math.dist(points[n][:2], candidate.head_xy) <= tolerance_mm]
        if len(matches) != 1:
            rec['status'] = 'head_center_unresolved'
            continue
        center = matches[0]
        ports = {p.node for p in candidate.ports}
        if center in ports or any(not 0 <= p.node < len(points)
                                 or math.dist(points[p.node][:2], p.xy) > tolerance_mm
                                 for p in candidate.ports):
            rec['status'] = 'source_changed'
            continue
        deleted_ports = {p.node for p in candidate.ports if blocked(p.xy, points[center])}
        # Cached source must still describe a radial tee, not an arbitrary fan.
        distances = [math.dist(p.xy, candidate.head_xy) for p in candidate.ports]
        if any(d <= tolerance_mm for d in distances):
            rec['status'] = 'port_geometry_needs_review'
            continue
        radial = [((p.xy[0] - hx) / d, (p.xy[1] - hy) / d)
                  for p, d in zip(candidate.ports, distances)]
        dots = [a[0] * b[0] + a[1] * b[1] for i, a in enumerate(radial) for b in radial[i + 1:]]
        if (any(abs(math.dist(p.xy, candidate.head_xy) - candidate.radius_mm) > 2
                for p in candidate.ports)
                or sum(d <= -cosine for d in dots) != 1
                or any(d > -cosine and abs(d) > math.sin(math.radians(8)) for d in dots)):
            rec['status'] = 'port_geometry_needs_review'
            continue
        if adj[center] - ports:
            rec['status'] = 'head_center_conflict'
            continue
        valid = True
        for port in candidate.ports:
            exterior = []
            for other in adj[port.node] - ports - {center}:
                dx, dy = points[other][0] - port.xy[0], points[other][1] - port.xy[1]
                length = math.hypot(dx, dy)
                if length > tolerance_mm and (dx * port.outward[0] + dy * port.outward[1]) / length >= cosine:
                    exterior.append(other)
            if len(exterior) != 1:
                valid = False
                break
        if not valid:
            rec['status'] = 'exterior_arm_changed'
            continue
        needed = {edge_key(p, center) for p in ports - deleted_ports}
        replaced = set()
        for i, a in enumerate(candidate.ports):
            for b in candidate.ports[i + 1:]:
                e = edge_key(a.node, b.node)
                if e in effective:
                    if abs(math.dist(a.xy, points[center]) + math.dist(points[center], b.xy)
                           - math.dist(a.xy, b.xy)) > tolerance_mm:
                        valid = False
                    replaced.add(e)
        if not valid:
            rec['status'] = 'non_axial_existing_link'
            continue
        deleted_necks = {edge_key(p, center) for p in deleted_ports} & effective
        if needed <= effective and not replaced and not deleted_necks:
            rec['status'] = 'user_deleted' if deleted_ports else 'already_connected'
            continue
        added = needed - effective
        # Split the former through-span even when a derived half-neck was
        # explicitly deleted. Otherwise the old unsplit span would bypass the
        # deletion on the next projection or reload. Keep the surviving half.
        for a, b in replaced | deleted_necks:
            effective.remove((a, b))
            adj[a].discard(b)
            adj[b].discard(a)
        for a, b in added:
            effective.add((a, b))
            adj[a].add(b)
            adj[b].add(a)
        rec.update(status='user_deleted' if deleted_ports else 'restored', center=center,
                   added_edges=[list(e) for e in sorted(added)],
                   split_edges=[list(e) for e in sorted(replaced)])
    base = (source.edges - source.added) | source.removed
    return ClosureResult(frozenset(effective), frozenset(effective - base), tuple(records),
                         frozenset(base - effective))
