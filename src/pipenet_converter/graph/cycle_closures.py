"""Restore only evidenced head-symbol gaps rejected by a legacy no-cycle gate.

This is an additive, opt-in projection. It never changes the stored tree graph,
connects an arbitrary crossing, or reconstructs an unexplained gap. Candidates
must come from the existing DXF same-material/opposed-arm/head-cover matcher.
"""
from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
import math

from .flow import Edge, edge_key


@dataclass(frozen=True)
class ClosureCandidate:
    """Source matcher evidence captured BEFORE its no-cycle filter."""

    a: int
    b: int
    xy_a: tuple[float, float]
    xy_b: tuple[float, float]
    head_xy: tuple[float, float]
    head_radius_mm: float
    evidence: str = "same_material_head_cover"

    @classmethod
    def from_record(cls, row: dict) -> ClosureCandidate:
        """Validate JSON cache records instead of accepting unproven endpoints."""
        candidate = cls(int(row["a"]), int(row["b"]), tuple(row["xy_a"]),
                        tuple(row["xy_b"]), tuple(row["head_xy"]),
                        float(row["head_radius_mm"]), str(row["evidence"]))
        if (candidate.evidence != "same_material_head_cover"
                or candidate.head_radius_mm <= 0
                or any(len(xy) != 2 for xy in (candidate.xy_a, candidate.xy_b, candidate.head_xy))
                or not all(math.isfinite(float(v)) for xy in
                           (candidate.xy_a, candidate.xy_b, candidate.head_xy) for v in xy)
                or not math.isfinite(candidate.head_radius_mm)):
            raise ValueError("루프 이음 후보의 도면 근거가 올바르지 않습니다.")
        return candidate


@dataclass(frozen=True)
class ClosureResult:
    """Effective edges plus accepted/rejected candidates for screen and audit."""

    edges: frozenset[Edge]
    added: frozenset[Edge]
    records: tuple[dict, ...]
    removed: frozenset[Edge] = frozenset()

    def report(self) -> dict:
        """Compact counts, without claiming a hydraulic interpretation."""
        return {"candidates": len(self.records),
                "restored": sum(r["status"] == "restored" for r in self.records),
                "already_connected": sum(r["status"] == "already_connected" for r in self.records),
                "review": sum(r["status"] not in ("restored", "already_connected", "duplicate_removed") for r in self.records),
                "duplicate_joins_removed": sum(r['status']=='duplicate_removed' for r in self.records),
                "added_edges": len(self.added), "split_edges": len(self.removed)}


def recover_cycle_closures(
    points: Sequence[Sequence[float]], edges: Iterable[Edge],
    candidates: Sequence[ClosureCandidate], *, head_centers: Iterable[int] = (),
    forbidden_segments: Sequence[Sequence[Sequence[float]]] = (),
    tolerance_mm: float = 0.2,
) -> ClosureResult:
    """Recover cycle-closing symbol spans; preserve centers and manual deletes.

    A half head-neck may already have been added by the legacy editor. Split the
    missing span at that EXISTING center rather than adding a bypass triangle.
    Both exterior pipe arms must still exist and endpoints must already belong
    to the same component. A deleted return route cannot be silently repaired.
    """
    if not math.isfinite(tolerance_mm) or tolerance_mm <= 0:
        raise ValueError("루프 이음 허용오차는 양의 유한한 mm 값이어야 합니다.")
    original = frozenset(edge_key(a, b) for a, b in edges)
    adj: dict[int, set[int]] = defaultdict(set)
    parents: dict[int, int] = {}

    def root(n: int) -> int:
        parents.setdefault(n, n)
        while parents[n] != n:
            parents[n] = parents[parents[n]]
            n = parents[n]
        return n

    for a, b in original:
        adj[a].add(b)
        adj[b].add(a)
        ra, rb = root(a), root(b)
        parents[ra] = rb
    centers = set(head_centers)
    effective, records = set(original), []
    for candidate in candidates:
        a, b = candidate.a, candidate.b
        rec = {"a": a, "b": b, "head_xy": list(candidate.head_xy),
               "evidence": candidate.evidence, "status": "review"}
        records.append(rec)
        if (not 0 <= a < len(points) or not 0 <= b < len(points) or a == b
                or math.dist(points[a][:2], candidate.xy_a) > tolerance_mm
                or math.dist(points[b][:2], candidate.xy_b) > tolerance_mm):
            rec["status"] = "source_changed"
            continue
        dx, dy = points[b][0] - points[a][0], points[b][1] - points[a][1]
        length = math.hypot(dx, dy)
        if length <= tolerance_mm:
            rec["status"] = "degenerate_span"
            continue
        ux, uy = dx / length, dy / length

        def project(xy):
            x, y = xy[0] - points[a][0], xy[1] - points[a][1]
            return x * ux + y * uy, abs(x * uy - y * ux)

        # Block even a partial deletion on this span. Source endpoint-only
        # deletion checks miss the newly split half-neck at the head center.
        blocked = False
        for seg in forbidden_segments:
            t1, d1 = project(seg[0]); t2, d2 = project(seg[1])
            if (max(d1, d2) <= tolerance_mm
                    and min(max(t1, t2), length) - max(min(t1, t2), 0) > tolerance_mm):
                blocked = True
                break
        if blocked:
            rec["status"] = "user_deleted"
            continue
        if edge_key(a, b) in original:
            rec["status"] = "already_connected"
            continue
        # Only adjacent existing head centers may split the gap. Proximity to
        # an unrelated pipe/symbol is not connection evidence.
        inner = []
        for n in (adj[a] | adj[b]) & centers:
            along, lat = project(points[n])
            if tolerance_mm < along < length - tolerance_mm and lat <= tolerance_mm:
                inner.append((along, n))
        inner.sort()
        if any(abs(x[0] - y[0]) <= tolerance_mm for x, y in zip(inner, inner[1:])):
            rec["status"] = "ambiguous_centers"
            continue
        chain = [a] + [n for _, n in inner] + [b]
        needed = {edge_key(u, v) for u, v in zip(chain, chain[1:])}
        if needed <= effective:
            rec["status"] = "already_connected"
            continue

        def exterior(n: int, left: bool) -> bool:
            found = []
            for other in adj[n] - set(chain):
                t, lat = project(points[other])
                if (t < -tolerance_mm if left else t > length + tolerance_mm):
                    # Source matcher already checked same material and axial
                    # continuity. Reject changed geometry rather than re-guess.
                    axial = -t if left else t - length
                    if lat <= tolerance_mm + axial * math.sin(math.radians(5)):
                        found.append(other)
            return len(found) == 1

        if not exterior(a, True) or not exterior(b, False):
            rec["status"] = "exterior_arm_changed"
            continue
        if root(a) != root(b):
            rec["status"] = "return_route_missing"
            continue
        effective.update(needed)
        rec.update(status="restored", chain=chain,
                   added_edges=[list(e) for e in sorted(needed - original)],
                   span_mm=length)
    return ClosureResult(frozenset(effective), frozenset(effective - original), tuple(records))
