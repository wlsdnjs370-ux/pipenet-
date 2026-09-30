"""Evidence-limited contacts for slightly displaced CAD fitting symbols.

All coordinates are millimetres. This is not a general proximity snap: only
a 270-degree fitting arc, a centred open-side pipe end and one perpendicular
through pipe can authorize a contact. Source geometry is never moved.
"""
from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
import math

Edge = tuple[int, int]


@dataclass(frozen=True)
class ArcContact:
    """A proven through-pipe split, retaining the original symbol position."""

    edge: Edge
    fraction: float
    xy: tuple[float, float]
    offset_mm: float
    branch_node: int


def offset_arc_contact(
    points: Sequence[Sequence[float]],
    nearby_edges: Sequence[Edge],
    adjacency: Mapping[int, set[int]],
    symbol: Mapping,
    arms: Sequence[int],
    *,
    legacy_offset_mm: float = 25.0,
    maximum_offset_mm: float = 50.0,
) -> ArcContact | None:
    """Return one unambiguous, small, open-side tee contact or leave it alone.

    The extra tolerance is capped at one quarter of the symbol radius and
    50 mm. A free pipe end must sit within 5 mm of the arc centre and follow
    its opening within 8 degrees. The trunk must extend beyond the symbol
    on both sides. Parallel competing trunks, heads, circles, unknown angles,
    short strokes and wrong-side/oblique pipes do not authorize a new join.
    Existing contacts within the legacy tolerance are left to that rule.
    """
    if symbol.get("k") != "호":
        return None
    try:
        cx, cy, r, sa, sweep = (float(symbol[k]) for k in
                                ("cx", "cy", "r", "sa", "sweep"))
    except (KeyError, TypeError, ValueError):
        return None
    if not all(map(math.isfinite, (cx, cy, r, sa, sweep))):
        return None
    limit = min(maximum_offset_mm, r / 4)
    if limit <= legacy_offset_mm or abs(sweep - 270) > 0.1:
        return None
    angle = math.radians(sa + sweep + (360 - sweep) / 2)
    ux, uy = math.cos(angle), math.sin(angle)
    cos_limit = math.cos(math.radians(8))
    sin_limit = math.sin(math.radians(8))
    branches = []
    for n in set(arms):
        if math.dist(points[n][:2], (cx, cy)) > 5 or len(adjacency.get(n, ())) != 1:
            continue
        other = next(iter(adjacency[n]))
        dx, dy = points[other][0] - points[n][0], points[other][1] - points[n][1]
        length = math.hypot(dx, dy)
        if length >= r / 2 and (dx * ux + dy * uy) / length >= cos_limit:
            branches.append(n)
    if len(branches) != 1:
        return None
    candidates = []
    for a, b in sorted({tuple(sorted(e)) for e in nearby_edges}):
        ax, ay = points[a][:2]
        dx, dy = points[b][0] - ax, points[b][1] - ay
        length = math.hypot(dx, dy)
        if length <= 2 * r or abs(dx * ux + dy * uy) / length > sin_limit:
            continue
        along = ((cx - ax) * dx + (cy - ay) * dy) / length
        if min(along, length - along) < r:
            continue
        fraction = along / length
        xy = (ax + fraction * dx, ay + fraction * dy)
        distance = math.dist(xy, (cx, cy))
        if distance > limit:
            continue
        # Count all nearby trunks for ambiguity, including legacy contacts.
        candidates.append(ArcContact((a, b), fraction, xy, distance, branches[0]))
    if len(candidates) != 1:
        return None
    contact = candidates[0]
    if contact.offset_mm <= legacy_offset_mm:
        return None
    vx, vy = points[contact.branch_node][0] - contact.xy[0], points[contact.branch_node][1] - contact.xy[1]
    gap = math.hypot(vx, vy)
    if gap <= 0 or (vx * ux + vy * uy) / gap < cos_limit:
        return None
    return contact
