"""Recover metre-scale straight-pipe geometry from unprojected display bearings.

Display distances are not authoritative. Each edge uses its declared length
and the real node elevations. All edges, including cycle-closing edges, are
checked; inconsistent cycles are reported rather than removed or stretched.
"""
from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass
import math
from typing import Mapping, Sequence


@dataclass(frozen=True)
class LayoutEdge:
    """Straight pipe constraint; length and elevations are in metres."""

    label: str
    start: str
    end: str
    length_m: float


def physical_layout(
    nodes: Mapping[str, tuple[float, float, float]],
    edges: Sequence[LayoutEdge],
    *, roots: Sequence[str] = (), tolerance_m: float = 0.005,
) -> dict[str, tuple[float, float, float]]:
    """Return XYZ metres using XY only as an unprojected direction reference.

    This function neither joins components nor changes hydraulic values.
    Disconnected components, a missing horizontal bearing, impossible rise,
    or incompatible loop produces a labelled validation error.
    """
    if tolerance_m <= 0 or not math.isfinite(tolerance_m):
        raise ValueError("좌표 허용오차는 양수여야 합니다.")
    if any(not all(math.isfinite(v) for v in xyz) for xyz in nodes.values()):
        raise ValueError("유한하지 않은 배관 좌표가 있습니다.")
    adjacency = defaultdict(list)
    vectors = {}
    for edge in edges:
        if edge.label in vectors:
            raise ValueError(f"{edge.label}: 중복 배관 라벨입니다.")
        if edge.start not in nodes or edge.end not in nodes:
            raise ValueError(f"{edge.label}: 연결 절점이 없습니다.")
        a, b = nodes[edge.start], nodes[edge.end]
        length, dz = edge.length_m, b[2] - a[2]
        if not math.isfinite(length) or length < 0 or abs(dz) > length + 1e-6:
            raise ValueError(f"{edge.label}: 배관 길이 {length:g} m보다 표고차 {abs(dz):g} m가 큽니다.")
        horizontal = math.sqrt(max(0.0, length * length - dz * dz))
        dx, dy = b[0] - a[0], b[1] - a[1]
        bearing = math.hypot(dx, dy)
        if horizontal <= 1e-6:
            dx = dy = 0.0
        elif bearing <= 1e-12:
            raise ValueError(f"{edge.label}: 수평 방향을 정할 원래 평면 좌표가 없습니다.")
        else:
            dx, dy = dx * horizontal / bearing, dy * horizontal / bearing
        vectors[edge.label] = (dx, dy, dz)
        adjacency[edge.start].append((edge.end, (dx, dy, dz)))
        adjacency[edge.end].append((edge.start, (-dx, -dy, -dz)))
    result = {}
    components = []
    for root in dict.fromkeys([*roots, *nodes]):
        if root not in nodes or root in result:
            continue
        components.append(root)
        result[root] = (0.0, 0.0, nodes[root][2])
        queue = deque([root])
        while queue:
            current = queue.popleft()
            x, y, _ = result[current]
            for neighbor, (dx, dy, _dz) in adjacency[current]:
                if neighbor in result:
                    continue
                result[neighbor] = (x + dx, y + dy, nodes[neighbor][2])
                queue.append(neighbor)
    if len(components) > 1:
        raise ValueError("급수망이 여러 덩어리로 끊겨 있습니다: " + ", ".join(components[:8]))
    for edge in edges:
        a, b = result[edge.start], result[edge.end]
        vector = vectors[edge.label]
        mismatch = math.dist(tuple(b[i] - a[i] for i in range(3)), vector)
        if mismatch > tolerance_m:
            raise ValueError(f"{edge.label}: 루프의 길이·방향이 {mismatch:.4f} m 어긋납니다. 연결은 유지하고 좌표를 확인하세요.")
    return result
