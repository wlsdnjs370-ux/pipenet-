"""Explicit shared-point contracts, independent of CAD/server state."""
from __future__ import annotations

import math
from typing import Sequence


def separate_valve_picks(
    valve_xy: Sequence[tuple[float, float]],
    source_xy: Sequence[tuple[float, float]],
    *, tolerance: float = 1e-6,
) -> tuple[list[int], list[int]]:
    """Partition picks into independent valves and existing source connections.

    Inputs share one coordinate system (CAD mm in the plan adapter). Only an
    explicitly coincident source/pick is the same boundary; nearby valves are
    not merged. A shared point remains a loss element, not a blind pipe branch.
    Returned indices preserve the original records and user metadata.
    """
    if not math.isfinite(tolerance) or tolerance < 0:
        raise ValueError("공통 절점 허용오차가 올바르지 않습니다.")
    if any(len(p) != 2 or not all(math.isfinite(v) for v in p)
           for p in [*valve_xy, *source_xy]):
        raise ValueError("공통 절점 좌표가 올바르지 않습니다.")
    independent, shared = [], []
    for i, point in enumerate(valve_xy):
        (shared if any(math.dist(point, source) <= tolerance for source in source_xy)
         else independent).append(i)
    return independent, shared
