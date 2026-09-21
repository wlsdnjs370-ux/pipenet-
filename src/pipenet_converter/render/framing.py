"""Choose one uniform display scale from an explicit reference subnetwork."""
from __future__ import annotations

import math
from typing import Iterable


def reference_display_scale(points: Iterable[tuple[float, float]], *,
                            canvas_units: float = 3000.0) -> float:
    """Return display units per CAD unit without stretching either axis.

    For a merged network the reference is its plan, before isometric projection,
    just like a standalone plan export. Tall risers must not shrink that scale.
    A point-only reference retains unit scale, matching standalone normalization.
    """
    values = list(points)
    if not values or any(not math.isfinite(v) for p in values for v in p):
        raise ValueError("표시 배율 기준 절점이 없거나 좌표가 올바르지 않습니다.")
    if not math.isfinite(canvas_units) or canvas_units <= 0:
        raise ValueError("표시 캔버스 크기는 양수여야 합니다.")
    xs, ys = zip(*values)
    span = max(max(xs) - min(xs), max(ys) - min(ys))
    return canvas_units / span if span > 1e-9 else 1.0
