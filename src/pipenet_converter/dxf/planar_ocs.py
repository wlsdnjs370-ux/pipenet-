"""DXF planar OCS transforms, composed before parent INSERT transforms."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Sequence
import math


@dataclass(frozen=True)
class PlanarOCS:
    """Map an entity's XY millimetres into its parent's XY coordinate system."""

    parent: Callable[[float, float], tuple[float, float]]
    scale: float
    x_sign: float

    def __call__(self, x: float, y: float) -> tuple[float, float]:
        return self.parent(self.x_sign * x, y)


def planar_ocs(parent: Callable, extrusion: Sequence[float]) -> PlanarOCS:
    """Resolve ±Z normals; do not silently render tilted circles as flat circles.

    DXF's arbitrary-axis rule gives X=(-1,0,0), Y=(0,1,0) for -Z.
    LINE and MTEXT are already WCS and must not pass through this transform.
    """
    x, y, z = map(float, extrusion)
    if not all(math.isfinite(v) for v in (x, y, z)) or abs(z) < 1e-12 or math.hypot(x, y) > abs(z)*1e-9:
        raise ValueError('기울어진 OCS 객체는 평면 배관 도면으로 처리할 수 없습니다. XY 평면 DXF로 변환해 주세요.')
    return PlanarOCS(parent, parent.scale, -1.0 if z < 0 else 1.0)
