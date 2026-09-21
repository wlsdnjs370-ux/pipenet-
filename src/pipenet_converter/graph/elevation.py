"""Explicit floor elevations for schematic sections (never infer metres from Y)."""
from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
import math
from typing import Sequence


@dataclass(frozen=True)
class FloorLevel:
    """One diagram level in CAD mm and its assigned real elevation in metres."""

    drawing_y_mm: float
    elevation_m: float


@dataclass(frozen=True)
class SectionElevation:
    """Piecewise linear section calibration with explicit boundary extrapolation.

    Fractional-floor offsets use the adjacent interval's scale, not a nearest
    floor's height. This is an assumption, not an installation-height survey.
    """

    levels: tuple[FloorLevel, ...]

    def __post_init__(self) -> None:
        if len(self.levels) < 2:
            raise ValueError("층고 환산에는 서로 다른 층 표기가 2개 이상 필요합니다.")
        if any(not math.isfinite(v) for p in self.levels
               for v in (p.drawing_y_mm, p.elevation_m)):
            raise ValueError("층 기준 좌표와 표고는 유한한 수여야 합니다.")
        if any(b.drawing_y_mm <= a.drawing_y_mm or b.elevation_m <= a.elevation_m
               for a, b in zip(self.levels, self.levels[1:])):
            raise ValueError("층 기준선의 순서 또는 표고가 서로 충돌합니다.")

    def at(self, drawing_y_mm: float) -> float:
        """Return metres at a diagram Y; tolerate extraction's mm rounding."""
        if not math.isfinite(drawing_y_mm):
            raise ValueError("계통도 좌표가 올바르지 않습니다.")
        for level in self.levels:
            if abs(drawing_y_mm - level.drawing_y_mm) <= 0.51:
                return level.elevation_m
        index = min(max(bisect_right([p.drawing_y_mm for p in self.levels],
                                    drawing_y_mm) - 1, 0), len(self.levels) - 2)
        a, b = self.levels[index:index + 2]
        return a.elevation_m + (drawing_y_mm - a.drawing_y_mm) * (
            b.elevation_m - a.elevation_m) / (b.drawing_y_mm - a.drawing_y_mm)


def uniform_floor_section(marks: Sequence[tuple[int, float]],
                          height_m: float) -> SectionElevation:
    """Build an explicitly assumed uniform profile; zero is not a floor number."""
    if not math.isfinite(height_m) or height_m <= 0:
        raise ValueError("가정 층고는 0보다 큰 유한한 수(m)여야 합니다.")
    floors: dict[int, float] = {}
    for floor, y in marks:
        if floor == 0:
            raise ValueError("0층은 지원하지 않습니다.")
        if floor in floors and abs(floors[floor] - y) > 1:
            raise ValueError(f"{floor}층 기준선이 여러 위치에 있습니다.")
        floors[floor] = y
    levels = tuple(FloorLevel(y, (floor if floor > 0 else floor + 1) * height_m)
                   for floor, y in sorted(floors.items()))
    return SectionElevation(levels)
