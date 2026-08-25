"""
domain/geometry.py

K-Fire 도메인 계층 — 순수 기하학/수학 유틸리티
PySide6 의존성 없음. UI·보고서·엔진 어디서든 재사용 가능.
"""
from __future__ import annotations
import math

# ---------------------------------------------------------------------------
# 상수
# ---------------------------------------------------------------------------

TOLERANCE = 1e-6  # 부동소수점 비교 오차 허용 범위

# 등각투영 기본 각도 (30°)
_ISO_ANGLE: float = math.radians(30)
ISO_COS: float = math.cos(_ISO_ANGLE)
ISO_SIN: float = math.sin(_ISO_ANGLE)


# ---------------------------------------------------------------------------
# 기본 수치 비교
# ---------------------------------------------------------------------------

def is_close(a: float, b: float, tol: float = TOLERANCE) -> bool:
    """두 실수가 오차 범위 내에서 같은지 확인"""
    return abs(a - b) <= tol


def is_same_point(
    p1: tuple[float, float, float],
    p2: tuple[float, float, float],
    tol: float = TOLERANCE,
) -> bool:
    """두 점(x,y,z)이 오차 범위 내에서 같은 위치인지 확인"""
    return (
        is_close(p1[0], p2[0], tol)
        and is_close(p1[1], p2[1], tol)
        and is_close(p1[2], p2[2], tol)
    )


# ---------------------------------------------------------------------------
# 거리 / 스냅
# ---------------------------------------------------------------------------

def dist_3d(
    p1: tuple[float, float, float],
    p2: tuple[float, float, float],
) -> float:
    """두 3D 점 사이의 유클리드 거리"""
    return math.sqrt(
        (p2[0] - p1[0]) ** 2 + (p2[1] - p1[1]) ** 2 + (p2[2] - p1[2]) ** 2
    )


def snap_value(val: float, step: float) -> float:
    """값을 그리드 스텝 단위로 스냅"""
    if step <= 0:
        return val
    return round(val / step) * step


def snap_point(
    pt: tuple[float, float, float],
    step: float,
) -> tuple[float, float, float]:
    """좌표(x,y,z)를 그리드 스텝 단위로 스냅"""
    return (
        snap_value(pt[0], step),
        snap_value(pt[1], step),
        snap_value(pt[2], step),
    )


# ---------------------------------------------------------------------------
# 벡터 연산
# ---------------------------------------------------------------------------

def normalize_vector(
    v: tuple[float, float, float],
) -> tuple[float, float, float]:
    """벡터 정규화 (길이를 1로 만듦)"""
    mag = math.sqrt(v[0] ** 2 + v[1] ** 2 + v[2] ** 2)
    if mag == 0:
        return (0.0, 0.0, 0.0)
    return (v[0] / mag, v[1] / mag, v[2] / mag)


def dot_product(
    v1: tuple[float, float, float],
    v2: tuple[float, float, float],
) -> float:
    """벡터 내적"""
    return v1[0] * v2[0] + v1[1] * v2[1] + v1[2] * v2[2]


def angle_between_vectors(
    v1: tuple[float, float, float],
    v2: tuple[float, float, float],
) -> float:
    """
    두 벡터 사이의 각도(Degree).
    0°: 같은 방향, 90°: 직각, 180°: 정반대 방향
    """
    u1 = normalize_vector(v1)
    u2 = normalize_vector(v2)
    dot = dot_product(u1, u2)
    dot = max(-1.0, min(1.0, dot))  # 부동소수점 오차 보정
    return math.degrees(math.acos(dot))


def get_vector(
    p1: tuple[float, float, float],
    p2: tuple[float, float, float],
) -> tuple[float, float, float]:
    """p1에서 p2로 향하는 벡터"""
    return (p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2])


# ---------------------------------------------------------------------------
# 등각투영 좌표 변환  (view_graphics.py · report_module.py 중복 제거용 단일 구현)
# ---------------------------------------------------------------------------

def to_screen_coord(
    x: float,
    y: float,
    z: float,
    scale: float = 1.0,
    cos_a: float = ISO_COS,
    sin_a: float = ISO_SIN,
) -> tuple[float, float]:
    """
    3D 좌표 (x, y, z) → 등각투영 2D 화면 좌표 (u, v) 변환.

    좌표계 규칙:
      u: x 오른쪽(+), y 왼쪽(-) → u = (x - y) · cos_a · scale
      v: 화면 아래가 (+) 이므로, z 위 = v 감소,
         x·y 증가 = 화면 위쪽으로 이동 (V자 투영)
         → v = (-(x + y) · sin_a - z) · scale

    기본값(cos_a, sin_a)은 30° 등각투영 기준.
    view_graphics.py 와 report_module.py 모두 이 함수를 호출하도록 점진적으로 교체 예정.
    """
    u = (x - y) * cos_a * scale
    v = (-(x + y) * sin_a - z) * scale
    return u, v


# ---------------------------------------------------------------------------
# __all__
# ---------------------------------------------------------------------------

__all__ = [
    "TOLERANCE",
    "ISO_COS",
    "ISO_SIN",
    "is_close",
    "is_same_point",
    "dist_3d",
    "snap_value",
    "snap_point",
    "normalize_vector",
    "dot_product",
    "angle_between_vectors",
    "get_vector",
    "to_screen_coord",
]
