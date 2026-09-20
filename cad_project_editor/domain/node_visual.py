"""노드 시각화 분류 — 순수 도메인 로직 (PySide6 의존 없음).

draw_all()에서 반복되던 80줄의 노드 타입 판별/색상 결정 코드를 단일 함수로 추출.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass
from domain.constants import fitting_category_from_id, is_check_valve_category
from domain.models import NodeTypeId

# RGB 튜플 (PySide6 QColor 대신 순수 값)
_COLOR_GRAY = (120, 120, 120)
_COLOR_BLUE = (0, 0, 255)
_COLOR_RED = (255, 0, 0)
_COLOR_DARK_GREEN = (0, 128, 0)
_COLOR_ORANGE = (255, 140, 0)
_COLOR_GOLD = (255, 215, 0)
_COLOR_DARK_ORCHID = (153, 50, 204)
_COLOR_PURPLE = (128, 0, 128)
_COLOR_SADDLE_BROWN = (139, 69, 19)
_COLOR_FOREST_GREEN = (34, 139, 34)
_COLOR_DODGER_BLUE = (30, 144, 255)
_COLOR_BLACK = (0, 0, 0)
_COLOR_WHITE = (255, 255, 255)
_COLOR_LIGHT_GRAY = (210, 210, 210)

# =========================================================
# 시스템 카테고리 고정 색상 테이블 (Head/WS/HN 전용)
# =========================================================
_SYSTEM_NOZZLE_COLORS: dict[str, tuple[int, int, int]] = {
    "head":        _COLOR_RED,
    "water_spray": _COLOR_FOREST_GREEN,
    "hose_nozzle": _COLOR_WHITE,
}

# =========================================================
# 사용자 정의 노즐 카테고리 자동 색상 팔레트 (12색, 고채도)
#
# 선정 기준: 예약색(Red/Blue/Green/Orange/Gold/Orchid/Purple/Brown/ForestGreen/DodgerBlue)의
# Hue ±15° 대역을 피해 Zone A(66°-100°), B(140°-195°), C(255°-272°), D(315°-345°) 에서 선정.
# =========================================================
_USER_NOZZLE_PALETTE: list[tuple[int, int, int]] = [
    (255,  20, 147),  # Deep Pink          H≈328° Zone D
    (  0, 206, 209),  # Dark Turquoise     H≈181° Zone B
    (154, 205,  50),  # Yellow Green       H≈ 80° Zone A
    (138,  43, 226),  # Blue Violet        H≈271° Zone C
    (  0, 255, 127),  # Spring Green       H≈150° Zone B
    (255, 105, 180),  # Hot Pink           H≈330° Zone D
    ( 64, 224, 208),  # Turquoise          H≈174° Zone B
    (  0, 250, 154),  # Medium Spring Grn  H≈157° Zone B
    (110,  20, 220),  # Purple Blue        H≈267° Zone C
    (221, 255,   0),  # Chartreuse Yellow  H≈ 68° Zone A
    (  0, 210, 180),  # Aquamarine         H≈171° Zone B
    (230,   0, 172),  # Fuchsia            H≈315° Zone D
]


def _auto_color_for_user_category(category_id: str) -> tuple[int, int, int]:
    """category_id의 MD5 해시로 팔레트에서 결정론적으로 색상을 선택한다.

    - 같은 category_id → 언제나 같은 색 (파일/세션 간 안정)
    - 보안 목적이 아니므로 MD5 사용 (표준 라이브러리, 플랫폼 무관)
    """
    digest = hashlib.md5(category_id.encode("utf-8")).hexdigest()
    idx = int(digest[:8], 16) % len(_USER_NOZZLE_PALETTE)
    return _USER_NOZZLE_PALETTE[idx]


@dataclass(slots=True)
class NodeVisual:
    """draw_all()에서 각 노드를 그릴 때 필요한 분류 정보."""
    color_rgb: tuple[int, int, int] = _COLOR_GRAY
    is_head: bool = False
    is_pump: bool = False
    is_valve: bool = False
    is_tank: bool = False
    is_prv: bool = False
    is_orifice: bool = False
    is_check_valve: bool = False
    is_hose_stream: bool = False


def classify_node_visual(
    tid: str | None,
    ntype_lower: str,
    is_active: bool = True,
    category_id: str = "",
    category_color_rgb: tuple[int, int, int] | None = None,
) -> NodeVisual:
    """노드의 type_id와 category_id로 시각화 분류를 결정한다.

    Parameters
    ----------
    tid : str | None
        node.type_id 또는 editor.get_node_type_id() 값
    ntype_lower : str
        노드의 type 문자열을 소문자로 정규화한 값 (폴백용, category_id 우선)
    is_active : bool
        node.is_active 값
    category_id : str
        노드의 라이브러리 카테고리 ID (예: "head", "water_spray", "alarm_valve")
    category_color_rgb : tuple[int, int, int] | None
        라이브러리에 저장된 사용자 카테고리 색상. 전달되면 최우선 사용.
        None이면 MD5 해시 기반 팔레트에서 결정론적으로 선택(폴백).

    Returns
    -------
    NodeVisual
        색상 및 카테고리 플래그
    """
    vis = NodeVisual()

    # 1. 펌프
    if tid == NodeTypeId.PUMP.value:
        vis.color_rgb = _COLOR_BLUE
        vis.is_pump = True
        return vis

    # 2. 헤드 / 노즐
    if tid in (NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value):
        vis.is_head = True
        _cat = category_id or ntype_lower  # category_id 우선, 없으면 ntype_lower 폴백
        if not is_active:
            vis.color_rgb = _COLOR_GRAY
        elif _cat in _SYSTEM_NOZZLE_COLORS:
            # 시스템 카테고리(head / water_spray / hose_nozzle) — 고정색
            vis.color_rgb = _SYSTEM_NOZZLE_COLORS[_cat]
        elif _cat and _cat not in ("nozzle", "head", "base"):
            # 사용자 정의 카테고리 — 라이브러리 저장 색 우선, 없으면 해시 폴백
            vis.color_rgb = category_color_rgb if category_color_rgb else _auto_color_for_user_category(_cat)
        else:
            # Head 스펙명("80(5.6)") 또는 폴백 — 헤드 고정색
            vis.color_rgb = _COLOR_RED
        return vis

    # 3. 수조 / 상수원
    if tid in (NodeTypeId.WT.value, NodeTypeId.RES.value):
        vis.color_rgb = _COLOR_DARK_GREEN
        vis.is_tank = True
        return vis

    # 4. PRV
    if tid == NodeTypeId.PRV.value:
        vis.color_rgb = _COLOR_ORANGE if is_active else _COLOR_GRAY
        vis.is_valve = True
        vis.is_prv = True
        return vis

    # 5. 오리피스
    if tid == NodeTypeId.ORIFICE.value:
        vis.color_rgb = _COLOR_GRAY if not is_active else _COLOR_GOLD
        vis.is_valve = True
        vis.is_orifice = True
        return vis

    # 6. 밸브류 — category_id 우선, ntype_lower 폴백
    if tid == NodeTypeId.VALVE.value:
        vis.is_valve = True
        _cat = category_id or ntype_lower
        if _cat in (
            "alarm_valve",
            "alarm valve",
            "preaction_valve",
            "preaction valve",
            "dry_pipe_valve",
            "dry pipe valve",
            "flow_switch_vane",
            "flow switch vane",
        ):
            vis.color_rgb = _COLOR_DARK_ORCHID
        elif (
            (category := fitting_category_from_id(_cat)) is not None
            and is_check_valve_category(category)
        ):
            vis.color_rgb = _COLOR_PURPLE
            vis.is_check_valve = True
        elif _cat in ("gate_valve", "gate valve", "butterfly_valve", "butterfly valve"):
            vis.color_rgb = _COLOR_SADDLE_BROWN
        else:
            # Globe Valve, Angle Valve, Ball Valve 등 기타 밸브
            vis.color_rgb = _COLOR_SADDLE_BROWN
        return vis

    # 7. 호스스트림 (HOSE_STREAM) — 호스노즐과 같은 원형 계열, 내부만 연회색
    if tid == NodeTypeId.HOSE_STREAM.value:
        vis.is_head = True
        vis.is_hose_stream = True
        vis.color_rgb = _COLOR_LIGHT_GRAY if is_active else _COLOR_GRAY
        return vis  # if/if 구조이므로 return 필수

    # 8. 기본
    vis.color_rgb = _COLOR_GRAY
    return vis
