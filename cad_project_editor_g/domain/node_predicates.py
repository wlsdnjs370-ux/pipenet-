# 파일명 domain/node_predicates.py
"""노드 타입 capability predicate — 흩어진 타입 분기의 단일 도메인 소유 지점.

Residual type-branch follow-up slice R1. 새 기기 타입이 추가될 때
호출부가 아니라 이 모듈만 수정하면 되도록 분기 결정을 한 곳으로 모은다.

⚠ 의미 보존 캡슐화 단계: 각 predicate의 내부 로직은 기존 인라인 패턴을
그대로 재현한다(예: pump 감지는 display 문자열 기반). type_id/ComponentSpec
기반으로 의미를 강화하는 것은 별도 decision slice(R4)에서 동치성 검증 후
진행한다. 상세는 docs/adr/current-component-spec/residual-type-branch-followup.md.
"""
from __future__ import annotations

from typing import Any


def is_pump_type_text(type_text: Any) -> bool:
    """display 타입 문자열이 펌프를 가리키는지 — 기존 `"pump" in str(t).lower()` 재현."""
    return "pump" in str(type_text if type_text is not None else "").lower()


def is_pump_node(node: Any) -> bool:
    """노드가 펌프인지 — 기존 `"pump" in str(getattr(node, "type", "")).lower()` 재현.

    None 입력은 False(호출부의 `node is not None and ...` 가드를 흡수).
    """
    if node is None:
        return False
    return is_pump_type_text(getattr(node, "type", ""))


def is_pump_graph_node(node: Any) -> bool:
    """펌프 노드 — type/type_id/category_id 종합(재생·3D 라벨 SSOT)."""
    if is_pump_node(node):
        return True
    tid = str(getattr(node, "type_id", "") or "").lower()
    cat = str(getattr(node, "category_id", "") or "").lower()
    type_text = str(getattr(node, "type", "") or "").lower()
    return tid == "pump" or "pump" in tid or "pump" in cat or "pump" in type_text


def is_tank_graph_node(node: Any) -> bool:
    """수조 노드 — type/type_id/category_id 종합(재생·3D 라벨 SSOT)."""
    if node is None:
        return False
    tid = str(getattr(node, "type_id", "") or "").lower()
    cat = str(getattr(node, "category_id", "") or "").lower()
    type_text = str(getattr(node, "type", "") or "")
    return (
        is_water_tank_type_text(type_text)
        or tid in ("wt", "tank", "reservoir", "res")
        or "tank" in cat
        or "wt" in cat
    )


def is_water_tank_type_text(type_text: Any) -> bool:
    """display 타입 문자열이 수조(WT/Tank)를 가리키는지 —
    기존 `"wt" in str(t).lower() or "tank" in str(t).lower()` 재현."""
    low = str(type_text if type_text is not None else "").lower()
    return "wt" in low or "tank" in low


# ── 방수(discharge) 계열: head / nozzle ──────────────────────────────
# NodeTypeId 실제 값: "head", "nozzle" (water_spray/hose_nozzle 카테고리는
# 모두 type_id "nozzle"에 매핑). type_id 비교는 방어적 소문자화 형태로 통일
# — 기존 호출부 중 일부(view_graphics, node_controller 일괄변경)는 이미
# str().lower() 후 비교했고, 일부(dialog_flow_distribution)는 직비교였다.
# 정규화된 type_id는 항상 소문자이므로 두 형태는 실데이터에서 동치이며,
# 소문자화 통일은 containment-neutral 확장이다.


def is_head_type_id(type_id: Any) -> bool:
    """type_id가 head인지 — 기존 `str(tid or "").lower() == "head"` 계열 재현."""
    return str(type_id if type_id is not None else "").lower() == "head"


def is_nozzle_type_id(type_id: Any) -> bool:
    """type_id가 nozzle인지 — head와 동일한 방어적 소문자화 비교."""
    return str(type_id if type_id is not None else "").lower() == "nozzle"


def is_discharge_type_id(type_id: Any) -> bool:
    """type_id가 방수형(head|nozzle)인지 —
    기존 `tid not in ("head", "nozzle")` 멤버십 검사들의 긍정형."""
    return is_head_type_id(type_id) or is_nozzle_type_id(type_id)


def is_discharge_node(node: Any) -> bool:
    """노드가 방수형(head|nozzle)인지 — type_id 기반.

    None 입력/`type_id` 속성 부재는 False(호출부의
    `getattr(node, "type_id", "base") or "base"` 폴백을 흡수).
    """
    if node is None:
        return False
    return is_discharge_type_id(getattr(node, "type_id", None))


def is_head_type_text(type_text: Any) -> bool:
    """display 타입 문자열이 정확히 head인지 —
    기존 `str(node.type).lower() == "head"` 재현(None은 "none"이 되어 False)."""
    return str(type_text).lower() == "head"


def is_base_type_id(type_id: Any) -> bool:
    """type_id가 기본(base) 노드인지 — 기존 `type_id in ("base",)` 직비교 재현.

    의도적으로 소문자화하지 않는다: 기존 호출부(병합삭제 확인창)는 비정규
    type_id를 '특수 노드'로 취급해 확인창을 띄우는 보수적 동작이었고,
    이를 그대로 보존한다.
    """
    return type_id == "base"
