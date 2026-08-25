"""체크밸브 카테고리 해석 정책.

라이브러리 부속 category를 우선하고, 유효한 노드 category_id를 폴백으로
사용한다. 둘 다 해석할 수 없을 때만 레거시 is_check_valve를 보존한다.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Callable, Literal

from domain.constants import (
    FittingCategory,
    fitting_category_from_id,
    is_check_valve_category,
)

logger = logging.getLogger(__name__)

FittingLookup = Callable[[str], Any | None]
ResolutionSource = Literal["fitting", "category_id", "unresolved"]


@dataclass(frozen=True)
class CheckValveIdentity:
    """노드의 canonical category와 파생 체크밸브 상태."""

    category: FittingCategory | None
    category_id: str
    is_check_valve: bool
    source: ResolutionSource

    @property
    def resolved(self) -> bool:
        return self.category is not None


def _category_from_fitting(fitting: Any | None) -> FittingCategory | None:
    if fitting is None:
        return None
    raw_category = getattr(fitting, "category", None)
    if isinstance(raw_category, FittingCategory):
        return raw_category
    return fitting_category_from_id(str(raw_category or ""))


def resolve_check_valve_identity(
    *,
    category_id: str,
    fitting_id: str | None,
    legacy_is_check_valve: bool,
    fitting_lookup: FittingLookup | None,
    node_id: str = "",
    warn_unresolved: bool = False,
) -> CheckValveIdentity:
    """부속 참조 → 노드 category 순으로 체크밸브 정체성을 해석한다.

    판정 불가일 때는 category를 조작하지 않고 레거시 플래그를 보존한다.
    """
    fitting_category = None
    if fitting_id and fitting_lookup is not None:
        fitting_category = _category_from_fitting(fitting_lookup(str(fitting_id)))
    if fitting_category is not None:
        return CheckValveIdentity(
            category=fitting_category,
            category_id=fitting_category.value,
            is_check_valve=is_check_valve_category(fitting_category),
            source="fitting",
        )

    stored_category = fitting_category_from_id(category_id)
    if stored_category is not None:
        return CheckValveIdentity(
            category=stored_category,
            category_id=stored_category.value,
            is_check_valve=is_check_valve_category(stored_category),
            source="category_id",
        )

    if warn_unresolved and (fitting_id or category_id or legacy_is_check_valve):
        logger.warning(
            "[CHECK_VALVE] category 판정 불가 — 레거시 플래그 보존 "
            "(node=%s, fitting_id=%s, category_id=%s, is_check_valve=%s)",
            node_id or "?",
            fitting_id or "",
            category_id or "",
            bool(legacy_is_check_valve),
        )
    return CheckValveIdentity(
        category=None,
        category_id=str(category_id or ""),
        is_check_valve=bool(legacy_is_check_valve),
        source="unresolved",
    )


def resolve_node_check_valve_identity(
    node: Any,
    fitting_lookup: FittingLookup | None,
    *,
    warn_unresolved: bool = False,
) -> CheckValveIdentity:
    return resolve_check_valve_identity(
        category_id=str(getattr(node, "category_id", "") or ""),
        fitting_id=getattr(node, "fitting_id", None),
        legacy_is_check_valve=bool(getattr(node, "is_check_valve", False)),
        fitting_lookup=fitting_lookup,
        node_id=str(getattr(node, "id", "") or ""),
        warn_unresolved=warn_unresolved,
    )


def canonical_check_valve_state(
    node: Any,
    fitting_lookup: FittingLookup | None,
    *,
    warn_unresolved: bool = False,
) -> dict[str, object]:
    """solver DTO에 넣을 canonical category와 파생 플래그."""
    if str(getattr(node, "type_id", "") or "").strip().lower() != "valve":
        return {
            "category_id": str(getattr(node, "category_id", "") or ""),
            "is_check_valve": False,
        }

    identity = resolve_node_check_valve_identity(
        node,
        fitting_lookup,
        warn_unresolved=warn_unresolved,
    )
    return {
        "category_id": identity.category_id,
        "is_check_valve": identity.is_check_valve,
    }


__all__ = [
    "CheckValveIdentity",
    "canonical_check_valve_state",
    "resolve_check_valve_identity",
    "resolve_node_check_valve_identity",
]
