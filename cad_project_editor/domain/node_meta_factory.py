"""속성적용 메타 팩토리 — display명 → flat meta 변환 체인의 단일 소유 지점 (R3a).

`NodeController.handle_node_attribute_apply`에 흩어져 있던 if/elif 타입 체인을
의미 보존 그대로 옮긴 것이다. 생성된 meta는 `update_node_via_graph`를 거쳐
`.kfp`에 그대로 저장되므로 출력은 추출 전과 byte-identical해야 한다
(characterization: tests/test_node_attribute_meta_factory.py).

새 기기 타입의 속성적용 meta는 이 모듈 한 곳에만 추가하면 된다.

도메인 계층 규칙: 라이브러리 조회(librarian)와 카테고리 정규화 함수는 호출자가
주입한다 — services/editor를 import하지 않는다.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Callable, Iterable, Protocol

from domain.constants import is_check_valve_category
from domain.models import _NODE_TYPE_LOOKUP


class NodeMetaLibraryLookup(Protocol):
    """팩토리가 필요로 하는 읽기 전용 라이브러리 조회 표면 (EditorCore가 충족)."""

    def find_nozzle_category_by_display(self, display_name: str) -> dict | None: ...

    def get_head_spec(self, head_name: str) -> dict | None: ...

    def get_nozzle_items_by_category(self, category_id: str) -> list[dict]: ...

    def get_fittings_v3_flat(self) -> list[Any]: ...


@dataclass
class AttributeApplyMeta:
    """속성적용 meta 생성 결과.

    Fields:
        meta          : update_node_via_graph에 전달할 flat meta dict
        type_id       : meta의 type_id (없으면 "base")
        reset_to_base : 기본 노드 분기 — 호출자가 update 전에
                        reset_node_to_base(name, "기본")를 수행해야 함
    """

    meta: dict[str, Any]
    type_id: str
    reset_to_base: bool = False


def build_attribute_apply_meta(
    base_meta: dict[str, Any],
    real_attr: str,
    *,
    library: NodeMetaLibraryLookup,
    canonicalize_category_id: Callable[[str, str], str],
) -> AttributeApplyMeta:
    """노드 meta 스냅샷(base_meta)과 표시명(real_attr)으로 속성적용 meta를 만든다.

    base_meta는 사전 reset된 노드의 `to_dict()` 결과를 그대로 받아
    in-place로 가공한다(추출 전 체인과 동일). 'Head' → 설계 기본 스펙명 해석은
    UI 잔류분이므로 real_attr는 이미 해석된 값이어야 한다.
    """
    meta = base_meta
    meta.pop("id", None)
    meta.pop("coords", None)
    meta.pop("elevation_m", None)

    meta["type"] = real_attr  # 임시: 스펙명, 아래에서 카테고리명으로 덮어씀

    low_attr = real_attr.lower()
    reset_to_base = False

    # ── SSOT 분기: JSON 라이브러리 조회 우선 ──────────────
    # 1순위: 노즐 라이브러리 — 카테고리 레벨 검색
    nozzle_cat = library.find_nozzle_category_by_display(real_attr)

    # Head 특수 케이스: real_attr이 항목명("80(5.6)")일 수 있음
    # (호출자에서 "Head" → 설계기본값 스펙명으로 변환되기 때문)
    head_spec = None
    if not nozzle_cat:
        head_spec = library.get_head_spec(real_attr)
        if head_spec:
            nozzle_cat = library.find_nozzle_category_by_display("Head")

    if nozzle_cat:
        cat_type_id = nozzle_cat.get("type_id", "nozzle")
        meta["is_active"] = True

        if cat_type_id == "head":
            # Head 카테고리
            meta["type_id"] = "head"
            meta["category_id"] = canonicalize_category_id(
                nozzle_cat.get("category_id", "head"),
                nozzle_cat.get("display_name", ""),
            )
            meta["type"] = "Head"
            meta["head_spec_name"] = real_attr
            if head_spec:
                k_val = float(head_spec.get("K_SI", 80.0))
                meta["k_factor_si"] = k_val
                meta["k_factor"] = k_val
        else:
            # Nozzle / Hose Nozzle 카테고리
            meta["type_id"] = "nozzle"
            cat_id = canonicalize_category_id(
                nozzle_cat.get("category_id", "nozzle"),
                nozzle_cat.get("display_name", ""),
            )
            meta["category_id"] = cat_id
            items = library.get_nozzle_items_by_category(cat_id)
            if items and not meta.get("nozzle_name"):
                default_nz = items[0]
                meta["nozzle_name"] = default_nz.get("display_name") or default_nz.get("name")
                meta["required_pressure_bar"] = float(default_nz.get("min_p", 1.0))
                k_val = float(default_nz.get("K_val", 0.0))
                meta["k_factor"] = k_val
                meta["k_factor_si"] = k_val

    # 2순위: v3 카테고리 SSOT에서 display_name → category_id 조회
    elif canonicalize_category_id("", real_attr):
        cat_id = canonicalize_category_id("", real_attr)
        meta.update({
            "type_id": "valve",
            "is_active": True,
            "category_id": cat_id,
        })
        valve_items = [
            item for item in library.get_fittings_v3_flat()
            if item.category.value == cat_id
        ]
        if valve_items:
            meta["fitting_id"] = valve_items[0].id
            meta["is_check_valve"] = is_check_valve_category(valve_items[0].category)

    # 3순위: 특수 노드 — 룩업 우선, 부분 매칭 폴백 (§1.6)
    elif _NODE_TYPE_LOOKUP.get(low_attr) == "pump" or "pump" in low_attr:
        meta.update({
            "type_id": "pump",
            "category_id": "pump",
            "is_active": True,
            "rated_q": 1000.0,
            "rated_p": 10.0,
            "shutoff_p": 12.0,
            "peak_q": 1500.0,
            "peak_p": 6.5,
        })

    elif _NODE_TYPE_LOOKUP.get(low_attr) == "wt" or "wt" in low_attr or "tank" in low_attr:
        meta.update({
            "type_id": "wt",
            "category_id": "wt",
            "is_active": True,
            "water_level": 2.0,
        })

    elif _NODE_TYPE_LOOKUP.get(low_attr) == "prv" or "prv" in low_attr:
        meta.update({
            "type_id": "prv",
            "category_id": "prv",
            "is_active": True,
            "pressure_setting_bar": 4.5,
            "loss_coefficient": 5.0,
        })

    elif _NODE_TYPE_LOOKUP.get(low_attr) == "orifice" or "orif" in low_attr:
        meta.update({
            "type_id": "orifice",
            "category_id": "orifice",
            "is_active": True,
            "hole_diameter": 25.0,
            "orifice_discharge_coeff": 0.65,
        })

    elif _NODE_TYPE_LOOKUP.get(low_attr) == "reservoir":
        meta.update({
            "type_id": "reservoir",
            "category_id": "reservoir",
            "is_active": True,
        })

    elif _NODE_TYPE_LOOKUP.get(low_attr) == "hose_stream":
        meta.update({
            "type_id": "hose_stream",
            "category_id": "hose_stream",
            "type": "Hose Stream",
            "is_active": True,
            "hose_stream_flow_lps": 400.0 / 60.0,  # 기본값 400 L/min
        })

    # 4순위: 기본 노드
    else:
        reset_to_base = True
        meta["type_id"] = "base"
        meta["category_id"] = ""
        meta["is_active"] = True
        # 잔존 펌프 메타 명시 삭제 — facade reset + normalize 정책의 보조 가드
        for _stale_key in (
            "pump_curve_data",
            "pump_library_id",
            "rated_q",
            "rated_p",
            "shutoff_p",
            "peak_q",
            "peak_p",
        ):
            meta.pop(_stale_key, None)

    # 안전장치: head_spec 조회 성공했는데 meta["type"]이 여전히 스펙명이면 카테고리명으로 교정
    if head_spec and meta.get("type_id") == "head" and meta.get("type") != "Head":
        meta["type"] = "Head"

    return AttributeApplyMeta(
        meta=meta,
        type_id=meta.get("type_id", "base"),
        reset_to_base=reset_to_base,
    )


# ──────────────────────────────────────────────────────────
# 계산 직전 라이브러리 동기화 meta (R3b)
# run_analysis/run_inverse/run_unified UseCase ×3 + CalcController에
# 중복돼 있던 _sync_library_to_nodes의 타입 체인을 단일 소유로 통합.
# ──────────────────────────────────────────────────────────
class LibrarySyncLookup(Protocol):
    """동기화 수집이 필요로 하는 읽기 전용 스펙 조회 표면 (EditorCore가 충족)."""

    def get_nozzle_spec(self, nozzle_name: str) -> dict | None: ...

    def get_head_spec(self, head_name: str) -> dict | None: ...


def nozzle_library_sync_meta(spec: dict[str, Any]) -> dict[str, Any]:
    """노즐 스펙 → 동기화 meta_update (k_factor_si + required_pressure_bar)."""
    meta_update: dict[str, Any] = {}
    k_val = spec.get("K_val")
    if spec.get("type") == "CALC":
        try:
            p_ref = float(spec.get("P_bar", 0.0))
            q_ref = float(spec.get("Q_lpm", 0.0))
        except (TypeError, ValueError):
            p_ref, q_ref = 0.0, 0.0
        if p_ref > 0:
            k_val = q_ref / math.sqrt(p_ref)
    if k_val is not None:
        try:
            meta_update["k_factor_si"] = float(k_val)
        except (TypeError, ValueError):
            pass
    try:
        _min_p = float(spec.get("min_p", 0.0))
    except (TypeError, ValueError):
        _min_p = 0.0
    if _min_p > 0:
        meta_update["required_pressure_bar"] = _min_p
    return meta_update


def head_library_sync_meta(
    spec: dict[str, Any], *, current_required_pressure_bar: float
) -> dict[str, Any]:
    """헤드 스펙 → 동기화 meta_update (k_factor_si만 — required는 전역 설정 담당)."""
    meta_update: dict[str, Any] = {}
    k_si = spec.get("K_SI")
    if k_si is not None:
        try:
            meta_update["k_factor_si"] = float(k_si)
        except (TypeError, ValueError):
            pass
    if current_required_pressure_bar > 0:
        meta_update["required_pressure_bar"] = 0.0
    return meta_update


def collect_library_sync_updates(
    nodes: dict[str, Any] | Iterable[tuple[str, Any]],
    library: LibrarySyncLookup,
) -> list[tuple[str, dict[str, Any]]]:
    """계산 직전 라이브러리 최신값 동기화 — 노드별 meta_update를 수집한다.

    적용(update_node_via_graph)은 호출자 책임. 타입 판정은 추출 전과 동일하게
    type_id 직비교(`tid == "nozzle"`/`"head"`)를 보존한다(의미 보존 캡슐화).
    """
    items = nodes.items() if hasattr(nodes, "items") else nodes
    updates: list[tuple[str, dict[str, Any]]] = []
    for nname, nnode in items:
        tid = getattr(nnode, "type_id", "")

        # (B) 노즐: k_factor_si + required_pressure_bar (min_p)
        if tid == "nozzle" and getattr(nnode, "nozzle_name", None):
            spec = library.get_nozzle_spec(nnode.nozzle_name)
            if spec:
                meta_update = nozzle_library_sync_meta(spec)
                if meta_update:
                    updates.append((nname, meta_update))

        # (C) 헤드: k_factor_si만 갱신 (required_pressure_bar는 전역 설정이 담당)
        elif tid == "head" and getattr(nnode, "head_spec_name", None):
            spec = library.get_head_spec(nnode.head_spec_name)
            if spec:
                meta_update = head_library_sync_meta(
                    spec,
                    current_required_pressure_bar=float(
                        getattr(nnode, "required_pressure_bar", 0.0)
                    ),
                )
                if meta_update:
                    updates.append((nname, meta_update))
    return updates
