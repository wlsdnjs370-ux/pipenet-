# -*- coding: utf-8 -*-
"""core/auto_design.py — 자동 구역계획 · 헤드배치 · 배관라우팅 테스트.

파일 I/O 가 전혀 없는 순수 계산 모듈이라 전 구간을 단위 테스트한다.
헤드 배치는 (2R)² ≥ S² + L² 과 skipping 1.8 m 두 하드 체크가 핵심이고,
라우팅은 가지당 헤드 수 상한과 구경 사다리가 핵심이다.

실행::

    python -m pytest tests/test_auto_design.py -q
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import auto_design as A  # noqa: E402
from phd_rules import PressureZone  # noqa: E402


def _square(side: float) -> list[tuple[float, float]]:
    return [(0.0, 0.0), (side, 0.0), (side, side), (0.0, side)]


def _room(**kw):
    side = kw.pop("side", 8.0)
    d = {
        "floor": "F1",
        "area_m2": side * side,
        "estimated_head_count": 7,
        "use": "other_low",
        "structure": "non_fire_resistant",
        "ceiling_h_m": 3.0,
        "polygon": _square(side),
    }
    d.update(kw)
    return d


def _zone(**kw):
    return A.AutoZonePlanner({"height_m": 30.0}).plan([_room(**kw)])[0]


# ────────────────────────────────────────────────────────────────────────────
# ② 구역계획 + 헤드 사양
# ────────────────────────────────────────────────────────────────────────────

def test_excluded_rooms_produce_no_zone():
    """NFTC 2.12 / HB §2.4.18 제외실은 구역 자체가 생기면 안 된다."""
    zones = A.AutoZonePlanner({}).plan([_room(hb_excluded=True), _room(floor="F2")])
    assert [z.floor_label for z in zones] == ["F2"]


def test_large_room_splits_into_multiple_zones():
    zones = A.AutoZonePlanner({}).plan([_room(area_m2=7000.0)])
    assert len(zones) == 3
    assert all(z.head_spec is not None for z in zones)


def test_head_spec_r_comes_from_nftc_table():
    assert _zone(use="apartment_living").head_spec.horizontal_distance_m == 3.2
    assert _zone(use="other_low", structure="fire_resistant").head_spec \
        .horizontal_distance_m == 2.3


def test_special_combustible_shrinks_r_to_1_7():
    z = _zone(use="factory", has_special_combustible=True)
    assert z.head_spec.horizontal_distance_m == 1.7


def test_fast_response_mandate_sets_rti_and_flag():
    z = _zone(use="apartment_living")
    assert z.nftc_2755_fast_required is True
    assert z.head_spec.rti_class == "fast"
    assert _zone(use="factory").head_spec.rti_class == "standard"


def test_dry_system_pendent_becomes_drypendent():
    """건식에 일반 하향형을 달면 배관에 물이 고여 동파한다."""
    z = _zone(system_type_hint="dry", head_orientation="pendent")
    assert z.head_spec.head_type == "drypendent"
    wet = _zone(system_type_hint="wet", head_orientation="pendent")
    assert wet.head_spec.head_type == "pendent"


def test_esfr_activated_for_tall_warehouse():
    z = _zone(use="warehouse", ceiling_h_m=12.0)
    assert z.head_spec.is_esfr is True
    assert z.head_spec.k_factor_lpm_bar05 >= 200


def test_ordinary_room_defaults_to_k80():
    assert _zone().head_spec.k_factor_lpm_bar05 == 80


def test_head_spec_trace_records_all_four_clauses():
    tr = _zone(use="warehouse", ceiling_h_m=12.0).head_spec.trace
    assert "2.7.3" in tr.nftc and "2.7.6" in tr.nftc


# ────────────────────────────────────────────────────────────────────────────
# ③ 헤드 자동 배치
# ────────────────────────────────────────────────────────────────────────────

def test_placement_satisfies_a_b_l_and_skipping():
    """(2R)² ≥ S² + L² 과 헤드 간격 1.8 m — 둘 다 어기면 살수 공백이 생긴다."""
    z = _zone()
    heads = A.AutoHeadPlacer(zone=z).place()
    assert heads
    r = z.head_spec.horizontal_distance_m
    for h in heads:
        assert h.a_b_l_check is True
        assert h.skipping_pass is True
        assert (2 * r) ** 2 >= h.cell_S ** 2 + h.cell_L ** 2


def test_placed_heads_are_at_least_1_8m_apart():
    heads = A.AutoHeadPlacer(zone=_zone()).place()
    for i, a in enumerate(heads):
        for b in heads[i + 1:]:
            assert math.hypot(a.x - b.x, a.y - b.y) >= A.AutoHeadPlacer.HEAD_SPACING_MIN_M


def test_all_heads_land_inside_the_zone_polygon():
    side = 8.0
    heads = A.AutoHeadPlacer(zone=_zone(side=side)).place()
    assert all(0.0 <= h.x <= side and 0.0 <= h.y <= side for h in heads)


def test_head_ids_are_unique_and_zone_scoped():
    z = _zone()
    heads = A.AutoHeadPlacer(zone=z).place()
    ids = [h.head_id for h in heads]
    assert len(ids) == len(set(ids))
    assert all(i.startswith(z.zone_id) for i in ids)


def test_larger_zone_needs_more_heads():
    small = len(A.AutoHeadPlacer(zone=_zone(side=6.0)).place())
    large = len(A.AutoHeadPlacer(zone=_zone(side=12.0)).place())
    assert large > small


def test_zone_without_head_spec_places_nothing():
    z = _zone()
    z.head_spec = None
    assert A.AutoHeadPlacer(zone=z).place() == []


def test_head_inside_obstacle_is_not_silently_clear():
    """장애물 안쪽 헤드는 변까지 거리가 멀다는 이유로 통과하면 안 된다.

    nftc_rules.validate_head_clearance 는 내부 헤드를 거리 0.0 으로 잡는데,
    auto_design 의 자체 거리 계산에는 그 처리가 빠져 있어 배치 단계에서만
    조용히 합격하던 결함이다.
    """
    placer = A.AutoHeadPlacer(zone=_zone(), obstacles=[
        {"id": "duct", "polygon": [(3.0, 3.0), (5.0, 3.0), (5.0, 5.0), (3.0, 5.0)]}])
    assert placer._min_dist_to_polygon((4.0, 4.0), _square(8.0)) == 0.0
    assert placer._clearance_ok((4.0, 4.0)) is False
    assert placer._clearance_ok((0.5, 0.5)) is True


def test_fully_blocked_zone_places_nothing():
    """구역 전체가 장애물이면 위반 헤드를 두느니 0개를 고른다 (현재 비용함수 계약)."""
    z = _zone()
    blob = {"id": "duct", "polygon": _square(8.0)}
    assert A.AutoHeadPlacer(zone=z, obstacles=[blob]).place() == []


def test_wall_obstacle_uses_10cm_not_60cm():
    z = _zone()
    far = {"id": "w", "polygon": [(-0.5, -0.5), (-0.2, -0.5), (-0.2, 8.5), (-0.5, 8.5)],
           "is_wall": True}
    heads = A.AutoHeadPlacer(zone=z, obstacles=[far]).place()
    assert all(h.nftc_2771_pass for h in heads)


def test_empty_polygon_yields_no_heads():
    z = _zone()
    z.polygon = []
    assert A.AutoHeadPlacer(zone=z).place() == []


# ────────────────────────────────────────────────────────────────────────────
# ④ 배관 라우팅
# ────────────────────────────────────────────────────────────────────────────

def _heads_in_row(zone, n, *, y=0.0, step=2.5, axis="EW"):
    return [
        A.HeadInstance(
            head_id=f"{zone.zone_id}-H{i:03d}", zone_id=zone.zone_id,
            x=i * step, y=y, z=0.0, spec=zone.head_spec, branch_axis=axis,
            cell_S=step, cell_L=step, nftc_2773_pass=True, nftc_2771_pass=True,
            skipping_pass=True, a_b_l_check=True)
        for i in range(n)
    ]


def test_branch_splits_at_8_heads():
    """가지당 헤드 8개 상한 — 넘으면 말단 압력이 확보되지 않는다."""
    z = _zone()
    z.heads = _heads_in_row(z, 10)
    branches, _ = A.AutoPipeRouter(zone=z, riser_xy=(0.0, -5.0)).route()
    assert len(branches) == 2
    assert [b.head_count_downstream for b in branches] == [8, 2]


def test_each_row_becomes_its_own_branch():
    z = _zone()
    z.heads = _heads_in_row(z, 4, y=0.0) + _heads_in_row(z, 4, y=3.0)
    branches, _ = A.AutoPipeRouter(zone=z, riser_xy=(0.0, -5.0)).route()
    assert len(branches) == 2


def test_branch_nominal_grows_with_head_count():
    z = _zone()
    ladder = {}
    for n in (2, 4, 6, 8):
        z.heads = _heads_in_row(z, n)
        branches, _ = A.AutoPipeRouter(zone=z, riser_xy=(0.0, -5.0)).route()
        ladder[n] = branches[0].nominal
    assert ladder == {2: "25A", 4: "32A", 6: "40A", 8: "50A"}


def test_branch_inner_diameter_matches_material_table():
    z = _zone()
    z.heads = _heads_in_row(z, 4)
    branches, _ = A.AutoPipeRouter(zone=z, riser_xy=(0.0, -5.0)).route()
    b = branches[0]
    assert b.material == "KSD 3507"
    assert b.inner_diameter_mm == pytest.approx(36.2)   # 32A / 3507
    assert b.c_factor == 120


def test_branch_carries_tee_endcap_and_hangers():
    z = _zone()
    z.heads = _heads_in_row(z, 6)
    branches, _ = A.AutoPipeRouter(zone=z, riser_xy=(0.0, -5.0)).route()
    b = branches[0]
    assert [f["type"] for f in b.fittings] == ["tee_branch", "endcap"]
    assert b.hangers_m
    assert max(b.hangers_m) <= b.length_m


def test_cross_main_nominal_grows_with_branch_count():
    z = _zone()
    z.heads = sum((_heads_in_row(z, 2, y=3.0 * k) for k in range(8)), [])
    _, cms = A.AutoPipeRouter(zone=z, riser_xy=(0.0, -5.0)).route()
    assert len(cms) == 8
    assert {c.nominal for c in cms} == {"100A"}


def test_zero_length_cross_main_is_dropped():
    """라이저가 가지 시작점과 같은 자리면 길이 0 배관이 생긴다 — 버려야 한다."""
    z = _zone()
    z.heads = _heads_in_row(z, 4)
    _, cms = A.AutoPipeRouter(zone=z, riser_xy=(0.0, 0.0)).route()
    assert all(c.length_m > 0 for c in cms)
    assert len(cms) == 0


def test_auto_riser_defaults_to_branch_centroid():
    z = _zone()
    z.heads = _heads_in_row(z, 2, y=0.0) + _heads_in_row(z, 2, y=4.0)
    _, cms = A.AutoPipeRouter(zone=z).route()
    assert {round(c.n2[1], 3) for c in cms} == {2.0}


def test_routing_empty_zone_returns_empty_lists():
    z = _zone()
    z.heads = []
    assert A.AutoPipeRouter(zone=z).route() == ([], [])


# ────────────────────────────────────────────────────────────────────────────
# 부속류 배치
# ────────────────────────────────────────────────────────────────────────────

def test_no_riser_position_places_no_valves():
    assert A.AutoFittingPlacer.place_for_zone(_zone(), system_type="wet") == []


@pytest.mark.parametrize("system_type, vtype, eq_len", [
    ("wet", "alarm", 12.9),
    ("preaction_single", "preaction", 10.1),
    ("dry", "dry", 10.1),
    ("deluge", "deluge", 10.1),
])
def test_system_type_selects_its_valve(system_type, vtype, eq_len):
    valves = A.AutoFittingPlacer.place_for_zone(
        _zone(), system_type=system_type, riser_position=(0.0, 0.0, 0.0))
    main = valves[0]
    assert main.type == vtype
    assert main.equivalent_length_m == eq_len


def test_osy_is_always_present():
    for st in ("wet", "dry", "deluge", "preaction_double"):
        valves = A.AutoFittingPlacer.place_for_zone(
            _zone(), system_type=st, riser_position=(0.0, 0.0, 0.0))
        assert any(v.type == "os_y" for v in valves)


def test_lsp_zone_gets_prv_at_4bar():
    valves = A.AutoFittingPlacer.place_for_zone(
        _zone(), system_type="wet", is_lsp=True, riser_position=(0.0, 0.0, 0.0))
    prv = next(v for v in valves if v.type == "prv")
    assert prv.spec["p2_bar"] == 4.0
    assert prv.spec["p1_bar"] is None   # 실측 전에는 미상 — 0.0 으로 채우지 않는다


def test_non_lsp_zone_gets_no_prv():
    valves = A.AutoFittingPlacer.place_for_zone(
        _zone(), system_type="wet", riser_position=(0.0, 0.0, 0.0))
    assert all(v.type != "prv" for v in valves)


# ────────────────────────────────────────────────────────────────────────────
# 전체 결선
# ────────────────────────────────────────────────────────────────────────────

def _network(**kw):
    params = dict(
        project_id="T-001",
        rooms=[_room(floor="F1"), _room(floor="F2")],
        obstacles=[],
        floors=[{"label": "F1", "z_m": 0.0}, {"label": "F2", "z_m": 3.5}],
        building_meta={"height_m": 30.0, "elevated_tank_z_m": 30.0},
    )
    params.update(kw)
    return A.design_full_network(**params)


def test_full_network_wires_all_stages():
    net = _network()
    assert net.project_id == "T-001"
    assert len(net.zones) == 2
    assert all(z.heads and z.branches and z.valves for z in net.zones)
    assert net.discretionary is not None
    assert net.hb_case is not None


def test_full_network_tags_zones_with_pressure_zone():
    net = _network()
    assert all(z.pressure_zone is not None for z in net.zones)


def test_freezing_risk_propagates_to_system_and_valves():
    net = _network(building_meta={"height_m": 30.0, "has_freezing_risk": True})
    assert net.system_type == "dry"
    assert any(v.type == "dry" for z in net.zones for v in z.valves)


def test_pump_record_matches_hb_case_location():
    net = _network()
    assert net.pumps[0]["location"] == net.hb_case.pump_location


def test_network_metadata_carries_decision_traces():
    net = _network()
    assert "system_decision_trace" in net.metadata
    assert "hb_case_trace" in net.metadata


def test_zone_centroid_of_empty_polygon_is_origin():
    assert A._zone_centroid([]) == (0.0, 0.0)


def test_zone_centroid_of_square():
    assert A._zone_centroid(_square(4.0)) == (2.0, 2.0)


def test_all_exports_exist():
    for name in A.__all__:
        assert hasattr(A, name), name
