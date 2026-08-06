# -*- coding: utf-8 -*-
"""지시서 §9.2·§9.3·§13.4 — 가지배관 골격과 교차배관 경로 (R1~R4).

지키는 것은 셋이다. **8개는 분기점 기준 한쪽**이라는 것(전체로 세어 자르면 적법한
16개짜리 가지배관이 둘로 쪼개진다), **최적화가 위법 형상을 만들지 않는다**는
것(총연장을 줄이려고 대칭 분기로 수렴하면 토너먼트가 된다), 그리고 **닿지 못한
분기점을 조용히 빼지 않는다**는 것(빼면 수리계산만 통과하고 미방호 구역이 남는다).
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.deterministic import pipe_routing as PR  # noqa: E402
from core.design.schema import BuildingDraft, floor_index  # noqa: E402

_SPACING = 3.2      # m — 헤드 간격. 라인 묶음 허용 오차는 이 값의 1/4 이다.
_STEP = _SPACING * 1000.0


class _Limits:
    """라우팅이 실제로 보는 값만. 전 필드를 되살리면 그 코드가 두 번째 진실이 된다."""

    def __init__(self, per_side=8, spacing=_SPACING):
        self.branch_heads_per_side_max = per_side
        self.head_spacing_square_m = spacing


def _room(heads, *, room_id="R-1F-001", axis_deg=0.0):
    return {"room_id": room_id, "metrics": {"axis_deg": axis_deg},
            "heads": [{"id": f"H-{i:03d}", "room_id": room_id, "x": x, "y": y,
                       "row": 0, "col": i} for i, (x, y) in enumerate(heads, 1)]}


def _grid(cols, rows):
    """cols 열 × rows 행. 열이 많을수록 교차배관이 x 를 따라 눕고, 그때 한 열이
    가지배관 하나가 된다 — 한쪽 8개 상한을 보려면 이 모양이어야 한다."""
    return [(c * _STEP, r * _STEP) for c in range(cols) for r in range(rows)]


def _rot(points, deg):
    t = math.radians(deg)
    cos_t, sin_t = math.cos(t), math.sin(t)
    return [(x * cos_t - y * sin_t, x * sin_t + y * cos_t) for x, y in points]


def _ys(heads):
    return {h["id"]: h["y"] for h in _room(heads)["heads"]}


def _plan(heads, *, per_side=8, valve=None, rooms=None):
    return PR.plan_branches("Z-1F-01", rooms or [_room(heads)],
                            valve, _Limits(per_side))


def _rect(x0, y0, x1, y1):
    return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]


def _floor(*polygons):
    """C530 이 보는 실은 폴리곤만 있으면 된다 — 벽은 그 경계에서 유도된다."""
    return [{"room_id": f"R-{n}", "polygon": list(p)}
            for n, p in enumerate(polygons, 1)]


def _routed(plan, rooms, *, cores=(), doors=(), min_dn=40):
    field = PR.RouteField(plan.axes, rooms, cores=cores, doors=doors)
    return PR.route_cross_mains(plan, field, min_dn=min_dn), field


def _ab(plan, path):
    return [plan.axes.project(p[0], p[1]) for p in path]


# ── C510 축 결정 ────────────────────────────────────────────────────────

def test_교차배관은_장변을_따라_눕는다():
    """가지배관은 짧고 많아야 한쪽 8개 상한에 여유가 생긴다(§9.2 C510)."""
    heads = [(c * _STEP, r * _STEP) for c in range(6) for r in range(2)]
    axes = PR.zone_axes([_room(heads)])
    assert axes.cross_span_mm > axes.branch_span_mm
    assert round(axes.to_dict()["cross_axis_deg"], 1) == 0.0


def test_복도형은_밸브를_봐도_축을_뒤집지_않는다():
    """장단변비 4 이상. 복도에서 가지배관이 장변을 따라가면 실 안으로 못 들어간다."""
    heads = [(c * _STEP, 0.0) for c in range(10)] + [(c * _STEP, _STEP) for c in range(10)]
    axes = PR.zone_axes([_room(heads)], valve_point=(5 * _STEP, -50 * _STEP))
    assert axes.corridor is True and axes.flipped is False


def test_밸브가_옆에_붙으면_교차배관이_그쪽으로_돈다():
    """교차배관은 밸브에서 곧게 뻗어야 주배관이 우회하지 않는다(§9.2 C510-3)."""
    heads = [(c * _STEP, r * _STEP) for c in range(6) for r in range(3)]
    straight = PR.zone_axes([_room(heads)], valve_point=(-20 * _STEP, _STEP))
    turned = PR.zone_axes([_room(heads)], valve_point=(2.5 * _STEP, -20 * _STEP))
    assert straight.flipped is False
    assert turned.flipped is True
    assert abs(straight.theta - turned.theta) == pytest.approx(3.14159 / 2, abs=1e-3)


def test_각도_후보는_실_격자에서만_고른다():
    """[문서정합 §9.2] 헤드는 실 격자 위에 있다. 뭉치의 OBB 각도를 새로 재면 어느
    실의 헤드 열과도 어긋난 교차배관이 나온다."""
    axes = PR.zone_axes([_room(_grid(1, 4), axis_deg=30.0)])
    assert axes.to_dict()["cross_axis_deg"] in (30.0, 120.0)


def test_헤드가_없으면_배관을_놓지_않는다():
    with pytest.raises(ValueError):
        PR.zone_axes([_room([])])


# ── R1·R2 — 8개는 한쪽 기준 ─────────────────────────────────────────────

def test_R1_한쪽_8개씩_총_16개는_분할하지_않는다():
    """§9.4 — 전체로 세어 자르면 적법한 가지배관이 둘로 쪼개진다."""
    plan = _plan(_grid(20, 16))
    assert len(plan.crosses) == 1
    assert len(plan.branches) == 20
    assert all((len(b.left), len(b.right)) == (8, 8) for b in plan.branches)
    assert plan.metrics["heads_per_side_max"] == 8
    assert plan.flags == []


def test_R2_한쪽이_9개면_분할한다():
    plan = _plan(_grid(20, 17))
    assert plan.metrics["heads_per_side_max"] <= 8
    assert len(plan.crosses) == 2
    assert plan.flags == []


def test_분기점을_옮겨_해결되면_교차배관을_더_놓지_않는다():
    """분할 전략 1. 한쪽만 넘치고 반대쪽에 여유가 있으면 tee 를 옮긴다."""
    plan = _plan(_grid(20, 12), per_side=8)
    assert len(plan.crosses) == 1
    assert all(max(len(b.left), len(b.right)) <= 8 for b in plan.branches)


def test_분기점_기준_양쪽이_교차배관을_사이에_둔다():
    heads = _grid(20, 16)
    plan = _plan(heads)
    (cross,) = plan.crosses
    ys = _ys(heads)
    for branch in plan.branches:
        assert all(ys[h] < cross.b_mm for h in branch.left)
        assert all(ys[h] > cross.b_mm for h in branch.right)


def test_모든_헤드가_정확히_한_가지배관에_실린다():
    """빠뜨린 헤드는 미방호 구역이 되고, 겹친 헤드는 물을 두 번 센다."""
    heads = [(c * _STEP, r * _STEP) for c in range(4) for r in range(20)]
    plan = _plan(heads)
    carried = [h for b in plan.branches for h in b.heads]
    assert sorted(carried) == sorted(h["id"] for h in _room(heads)["heads"])


def test_가지배관은_교차배관마다_따로_선다():
    """교차배관이 둘이면 같은 헤드 열도 독립 가지배관 둘이다 — 하나로 이으면 루프다."""
    plan = _plan(_grid(20, 20))
    assert len(plan.crosses) == 2
    assert {b.cross_id for b in plan.branches} == {c.id for c in plan.crosses}
    assert len(plan.branches) == 40


def test_상한을_못_맞추면_조용히_넘기지_않는다():
    """한 격자점에 헤드가 몰려 분할로도 안 풀리는 경우. 플래그 없이 넘기면 그
    가지배관은 수리계산만 통과하고 실제로는 물이 안 간다."""
    plan = _plan([(0.0, 0.0)] * 5, per_side=2)
    assert [f["code"] for f in plan.flags] == ["BRANCH_SPLIT_FAILED"]


# ── C530 벽·문·장애물 ──────────────────────────────────────────────────

def test_실_경계에서_문을_빼면_벽이_남는다():
    """[문서정합 §9.2] building.json 에 벽이 없다. 실 경계 중 문이 아닌 변이 벽이다."""
    rooms = _floor(_rect(0, 0, 8000, 6000), _rect(8000, 0, 14000, 6000))
    assert len(PR.wall_segments(rooms, [])) == 7      # 맞닿은 변은 하나로 센다
    door = [{"p1": (8000, 0), "p2": (8000, 6000)}]
    assert len(PR.wall_segments(rooms, door)) == 6


def test_좁고_길쭉한_실만_격자를_낮춘다():
    """폭만 보면 1.5m 창고 하나에 건물 전체 격자가 내려가고 탐색이 네 배가 된다."""
    assert PR._auto_grid_mm(_floor(_rect(0, 0, 20000, 1500))) == PR.GRID_MM_NARROW
    assert PR._auto_grid_mm(_floor(_rect(0, 0, 1500, 1500))) == PR.GRID_MM


# ── C530 경로 ───────────────────────────────────────────────────────────

def test_막힌_것이_없으면_교차배관은_곧다():
    plan, _ = _routed(_plan(_grid(5, 4)),
                      _floor(_rect(-2000, -2000, 14800, 11600)))
    (cross,) = plan.crosses
    assert len(cross.path) == 2
    assert plan.metrics["cross_turns"] == 0
    assert plan.metrics["cross_length_m"] == pytest.approx(12.8)
    assert cross.min_dn == 40                        # 교차배관 법정 최소 40mm


def test_경로는_구역_축에_평행하다():
    """[문서정합 §9.2] 직교 격자를 세계 좌표에 깔면 기운 건물에서 계단이 나온다."""
    heads = _rot(_grid(5, 4), 30.0)
    plan, _ = _routed(_plan(heads, rooms=[_room(heads, axis_deg=30.0)]),
                      _floor(_rot(_rect(-2000, -2000, 14800, 11600), 30.0)))
    (cross,) = plan.crosses
    path = _ab(plan, cross.path)
    assert len(path) >= 2
    for p, q in zip(path, path[1:]):
        assert abs(p[0] - q[0]) < 1e-6 or abs(p[1] - q[1]) < 1e-6


def test_벽은_비용을_물고_지나간다():
    """배관은 슬리브로 벽을 뚫는다. 절대 장애물로 두면 경로가 아예 안 나온다."""
    rooms = _floor(_rect(-2000, -2000, 8000, 11600),
                   _rect(8000, -2000, 14800, 11600))
    plan, _ = _routed(_plan(_grid(5, 4)), rooms)
    assert plan.metrics["wall_pierces"] == 1
    assert plan.metrics["cross_turns"] == 0          # 뚫는 편이 도는 것보다 싸다


def test_문은_벽이_아니다():
    rooms = _floor(_rect(-2000, -2000, 8000, 11600),
                   _rect(8000, -2000, 14800, 11600))
    door = [{"p1": (8000, -2000), "p2": (8000, 11600)}]
    plan, _ = _routed(_plan(_grid(5, 4)), rooms, doors=door)
    assert plan.metrics["wall_pierces"] == 0


def test_코어는_뚫지_않고_돌아간다():
    core = [{"polygon": _rect(7600, 3200, 8800, 6400)}]
    plan, field = _routed(_plan(_grid(5, 4)),
                          _floor(_rect(-2000, -2000, 14800, 11600)), cores=core)
    (cross,) = plan.crosses
    path = _ab(plan, cross.path)
    assert plan.metrics["cross_turns"] >= 2
    assert all(field.clear(p, q) for p, q in zip(path, path[1:]))
    assert plan.flags == []


def test_헤드_면제실도_절대_장애물이다():
    """계단·EV 는 헤드를 안 놓는 실이지 배관이 지나가도 되는 실이 아니다."""
    rooms = _floor(_rect(-2000, -2000, 14800, 11600))
    rooms.append({"room_id": "R-EV", "polygon": _rect(7200, 3200, 9200, 6400),
                  "head_exempt": True})
    plan, _ = _routed(_plan(_grid(5, 4)), rooms)
    assert plan.metrics["cross_turns"] >= 2


# ── C530 재배정 ─────────────────────────────────────────────────────────

_ODD = 8 * _STEP     # 짧은 열의 축 좌표. 8×8 격자(0…7)의 바로 오른쪽이다.


def _short_column(levels):
    """8열 × 8행 격자 오른쪽에 짧은 열 하나. 그 열의 분기점이 재배정 시험대다.

    per_side=2 에서 교차배관은 넷씩 두 구간으로 갈리므로, 짧은 열이 첫 구간에 몇
    개를 갖느냐가 옮길 수 있느냐를 가른다 — 옮기면 전부 한쪽에 실리기 때문이다.
    """
    return ([(c * _STEP, r * _STEP) for c in range(8) for r in range(8)]
            + [(_ODD, r * _STEP) for r in levels])


_BLOCK_TEE = [{"polygon": _rect(_ODD - 1200, 3600, _ODD + 1200, 6000)}]
_WIDE = _floor(_rect(-2000, -2000, _ODD + 2400, 25000))


def test_닿지_못한_분기점은_인접_교차배관으로_옮긴다():
    plan, _ = _routed(_plan(_short_column([0, 1]), per_side=2),
                      _WIDE, cores=_BLOCK_TEE)
    near, far = plan.crosses
    (moved,) = [b for b in plan.branches if b.a_mm > 7 * _STEP]
    assert plan.metrics["branches_reassigned"] == 1
    assert moved.cross_id == far.id != near.id
    assert moved.per_side_max <= 2
    assert len(far.spurs) == 1
    assert plan.flags == []


def test_옮겨서_한쪽_상한이_깨지면_옮기지_않는다():
    """§9.4 를 깨는 재배정은 해결이 아니다. 조용히 옮기면 그 가지배관은 물이 안 간다."""
    plan, _ = _routed(_plan(_short_column([0, 1, 2, 3]), per_side=2),
                      _WIDE, cores=_BLOCK_TEE)
    assert plan.metrics["branches_reassigned"] == 0
    (flag,) = plan.flags
    assert flag["code"] == "ROUTING_UNREACHABLE"
    assert len(flag["heads"]) == 4


def test_닿지_못한_헤드를_조용히_빼지_않는다():
    """빼면 수리계산만 통과하고 미방호 구역이 남는다 — 플래그가 헤드를 들고 있어야 한다."""
    heads = _short_column([0, 1, 2, 3])
    plan, _ = _routed(_plan(heads, per_side=2), _WIDE, cores=_BLOCK_TEE)
    carried = [h for b in plan.branches for h in b.heads]
    assert sorted(carried) == sorted(h["id"] for h in _room(heads)["heads"])
    assert set(plan.flags[0]["heads"]) <= set(carried)


# ── C540 주배관 ─────────────────────────────────────────────────────────

_OPEN = _floor(_rect(-6000, -6000, 14800, 11600))
_VALVE = (-5000.0, -5000.0)


def _mained(plan, rooms, *, valve=_VALVE, cores=(), min_dn=100):
    plan, field = _routed(plan, rooms, cores=cores)
    return PR.route_main(plan, field, valve, min_dn=min_dn), field


def test_주배관은_밸브에서_교차배관_끝으로_간다():
    plan, _ = _mained(_plan(_grid(5, 4)), _OPEN)
    (cross,) = plan.crosses
    assert plan.main.path[0] == pytest.approx(list(_VALVE))
    assert plan.main.path[-1] == pytest.approx(cross.path[0])
    assert plan.main.spurs == []
    assert plan.main.min_dn == 100
    assert plan.metrics["main_length_m"] > 0


def test_주배관은_교차배관_한복판으로_들어가지_않는다():
    """한복판으로 들어가면 그 지점에서 교차배관이 양쪽으로 갈라져 §9.3 형상이 된다."""
    plan, _ = _mained(_plan(_grid(5, 4)), _OPEN, valve=(6400.0, -5000.0))
    (cross,) = plan.crosses
    assert plan.main.path[-1] in (cross.path[0], cross.path[-1])


def test_교차배관이_여럿이면_주배관에서_분기한다():
    """[문서정합 §9.2 C540] 각각을 밸브까지 따로 끌면 주배관이 나란히 겹친다."""
    rooms = _floor(_rect(-6000, -6000, 20 * _STEP + 2000, 17 * _STEP + 2000))
    plan, _ = _mained(_plan(_grid(20, 17)), rooms)
    assert len(plan.crosses) == 2
    (spur,) = plan.main.spurs
    assert list(PR._nearest_on(plan.main.path, spur[0])) == pytest.approx(spur[0])
    ends = [c.path[0] for c in plan.crosses] + [c.path[-1] for c in plan.crosses]
    assert spur[-1] in ends
    assert plan.flags == []


def test_주배관이_못_닿으면_헤드를_들고_보고한다():
    """구역 전체가 미방호다. 주배관 없이 그래프를 완성하면 그 사실이 사라진다."""
    bar = [{"polygon": _rect(-20000, -4000, 30000, -3000)}]
    plan, _ = _mained(_plan(_grid(5, 4)), _OPEN, cores=bar)
    (flag,) = plan.flags
    assert flag["code"] == "MAIN_UNREACHABLE"
    assert len(flag["heads"]) == 20
    assert plan.main.path == []


# ── C540 층 표고 ────────────────────────────────────────────────────────

def _draft(floors):
    return BuildingDraft.from_dict({"source": {"floors": floors}})


def test_층_라벨은_모듈_A_규약으로_읽는다():
    assert [floor_index(s) for s in ("1F", "B1F", "지하2층", "옥탑")] == [1, -1, -2, 99]
    # 지하 표기가 층수와 떨어져 있어도 지하다 — 끝의 "1층"만 보면 지상으로 잡힌다.
    assert floor_index("지하주차장 1층") == -1
    assert floor_index("") is None and floor_index("기계실") is None


def test_지상1층_바닥이_0이고_지하는_아래로_쌓인다():
    elevation, missing = _draft([
        {"label": "B2F", "height_mm": 3000}, {"label": "B1F", "height_mm": 4000},
        {"label": "1F", "height_mm": 4500}, {"label": "2F", "height_mm": 3200},
    ]).floor_elevations()
    assert elevation == {"B2F": -7.0, "B1F": -4.0, "1F": 0.0, "2F": 4.5}
    assert missing == []


def test_옥상은_층수가_아니라_최상층_바로_위다():
    elevation, missing = _draft([
        {"label": "1F", "height_mm": 4500}, {"label": "2F", "height_mm": 3200},
        {"label": "옥상"},
    ]).floor_elevations()
    assert elevation["옥상"] == pytest.approx(7.7)
    assert missing == []


def test_층고를_모르면_그_층부터_빼고_돌려준다():
    """지어내 메우면 그 값이 입상관 길이가 되고 그대로 낙차 압력이 된다(G9)."""
    elevation, missing = _draft([
        {"label": "1F", "height_mm": 4500}, {"label": "2F"},
        {"label": "3F", "height_mm": 3200}, {"label": "4F", "height_mm": 3200},
    ]).floor_elevations()
    assert elevation == {"1F": 0.0, "2F": 4.5}
    assert missing == ["3F", "4F"]


# ── C540 입상관 ─────────────────────────────────────────────────────────

def _riser(levels, *, polygon=None):
    return PR.plan_riser({"core_id": "CR-1", "polygon": polygon
                          or _rect(0, 0, 4000, 4000)}, levels)


def test_입상관은_코어_안에_선다():
    """ㄱ자 코어에서 무게중심은 밖으로 나간다. 입상관은 샤프트 안에 서야 한다."""
    ell = [(0, 0), (10000, 0), (10000, 2000), (2000, 2000), (2000, 10000), (0, 10000)]
    riser = _riser([], polygon=ell)
    assert PR.point_in_polygon(riser.point, ell)


def test_배관장은_도면상_길이와_표고차_중_큰_쪽이다():
    """표시용으로 눌러 그린 층에서 도면상 길이만 쓰면 낙차가 사라진다(§9.2 C540)."""
    riser = _riser([PR.RiserLevel("1F", 0.0, 0.0),
                    PR.RiserLevel("B1F", -4.0, -1.0),
                    PR.RiserLevel("2F", 4.5, 20.0)])
    lengths = {seg["id"]: seg["length_m"] for seg in PR.riser_segments(riser)}
    assert lengths == {"RS-CR-1-B1F-1F": 4.0, "RS-CR-1-1F-2F": 20.0}


def test_최하층_종점이_최종_급수원이다():
    riser = _riser([PR.RiserLevel("1F", 0.0, 0.0), PR.RiserLevel("B1F", -4.0, -4.0)])
    assert riser.source_node == "RS-CR-1-B1F"
    assert PR.riser_segments(riser)[0]["from"] == riser.source_node


def test_표고가_없는_입상관은_급수원도_없다():
    assert _riser([]).source_node is None and PR.riser_segments(_riser([])) == []


# ── R3·R4 — 토너먼트 금지 ───────────────────────────────────────────────

def _undirected(edges):
    graph: dict = {}
    for a, b in edges:
        graph.setdefault(a, set()).add(b)
        graph.setdefault(b, set()).add(a)
    return graph


def test_R3_대칭_이분_분기는_토너먼트다():
    """총연장 최소화(MST/Steiner)가 수렴하는 형상. 최적화가 위법을 만든다(§9.3)."""
    graph = _undirected([("S", "A"), ("A", "B"), ("A", "C"),
                         ("B", "B1"), ("B", "B2"), ("C", "C1"), ("C", "C2")])
    assert PR.check_tournament(graph, "S")


def test_R4_빗살은_통과한다():
    """교차배관 1개 + 거기서 갈라지는 가지배관들. 허용 형상은 이것뿐이다."""
    edges = [("S", "M1")]
    for n in range(1, 5):
        edges += [(f"M{n}", f"M{n + 1}"), (f"M{n}", f"B{n}"), (f"B{n}", f"H{n}")]
    assert PR.check_tournament(_undirected(edges), "S") == []


def test_교차배관_위의_이웃한_분기점은_토너먼트가_아니다():
    """[문서정합 §9.3] 명세 예시("연속 2회 분기")를 그대로 쓰면 빗살이 전부
    걸린다 — 교차배관의 분기점들은 서로 이웃한 차수 3 노드다."""
    edges = [("S", "T1"), ("T1", "T2"), ("T1", "b1"), ("T2", "b2")]
    assert PR.check_tournament(_undirected(edges), "S") == []


def test_급수원에서_교차배관_둘로_갈라지는_것은_주배관_배열이다():
    edges = [("S", "L1"), ("S", "R1")]
    for side in ("L", "R"):
        edges += [(f"{side}1", f"{side}2"), (f"{side}1", f"{side}b1"),
                  (f"{side}2", f"{side}b2")]
    assert PR.check_tournament(_undirected(edges), "S") == []


def _hierarchy():
    """급수원 → 주배관 → 교차배관 둘 → 분기점 → 가지배관. 실제 설계의 위계다."""
    edges = [("PUMP", "AV"), ("AV", "M0"), ("M0", "M1"), ("M1", "M2")]
    tees = []
    for side, hub in (("L", "M1"), ("R", "M2")):
        edges.append((hub, f"{side}C1"))
        for n in (1, 2):
            tee = f"{side}T{n}"
            edges += [(f"{side}C{n}", tee), (f"{side}C{n}", f"{side}C{n + 1}"),
                      (tee, f"{side}H{n}")]
            tees.append(tee)
    return _undirected(edges), tees


def test_주배관_위계는_가지배관_배열이_아니다():
    """[문서정합 §9.3] 주배관이 들어오면 급수원 하나를 빼는 것으로 모자란다 —
    주배관이 교차배관 둘로 갈라지고 그 둘이 각각 분기점을 품는 순간 적법한 위계가
    통째로 걸린다. NFTC 2.5.10.1 이 금하는 것은 가지배관 배열이다."""
    graph, tees = _hierarchy()
    assert PR.check_tournament(graph, "PUMP")                     # 전체로 보면 걸린다
    assert PR.check_tournament(graph, "PUMP", roots=tees) == []


def test_분기점_하류가_또_갈라지면_범위를_줘도_걸린다():
    """범위를 좁힌 것이지 검사를 끈 것이 아니다."""
    graph = _undirected([("S", "T"), ("T", "a"), ("a", "b"), ("a", "c"),
                         ("b", "b1"), ("b", "b2"), ("c", "c1"), ("c", "c2")])
    assert PR.check_tournament(graph, "S", roots=["T"])


def test_없는_급수원은_검사하지_않는다():
    assert PR.check_tournament(_undirected([("A", "B")]), "S") == []


def test_우리가_만든_빗살은_토너먼트가_아니다():
    """C520 산출물을 그래프로 세워 스스로 검사한다. 검사가 우리 산출물을 통과하지
    못하면 둘 중 하나가 틀린 것이다."""
    plan = _plan(_grid(5, 6))
    edges = []
    for cross in plan.crosses:
        mine = sorted((b for b in plan.branches if b.cross_id == cross.id),
                      key=lambda b: b.a_mm)
        edges.append(("SRC", f"{cross.id}-0"))
        for n, branch in enumerate(mine, start=1):
            tee = f"{cross.id}-{n}"
            edges.append((f"{cross.id}-{n - 1}", tee))
            for side in (branch.left, branch.right):
                prev = tee
                for head in side:
                    edges.append((prev, head))
                    prev = head
    assert PR.check_tournament(_undirected(edges), "SRC") == []
