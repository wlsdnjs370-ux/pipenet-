# -*- coding: utf-8 -*-
"""지시서 §7 — C3 유수검지장치와 방호구역.

지키는 것은 세 가지다. **밸브는 사람이 찍는다**(후보는 확정된 코어뿐이고 설비
종류·설치 요건은 기본값을 넣지 않는다), **거리는 문을 지나는 경로거리다**(직선으로
재면 벽 하나 사이의 실이 가까워 보인다), **닿지 않는 실은 억지로 붙이지 않는다**
(붙이면 어느 유수검지장치에도 물리지 않은 헤드가 C4·C5 를 통과한다).

도면은 문으로 이어진 세 실(코어 A — B — C)과 문이 없는 별채 D 다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

import routes.r30_design as r30_design  # noqa: E402
from core.design import session as S  # noqa: E402

_REQ_OK = {"mount_height_ok": True, "door_size_ok": True,
           "room_temp_ok": True, "signage_ok": True}

# A(코어) 0~3m — B 3~8m — C 8~13m 이 x 축으로 붙어 있고, D 는 20m 밖에 떨어져 있다.
_SPANS = {"A": (0, 3000), "B": (3000, 8000), "C": (8000, 13000), "D": (20000, 23000)}


def _rect(x0, x1):
    return [[x0, 0], [x1, 0], [x1, 3000], [x0, 3000]]


def _room(rid, x0, x1, area):
    return {
        "id": rid, "floor": "1F", "polygon": _rect(x0, x1), "area_m2": area,
        "use": "판매시설", "ceiling": {"has_finish": False, "slab_height_mm": 3200},
        "ambient_temp_max_c": 29.0,
        "provenance": {"use": "GATE", "ceiling.has_finish": "GATE",
                       "ceiling.slab_height_mm": "default",
                       "ambient_temp_max_c": "default"},
    }


def _door(x):
    """A|B, B|C 경계의 공유 변. C160 이 내는 가상 간선과 같은 모양이다."""
    return {"p1": [x, 0], "p2": [x, 3000], "kind": "door",
            "confidence": 0.9, "is_virtual": True}


def _building(*, areas=None, doors=("AB", "BC"), cores=None):
    area = {"A": 9.0, "B": 15.0, "C": 15.0, "D": 9.0, **(areas or {})}
    edges = {"AB": _door(3000), "BC": _door(8000)}
    return {
        "schema": "fncadnet.building/1",
        "source": {"dxf": "1F.dxf", "floors": [{"label": "1F", "dxf": "1F.dxf"}]},
        "building": {"floors_total": 12, "structure": "내화구조", "use": "판매시설"},
        "rooms": [_room(f"R-1F-{rid}", *_SPANS[rid], area[rid]) for rid in "ABCD"],
        "virtual_edges": [edges[k] for k in doors],
        "cores": cores if cores is not None else [
            {"id": "SH-01", "kind": "shaft", "polygon": _rect(*_SPANS["A"]),
             "area_m2": 9.0, "confirmed": True, "provenance": {"confirmed": "GATE"}},
        ],
        "obstacles": {"status": "partial", "provenance": {"status": "GATE"}},
        "gate": {"passed": True, "operator": "jinwon", "unresolved": []},
    }


@pytest.fixture()
def client(tmp_path):
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=True)
    app.config["DESIGN_SESSION_DIR"] = tmp_path
    return app.test_client()


def _sess(client, sid):
    return S.DesignSession.open(client.application.config["DESIGN_SESSION_DIR"], sid)


def _seed(client, *, constraints=True, **kw) -> str:
    sid = client.post("/api/design/session",
                      json={"operator": "jinwon"}).get_json()["session_id"]
    _sess(client, sid).write("building.json", _building(**kw))
    if constraints:
        res = client.post("/api/design/c2/constraints", json={"session_id": sid})
        assert res.status_code == 200, res.get_json()
    return sid


def _place(client, sid, *, point=(1500, 1500), core_id="SH-01",
           system_type="습식", **kw):
    valve = {"core_id": core_id, "point": list(point), "system_type": system_type,
             "requirements_confirmed": dict(_REQ_OK)}
    valve.update(kw)
    return client.post("/api/design/c3/valves",
                       json={"session_id": sid, "operator": "jinwon",
                             "valves": [valve]})


# ── C3.1 밸브 후보와 확정 (§7.1) ────────────────────────────────────────

def test_후보는_확정된_샤프트_계단_코어뿐이다(client):
    """`confirmed` 가 `None` 인 코어를 후보로 내면 GATE 가 왜 있는지가 사라진다."""
    sid = _seed(client, cores=[
        {"id": "SH-01", "kind": "shaft", "polygon": _rect(*_SPANS["A"]),
         "area_m2": 9.0, "confirmed": True},
        {"id": "SH-02", "kind": "shaft", "polygon": _rect(*_SPANS["B"]),
         "area_m2": 15.0, "confirmed": None},
        {"id": "EV-01", "kind": "elevator", "polygon": _rect(*_SPANS["C"]),
         "area_m2": 15.0, "confirmed": True},
    ])
    body = client.get(f"/api/design/c3/candidates/{sid}").get_json()
    assert [c["core_id"] for c in body["candidates"]] == ["SH-01"]
    assert "부압식" in body["system_types"]
    assert [r["key"] for r in body["requirements"]] == list(
        r30_design.Z.MANUAL_REQUIREMENTS)


def test_제안하는_밸브_자리는_실_안에_있다():
    """오목한 실은 무게중심이 밖으로 나간다(양주옥정 코어 11개 중 1개).

    밖으로 나간 점을 화면이 제안하면, 사람이 제안대로 찍었을 때 '실 밖' 이라며
    거절당한다.
    """
    from core.design.recognize.spatial import (
        centroid, point_in_polygon, representative_point)

    u_shape = [[0, 0], [3000, 0], [3000, 10000], [2500, 10000],
               [2500, 500], [500, 500], [500, 10000], [0, 10000]]
    assert not point_in_polygon(centroid(u_shape), u_shape)   # 무게중심은 밖이다
    assert point_in_polygon(representative_point(u_shape), u_shape)


def test_후보가_아닌_코어에는_밸브를_놓을_수_없다(client):
    sid = _seed(client, cores=[
        {"id": "SH-02", "kind": "shaft", "polygon": _rect(*_SPANS["A"]),
         "area_m2": 9.0, "confirmed": None},
    ])
    res = _place(client, sid, core_id="SH-02")
    assert res.status_code == 422
    assert res.get_json()["code"] == "VALVE_CORE_UNCONFIRMED"


def test_설비_종류를_고르지_않으면_습식으로_확정하지_않는다(client):
    """동결·수손 우려는 도면에 없다. 기본값이 곧 잘못된 설비 종류의 확정이다."""
    sid = _seed(client)
    res = _place(client, sid, system_type="")
    assert res.status_code == 422
    assert res.get_json()["code"] == "SYSTEM_TYPE_REQUIRED"


def test_묻지_않은_설치_요건을_거짓으로_접지_않는다(client):
    sid = _seed(client)
    res = _place(client, sid, requirements_confirmed={"mount_height_ok": True})
    assert res.status_code == 422
    body = res.get_json()
    assert body["code"] == "VALVE_REQUIREMENTS_REQUIRED"
    assert set(body["fields"]) == {"door_size_ok", "room_temp_ok", "signage_ok"}


def test_요건에_아니라고_답하면_확정은_되되_미충족으로_남는다(client):
    """'아니오'는 사람의 답이다. 막지는 않고 그대로 기록해 화면이 보여준다."""
    sid = _seed(client)
    res = _place(client, sid,
                 requirements_confirmed={**_REQ_OK, "signage_ok": False})
    assert res.status_code == 200
    body = res.get_json()
    assert body["requirements_unmet"] == [{"valve_id": "AV-1F-01",
                                           "fields": ["signage_ok"]}]


def test_실_밖을_찍으면_가까운_실로_끌어붙이지_않는다(client):
    sid = _seed(client)
    res = _place(client, sid, point=(50000, 50000))
    assert res.status_code == 422
    assert res.get_json()["code"] == "VALVE_OUTSIDE_ROOMS"


def test_기준을_굽기_전에는_구역을_나누지_않는다(client):
    sid = _seed(client, constraints=False)
    res = _place(client, sid)
    assert res.status_code == 409
    assert res.get_json()["code"] == "CONSTRAINTS_REQUIRED"


def test_밸브를_다시_찍으면_옛_구역은_지워진다(client):
    """밸브가 바뀌면 구역은 근거를 잃는다. 남기면 옛 구역을 새 밸브의 것으로 읽는다."""
    sid = _seed(client)
    _place(client, sid)
    client.post("/api/design/c3/zones", json={"session_id": sid})
    assert _sess(client, sid).read("design.json")[0]["zones"]

    _place(client, sid, point=(2000, 2000))
    assert "zones" not in _sess(client, sid).read("design.json")[0]


# ── C3.2 방호구역 (§7.2) ────────────────────────────────────────────────

def test_구역은_문으로_닿는_실만_모은다(client):
    sid = _seed(client)
    _place(client, sid)
    res = client.post("/api/design/c3/zones", json={"session_id": sid})
    assert res.status_code == 200
    body = res.get_json()
    zone = body["zones"][0]
    assert zone["rooms"] == ["R-1F-A", "R-1F-B", "R-1F-C"]
    assert body["unreached"] == ["R-1F-D"]
    assert [f["code"] for f in body["flags"]] == ["ROOM_UNREACHABLE"]
    # 경로거리: A중심(1.5) → 문(3.0) → B중심(5.5) → 문(8.0) → C중심(10.5)
    assert body["room_distance_m"]["R-1F-C"] == pytest.approx(9.0)


def test_문이_없으면_닿았다고_말하지_않는다(client):
    """벽 하나를 사이에 둔 실은 직선으로는 가깝지만 배관은 지나갈 수 없다."""
    sid = _seed(client, doors=("AB",))
    _place(client, sid)
    body = client.post("/api/design/c3/zones",
                       json={"session_id": sid}).get_json()
    assert body["zones"][0]["rooms"] == ["R-1F-A", "R-1F-B"]
    assert body["unreached"] == ["R-1F-C", "R-1F-D"]


def test_같은_변을_왕복하는_실은_자기_자신과_이어지지_않는다():
    """실도면에는 경계가 같은 변을 두 번 지나는 슬리버 face 가 있다(양주옥정 2건).

    변마다 실 이름을 세기만 하면 그런 실은 '두 실이 공유하는 변' 으로 보여 자기
    자신으로 가는 문이 생긴다.
    """
    from core.design.deterministic import zoning as Z
    from core.design.schema import BuildingDraft

    draft = BuildingDraft.from_dict({
        "rooms": [_room("R-1F-S", 0, 3000, 9.0)],
        # 사각형에서 나갔다가 같은 선을 되짚어 돌아오는 돌기.
        "virtual_edges": [{"p1": [-2000, 1500], "p2": [0, 1500], "kind": "door"}],
    })
    draft.rooms[0].polygon = [[0, 0], [3000, 0], [3000, 3000], [0, 3000],
                              [0, 1500], [-2000, 1500], [0, 1500]]
    graph = Z.build_room_graph(draft)
    assert graph.doors == {}
    assert graph.orphans == ["R-1F-S"]


def test_면적_초과는_자동으로_쪼개지_않고_사람에게_돌린다(client):
    """[문서정합 §7.2] 자동 분할은 유수검지장치 없는 방호구역을 만든다."""
    sid = _seed(client, areas={"B": 3200.0})
    _place(client, sid)
    body = client.post("/api/design/c3/zones",
                       json={"session_id": sid}).get_json()
    assert len(body["zones"]) == 1          # 쪼개지 않았다
    assert body["zones"][0]["reachability"] == "area_exceeded"
    assert "ZONE_AREA_EXCEEDED" in [f["code"] for f in body["flags"]]


def test_밸브_없이는_구역을_나누지_않는다(client):
    sid = _seed(client)
    res = client.post("/api/design/c3/zones", json={"session_id": sid})
    assert res.status_code == 409
    assert res.get_json()["code"] == "VALVE_REQUIRED"


def test_남은_문제가_없을_때만_다음_단계가_열린다(client):
    """닿지 않는 실을 안은 채로 C4 헤드 배치가 열리면 그 헤드는 근거가 없다."""
    sid = _seed(client)
    _place(client, sid)
    client.post("/api/design/c3/zones", json={"session_id": sid})
    assert _sess(client, sid).status()["meta"]["stage"] == "c3"   # D 가 남았다

    raw, version = _sess(client, sid).read("building.json")
    raw["rooms"] = [r for r in raw["rooms"] if r["id"] != "R-1F-D"]
    _sess(client, sid).write("building.json", raw, if_version=version)

    body = client.post("/api/design/c3/zones",
                       json={"session_id": sid}).get_json()
    assert body["flags"] == []
    assert _sess(client, sid).status()["meta"]["stage"] == "c4"


def test_구역_배정은_다시_돌려도_같다(client):
    """거리가 같은 실이 실행마다 다른 밸브에 붙으면 감사에서 재현되지 않는다."""
    sid = _seed(client)
    client.post("/api/design/c3/valves", json={
        "session_id": sid, "operator": "jinwon", "valves": [
            {"core_id": "SH-01", "point": [1500, 1500], "system_type": "습식",
             "requirements_confirmed": dict(_REQ_OK)},
            {"core_id": "SH-01", "point": [1000, 1000], "system_type": "건식",
             "requirements_confirmed": dict(_REQ_OK)},
        ]})
    first = client.post("/api/design/c3/zones", json={"session_id": sid}).get_json()
    again = client.post("/api/design/c3/zones", json={"session_id": sid}).get_json()
    assert first["zones"] == again["zones"]
    # 두 밸브가 같은 실에 있어도 한쪽만 실을 가져간다 — 나머지는 빈 구역이다.
    assert sorted(len(z["rooms"]) for z in first["zones"]) == [0, 3]
    assert "ZONE_EMPTY" in [f["code"] for f in first["flags"]]


# ── C3.4 헤드 사양 (§7.4) ───────────────────────────────────────────────

def test_헤드_사양은_기준을_읽기만_한다(client):
    """여기서 R·표시온도를 다시 정하면 한 건물에 두 벌의 기준이 생긴다."""
    from core.design.deterministic import zoning as Z
    from core.design.schema import Room

    sid = _seed(client)
    constraints, _v = _sess(client, sid).read("constraints.json")
    spec = Z.head_spec(Room.from_dict(_room("R-1F-A", 0, 3000, 9.0)), constraints)
    assert spec["orientation"] == "upright"       # 반자가 없다
    assert spec["flex"] is None                   # 신축배관은 반자 속에만
    assert spec["temp_rating_c"] == constraints["temp_rating_c"]
    assert spec["k_factor"] == constraints["k_factor"]


def test_반자_유무를_모르면_헤드_사양을_정하지_않는다(client):
    from core.design.deterministic import zoning as Z
    from core.design.schema import Room

    sid = _seed(client)
    constraints, _v = _sess(client, sid).read("constraints.json")
    room = _room("R-1F-A", 0, 3000, 9.0)
    room["ceiling"] = {"slab_height_mm": 3200}
    with pytest.raises(Z.ZoningError) as exc:
        Z.head_spec(Room.from_dict(room), constraints)
    assert exc.value.code == "MISSING_BUILDING_FACT"
