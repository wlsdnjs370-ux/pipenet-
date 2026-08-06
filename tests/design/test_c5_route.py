# -*- coding: utf-8 -*-
"""지시서 §9·§13.5 — C5 배관망 라우트.

지키는 것은 셋이다. **헤드 없이는 관을 놓지 않는다**(C4 가 아직이면 어디로 물을
보낼지 모른다), **근거 없는 값은 채우지 않는다**(층고를 못 구하면 그 층은 빠지고
플래그로 말한다), **낸 표는 모듈 A 가 삼킬 수 있어야 한다**(nodes/pipes/nozzles
여섯 표가 그대로 SDF 가 된다).

도면은 test_c4_route 와 같다 — 문으로 이어진 A·B·C 와 떨어진 별채 D.

`_inner_dia_mm` 이 평면 import 인 `hydraulic_solver` 를 열므로 `core` 도
sys.path 에 올린다. 서버(`대조 서버.py`)가 이미 지는 부담과 같다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_ROOT / "core"))

import routes.r30_design as r30_design  # noqa: E402
from core.design import session as S  # noqa: E402

_REQ_OK = {"mount_height_ok": True, "door_size_ok": True,
           "room_temp_ok": True, "signage_ok": True}
_SPANS = {"A": (0, 3000), "B": (3000, 8000), "C": (8000, 13000), "D": (20000, 23000)}
_AREAS = {"A": 9.0, "B": 15.0, "C": 15.0, "D": 9.0}


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
    return {"p1": [x, 0], "p2": [x, 3000], "kind": "door",
            "confidence": 0.9, "is_virtual": True}


def _building(*, rooms="ABC"):
    return {
        "schema": "fncadnet.building/1",
        "source": {"dxf": "1F.dxf",
                   "floors": [{"label": "1F", "dxf": "1F.dxf", "height_mm": 3200}]},
        "building": {"floors_total": 12, "structure": "내화구조", "use": "판매시설"},
        "rooms": [_room(f"R-1F-{rid}", *_SPANS[rid], _AREAS[rid]) for rid in rooms],
        "virtual_edges": [_door(3000), _door(8000)],
        "cores": [{"id": "SH-01", "kind": "shaft", "polygon": _rect(*_SPANS["A"]),
                   "area_m2": 9.0, "confirmed": True,
                   "provenance": {"confirmed": "GATE"}}],
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
        assert client.post("/api/design/c2/constraints",
                           json={"session_id": sid}).status_code == 200
    return sid


def _placed(client, **kw) -> str:
    """밸브·구역·헤드까지 끝난 세션. 여기가 C5 의 출발점이다."""
    sid = _seed(client, **kw)
    assert client.post("/api/design/c3/valves", json={
        "session_id": sid, "operator": "jinwon",
        "valves": [{"core_id": "SH-01", "point": [1500, 1500],
                    "system_type": "습식",
                    "requirements_confirmed": dict(_REQ_OK)}],
    }).status_code == 200
    assert client.post("/api/design/c3/zones",
                       json={"session_id": sid}).status_code == 200
    assert client.post("/api/design/c4/heads",
                       json={"session_id": sid,
                             "operator": "jinwon"}).status_code == 200
    return sid


def _route(client, sid, **body):
    return client.post("/api/design/c5/route",
                       json={"session_id": sid, "operator": "jinwon", **body})


# ── 전제 ────────────────────────────────────────────────────────────────

def test_기준을_굽기_전에는_라우팅하지_않는다(client):
    sid = _seed(client, constraints=False)
    assert _route(client, sid).get_json()["code"] == "CONSTRAINTS_REQUIRED"


def test_구역이_없으면_라우팅하지_않는다(client):
    sid = _seed(client)
    res = _route(client, sid)
    assert res.status_code == 409
    assert res.get_json()["code"] == "ZONE_REQUIRED"


def test_헤드가_없으면_라우팅하지_않는다(client):
    """놓을 헤드를 모르면 가지배관의 길이도 관경도 정할 수 없다."""
    sid = _seed(client)
    assert client.post("/api/design/c3/valves", json={
        "session_id": sid, "operator": "jinwon",
        "valves": [{"core_id": "SH-01", "point": [1500, 1500],
                    "system_type": "습식",
                    "requirements_confirmed": dict(_REQ_OK)}],
    }).status_code == 200
    assert client.post("/api/design/c3/zones",
                       json={"session_id": sid}).status_code == 200
    assert _route(client, sid).get_json()["code"] == "HEADS_REQUIRED"


# ── 배관망 ──────────────────────────────────────────────────────────────

def test_코어마다_망_하나가_나온다(client):
    sid = _placed(client)
    body = _route(client, sid).get_json()
    assert body["ok"] is True
    (net,) = body["networks"]
    assert net["riser"]["core_id"] == "SH-01"
    assert net["metrics"]["heads"] > 0
    assert net["metrics"]["length_m"] > 0


def test_모듈A_여섯_표가_그대로_실린다(client):
    """이 표들이 PR-10 에서 SDF·KFP·HAS 가 된다. 이름이 어긋나면 거기서 조용히 빈다."""
    sid = _placed(client)
    (net,) = _route(client, sid).get_json()["networks"]
    graph = net["graph"]
    assert set(graph) >= {"nodes", "pipes", "nozzles", "fittings",
                          "equipment", "meta"}
    assert len(graph["nozzles"]) == net["metrics"]["heads"]
    assert all({"label", "x", "y", "elevation"} <= set(n) for n in graph["nodes"])
    assert all({"in", "out", "length", "dia"} <= set(p) for p in graph["pipes"])
    assert sum(1 for n in graph["nodes"] if n["io_node"] == "Input") == 1


def test_모든_헤드가_급수원에서_닿는다(client):
    """닿지 않은 헤드는 그만큼 미방호다 — 조용히 넘기지 않고 플래그로 세운다."""
    sid = _placed(client)
    body = _route(client, sid).get_json()
    assert "HEAD_UNREACHED" not in {f["code"] for f in body["flags"]}


def test_관경은_급수원_쪽이_더_굵거나_같다(client):
    """단조성이 깨지면 가는 관에서 굵은 관으로 흐르는 그림이 된다."""
    sid = _placed(client)
    (net,) = _route(client, sid).get_json()["networks"]
    pipes = net["graph"]["pipes"]
    assert all(int(p["dia"]) > 0 for p in pipes)

    downstream: dict[str, list[dict]] = {}
    for pipe in pipes:
        downstream.setdefault(pipe["in"], []).append(pipe)
    for pipe in pipes:
        for child in downstream.get(pipe["out"], ()):
            assert int(child["dia"]) <= int(pipe["dia"])


def test_분기점이_하나뿐인_구역도_급수원에_이어진다(client):
    """실 하나에 헤드 한 줄이면 교차배관은 이을 두 분기점이 없어 점 하나로 줄어든다.
    그 점을 '경로 없음' 으로 읽고 건너뛰면 주배관이 아예 놓이지 않아 구역 전체가
    급수원과 끊긴다 — 표는 차 있는데 물이 안 가는 그림이다."""
    sid = _placed(client, rooms="A")
    (net,) = _route(client, sid).get_json()["networks"]
    roles = {p["role"] for p in net["graph"]["pipes"]}
    assert {"branch", "main"} <= roles
    assert net["metrics"]["length_m"] > 0


def test_결과가_세션에_남는다(client):
    sid = _placed(client)
    _route(client, sid)
    sess = _sess(client, sid)
    design, _v = sess.read("design.json")
    assert len(design["networks"]) == 1

    graph, _v = sess.read("graph.json")
    assert graph["schema"] == "fncadnet.graph/1"
    assert graph["networks"][0]["graph"]["pipes"]


def test_막는_플래그가_없으면_방출_단계가_열린다(client):
    sid = _placed(client)
    body = _route(client, sid).get_json()
    assert body["blocking"] == []
    assert _sess(client, sid).status()["meta"]["stage"] == "emit"
    # 단계를 화면이 따로 셈하면 서버와 갈린다 — 정한 값을 그대로 실어 보낸다.
    assert body["stage"] == "emit"


def test_앞_단계가_막고_있으면_배관망이_깨끗해도_방출은_안_열린다(client):
    """C3 가 문제를 안고 있으면 stage 는 거기 서 있다. 배관망만 보고 방출을 열면
    닿지 않는 실을 안은 채 솔버 파일이 나간다."""
    sid = _placed(client)
    _sess(client, sid).update_meta(stage="c3")
    body = _route(client, sid).get_json()
    assert body["blocking"] == [] and body["stage"] == "c3"
    assert _sess(client, sid).status()["meta"]["stage"] == "c3"


def test_라우팅_기록이_감사로그에_남는다(client):
    sid = _placed(client)
    _route(client, sid)
    (entry,) = [e for e in _sess(client, sid).audit_entries()
                if e["event"] == "network_routed"]
    assert entry["stage"] == "C5"
    assert entry["detail"]["networks"] == 1
    assert entry["detail"]["metrics"][0]["pipes"] > 0
