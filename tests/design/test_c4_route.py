# -*- coding: utf-8 -*-
"""지시서 §8·§13.5 — C4 헤드 배치 라우트.

지키는 것은 세 가지다. **구역 없이는 헤드를 놓지 않는다**(어느 밸브도 담당하지
않는 실의 헤드에는 C5 가 물을 보낼 수 없다), **경고와 실패는 다르다**(장애물
미확인은 다음 단계를 막지 않고, 피복이 뚫린 것은 막는다), **검사는 배치와 함께
남는다**(나중에 따로 돌리는 검사는 아무도 안 돌린다).

도면은 test_c3_route 와 같다 — 문으로 이어진 A·B·C 와 떨어진 별채 D.
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
    return {"p1": [x, 0], "p2": [x, 3000], "kind": "door",
            "confidence": 0.9, "is_virtual": True}


_AREAS = {"A": 9.0, "B": 15.0, "C": 15.0, "D": 9.0}


def _building(*, obstacles=None, rooms="ABCD"):
    return {
        "schema": "fncadnet.building/1",
        "source": {"dxf": "1F.dxf", "floors": [{"label": "1F", "dxf": "1F.dxf"}]},
        "building": {"floors_total": 12, "structure": "내화구조", "use": "판매시설"},
        "rooms": [_room(f"R-1F-{rid}", *_SPANS[rid], _AREAS[rid]) for rid in rooms],
        "virtual_edges": [_door(3000), _door(8000)],
        "cores": [{"id": "SH-01", "kind": "shaft", "polygon": _rect(*_SPANS["A"]),
                   "area_m2": 9.0, "confirmed": True,
                   "provenance": {"confirmed": "GATE"}}],
        "obstacles": obstacles or {"status": "partial",
                                   "provenance": {"status": "GATE"}},
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


def _zoned(client, **kw) -> str:
    """밸브를 찍고 구역까지 나눈 세션. 여기가 C4 의 출발점이다."""
    sid = _seed(client, **kw)
    assert client.post("/api/design/c3/valves", json={
        "session_id": sid, "operator": "jinwon",
        "valves": [{"core_id": "SH-01", "point": [1500, 1500],
                    "system_type": "습식",
                    "requirements_confirmed": dict(_REQ_OK)}],
    }).status_code == 200
    assert client.post("/api/design/c3/zones",
                       json={"session_id": sid}).status_code == 200
    return sid


def _heads(client, sid, **body):
    return client.post("/api/design/c4/heads",
                       json={"session_id": sid, "operator": "jinwon", **body})


# ── 전제 ────────────────────────────────────────────────────────────────

def test_기준을_굽기_전에는_배치하지_않는다(client):
    sid = _seed(client, constraints=False)
    res = _heads(client, sid)
    assert res.status_code == 409
    assert res.get_json()["code"] == "CONSTRAINTS_REQUIRED"


def test_구역이_없으면_배치하지_않는다(client):
    """구역 없이 놓은 헤드는 어느 유수검지장치의 것인지 말할 수 없다."""
    sid = _seed(client)
    res = _heads(client, sid)
    assert res.status_code == 409
    assert res.get_json()["code"] == "ZONE_REQUIRED"


def test_배치_전에는_검사표가_없다(client):
    sid = _seed(client)
    res = client.get(f"/api/design/checks/{sid}")
    assert res.status_code == 409
    assert res.get_json()["code"] == "CHECKS_NOT_RUN"


# ── 배치 ────────────────────────────────────────────────────────────────

def test_구역에_든_실만_배치하고_빠진_실은_세어_보인다(client):
    """별채 D 는 어느 밸브에서도 문으로 닿지 않는다. 지우지도, 억지로 붙이지도 않는다."""
    sid = _zoned(client)
    body = _heads(client, sid).get_json()
    assert body["ok"] is True
    assert body["rooms_without_zone"] == ["R-1F-D"]
    assert [r["room_id"] for r in body["rooms"]] == ["R-1F-A", "R-1F-B", "R-1F-C"]
    assert all(r["heads"] for r in body["rooms"])


def test_헤드_아이디는_실을_넘어서도_유일하다(client):
    """C5 가 이 아이디로 가지배관을 묶는다. 겹치면 다른 실의 헤드가 한 가지에 붙는다."""
    sid = _zoned(client)
    ids = [h["id"] for r in _heads(client, sid).get_json()["rooms"] for h in r["heads"]]
    assert len(set(ids)) == len(ids)


def test_배치가_끝나면_다음_단계가_열린다(client):
    sid = _zoned(client, rooms="ABC")
    assert _heads(client, sid).status_code == 200
    assert _sess(client, sid).status()["meta"]["stage"] == "c5"


def test_c3가_문제를_안고_있으면_단계를_넘기지_않는다(client):
    """별채 D 가 남아 C3 는 아직 c3 에 서 있다. 여기서 c5 로 밀면 닿지 않는 실을
    안은 채 C5 가 열리고, 건너뛴 C4 는 아무도 다시 보지 않는다."""
    sid = _zoned(client)
    assert _heads(client, sid).get_json()["rooms_without_zone"] == ["R-1F-D"]
    assert _sess(client, sid).status()["meta"]["stage"] == "c3"


def test_장애물_미확인은_경고이지_실패가_아니다(client):
    """§8.5 — 플래그는 남기되 다음 단계를 막지는 않는다."""
    sid = _zoned(client, rooms="ABC",
                 obstacles={"status": "none", "provenance": {"status": "GATE"}})
    body = _heads(client, sid).get_json()
    assert {f["code"] for f in body["flags"]} == {"OBSTACLE_UNVERIFIED"}
    assert body["blocking"] == []
    assert _sess(client, sid).status()["meta"]["stage"] == "c5"


def test_결과와_검사표가_세션에_남는다(client):
    sid = _zoned(client)
    _heads(client, sid)
    design, _v = _sess(client, sid).read("design.json")
    assert [r["room_id"] for r in design["heads"]] == ["R-1F-A", "R-1F-B", "R-1F-C"]

    body = client.get(f"/api/design/checks/{sid}").get_json()
    assert [c["code"] for c in body["checks"]] == [
        "TYPICAL_FLOOR_MISMATCH", "ROOM_AREA_MISMATCH", "HEAD_COUNT_DEVIATION",
        "EXEMPT_AREA_EXCESS", "PIPE_LENGTH_PER_HEAD"]


def test_연면적을_주면_면적_대조가_미검증에서_풀린다(client):
    sid = _zoned(client)
    body = _heads(client, sid, gross_floor_area_m2=48.0).get_json()
    (rec,) = [c for c in body["checks"]["checks"] if c["code"] == "ROOM_AREA_MISMATCH"]
    assert rec["status"] == "pass"


def test_배치_기록이_감사로그에_남는다(client):
    sid = _zoned(client)
    _heads(client, sid)
    (entry,) = [e for e in _sess(client, sid).audit_entries()
                if e["event"] == "heads_placed"]
    assert entry["stage"] == "C4"
    assert entry["detail"]["rooms"] == 3 and entry["detail"]["skipped"] == 1
    assert entry["detail"]["heads"] > 0
