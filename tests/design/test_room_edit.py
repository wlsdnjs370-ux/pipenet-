# -*- coding: utf-8 -*-
"""지시서 §12.6 — 실 편집.

사람이 고친 것이 기록에만 남고 폴리곤에 반영되지 않으면, 화면에는 합쳐진 실이
보이는데 C4 는 갈라진 채로 헤드를 깐다. 그래서 여기서 보는 것은 두 가지다 —
**편집이 실제로 draft 를 바꾸는가**, 그리고 **모르는 것을 지어내지 않는가**.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

import routes.r30_design as r30_design  # noqa: E402
from core.design import gate as G  # noqa: E402
from core.design import room_edit as E  # noqa: E402
from core.design import session as S  # noqa: E402
from core.design.schema import BuildingDraft  # noqa: E402

# 4m × 3m 두 칸이 x=4000 에서 변을 맞댄다.
LEFT = [[0, 0], [4000, 0], [4000, 3000], [0, 3000]]
RIGHT = [[4000, 0], [8000, 0], [8000, 3000], [4000, 3000]]
FAR = [[10000, 0], [14000, 0], [14000, 3000], [10000, 3000]]
ABOVE = [[4000, 3000], [8000, 3000], [8000, 6000], [4000, 6000]]
U_SHAPE = [[0, 0], [6000, 0], [6000, 3000], [4000, 3000],
           [4000, 1000], [2000, 1000], [2000, 3000], [0, 3000]]


# ── 기하 ────────────────────────────────────────────────────────────────

def test_맞닿은_실은_공유변을_지워_합친다():
    ring = E.merge_polygons([LEFT, RIGHT])
    assert E.polygon_area_m2(ring) == pytest.approx(24.0)
    assert not any(abs(x - 4000) < 1 and 0 < y < 3000 for x, y in ring)


def test_떨어진_실은_합치지_않는다():
    """union 을 흉내 내 바깥 사각형을 만들면 없는 면적이 헤드 개수가 된다."""
    with pytest.raises(E.RoomEditError):
        E.merge_polygons([LEFT, FAR])


def test_꼭짓점만_맞닿아도_합치지_않는다():
    with pytest.raises(E.RoomEditError):
        E.merge_polygons([LEFT, ABOVE])


def test_선이_경계를_두_번_지나면_자른다():
    left, right = E.split_polygon(LEFT, [2000, -1000], [2000, 4000])
    assert E.polygon_area_m2(left) == pytest.approx(6.0)
    assert E.polygon_area_m2(right) == pytest.approx(6.0)


def test_꼭짓점을_지나는_선은_거절한다():
    with pytest.raises(E.RoomEditError):
        E.split_polygon(LEFT, [0, -1000], [0, 4000])


def test_조각이_셋_이상_나오면_거절한다():
    """ㄷ 자 실을 가로로 자르면 조각이 셋이다. 어느 둘로 묶을지는 알 수 없다."""
    with pytest.raises(E.RoomEditError):
        E.split_polygon(U_SHAPE, [-1000, 2000], [7000, 2000])


# ── draft 반영 ──────────────────────────────────────────────────────────

def _raw(rooms):
    return {
        "schema": "fncadnet.building/1",
        "source": {"floors": [{"label": "1F"}]},
        "rooms": rooms, "cores": [], "obstacles": {"status": None},
    }


def _room(rid, polygon, **kw):
    return {"id": rid, "floor": "1F", "polygon": polygon,
            "confidence": {"polygon": 0.8}, "provenance": {"polygon": "C170"}, **kw}


def _draft(rooms) -> BuildingDraft:
    return BuildingDraft.from_dict(_raw(rooms))


def test_합치면_실이_하나로_바뀐다():
    draft = _draft([_room("R-1F-001", LEFT), _room("R-1F-002", RIGHT)])
    G.apply_edits(draft, [{"op": "merge", "rooms": ["R-1F-001", "R-1F-002"],
                           "into": "R-1F-001M"}])
    assert [r.id for r in draft.rooms] == ["R-1F-001M"]
    merged = draft.rooms[0]
    assert merged.area_m2 == pytest.approx(24.0)
    assert merged.provenance["polygon"] == G.PROV_GATE


def test_합칠_때_어긋나는_값은_비운다():
    """사무실과 창고를 합쳐 한쪽 용도를 고르면 근거 없는 확정이 된다."""
    draft = _draft([
        _room("R-1F-001", LEFT, use="업무시설", provenance={"use": "GATE"}),
        _room("R-1F-002", RIGHT, use="창고시설", provenance={"use": "GATE"}),
    ])
    rec = G.apply_edits(draft, [{"op": "merge", "rooms": ["R-1F-001", "R-1F-002"]}])[0]
    merged = draft.rooms[0]
    assert merged.use is None and "use" in rec["cleared"]
    assert f"{merged.id}.use" in G.unresolved(draft)


def test_같은_값이면_확정도_함께_물려받는다():
    draft = _draft([
        _room("R-1F-001", LEFT, use="업무시설", provenance={"use": "GATE"}),
        _room("R-1F-002", RIGHT, use="업무시설", provenance={"use": "GATE"}),
    ])
    G.apply_edits(draft, [{"op": "merge", "rooms": ["R-1F-001", "R-1F-002"]}])
    merged = draft.rooms[0]
    assert merged.use == "업무시설"
    assert f"{merged.id}.use" not in G.unresolved(draft)


def test_자르면_두_실이_부모_자리에_들어간다():
    draft = _draft([_room("R-1F-000", FAR), _room("R-1F-001", LEFT),
                    _room("R-1F-002", RIGHT)])
    G.apply_edits(draft, [{"op": "split", "room": "R-1F-001",
                           "line": [[2000, -1000], [2000, 4000]]}])
    assert [r.id for r in draft.rooms] == [
        "R-1F-000", "R-1F-001a", "R-1F-001b", "R-1F-002"]
    assert sum(r.area_m2 for r in draft.rooms[1:3]) == pytest.approx(12.0)


def test_자른_자식은_부모의_확정을_이어받는다():
    draft = _draft([_room("R-1F-001", LEFT, use="업무시설", name="사무실",
                          provenance={"use": "GATE"})])
    G.apply_edits(draft, [{"op": "split", "room": "R-1F-001",
                           "line": [[2000, -1000], [2000, 4000]]}])
    assert [r.use for r in draft.rooms] == ["업무시설", "업무시설"]
    assert not [k for k in G.unresolved(draft) if k.startswith("R-1F-001")
                and k.endswith(".use")]


def test_지우면_결손에서도_사라진다():
    draft = _draft([_room("R-1F-001", LEFT), _room("R-1F-002", RIGHT)])
    G.apply_edits(draft, [{"op": "delete", "room": "R-1F-002"}])
    assert [r.id for r in draft.rooms] == ["R-1F-001"]
    assert not [k for k in G.unresolved(draft) if k.startswith("R-1F-002")]


def test_이미_있는_id_로_바꾸지_않는다():
    """조용히 다른 id 를 붙이면 같은 요청의 `values` 가 없는 실을 가리킨다."""
    draft = _draft([_room("R-1F-001", LEFT), _room("R-1F-002", RIGHT)])
    with pytest.raises(ValueError):
        G.apply_edits(draft, [{"op": "merge", "rooms": ["R-1F-001", "R-1F-002"],
                               "into": "R-1F-002"}])


def test_모르는_편집은_거절한다():
    draft = _draft([_room("R-1F-001", LEFT)])
    with pytest.raises(ValueError):
        G.apply_edits(draft, [{"op": "rotate", "room": "R-1F-001"}])


# ── 서버 ────────────────────────────────────────────────────────────────

@pytest.fixture()
def client(tmp_path):
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=True)
    app.config["DESIGN_SESSION_DIR"] = tmp_path
    return app.test_client()


def _seed(client) -> str:
    sid = client.post("/api/design/session", json={"operator": "jinwon"}).get_json()["session_id"]
    root = client.application.config["DESIGN_SESSION_DIR"]
    S.DesignSession.open(root, sid).write(
        "building.json", _raw([_room("R-1F-001", LEFT), _room("R-1F-002", RIGHT)]))
    return sid


def test_편집은_한_건씩_바로_반영하고_폴리곤을_돌려준다(client):
    sid = _seed(client)
    res = client.post("/api/design/gate/edit", json={
        "session_id": sid, "operator": "jinwon",
        "edit": {"op": "merge", "rooms": ["R-1F-001", "R-1F-002"]}})
    assert res.status_code == 200
    body = res.get_json()
    assert len(body["rooms"]) == 1
    assert body["rooms"][0]["area_m2"] == pytest.approx(24.0)

    sess = S.DesignSession.open(client.application.config["DESIGN_SESSION_DIR"], sid)
    assert "room_edit" in [e["event"] for e in sess.audit_entries()]
    saved, _version = sess.read("building.json")
    assert saved["gate"]["edits"][0]["op"] == "merge"


def test_합칠_수_없으면_확정_전에_말한다(client):
    """확정 순간에야 거절하면 사람은 이미 다 채운 뒤다."""
    sid = _seed(client)
    res = client.post("/api/design/gate/edit", json={
        "session_id": sid,
        "edit": {"op": "merge", "rooms": ["R-1F-001", "R-1F-001"]}})
    assert res.status_code == 400 and res.get_json()["code"] == "INVALID_EDIT"


def test_확정에_실린_편집은_값보다_먼저_반영된다(client):
    """실을 자른 뒤 그 자식의 용도를 같은 요청에 실을 수 있어야 한다."""
    sid = _seed(client)
    res = client.post("/api/design/gate/confirm", json={
        "session_id": sid, "operator": "jinwon",
        "edits": [{"op": "split", "room": "R-1F-001",
                   "line": [[2000, -1000], [2000, 4000]]}],
        "values": {"R-1F-001a": {"use": "업무시설"}}})
    assert res.status_code == 422        # 나머지 결손은 아직 남아 있다
    unresolved = res.get_json()["unresolved"]
    assert "R-1F-001a.use" not in unresolved
    assert "R-1F-001b.use" in unresolved
