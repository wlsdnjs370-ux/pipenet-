# -*- coding: utf-8 -*-
"""지시서 §4 — 인간 확정 게이트.

게이트의 요점은 두 가지다. **인식기가 채운 값은 제안이지 확정이 아니라는 것**과,
**UI 잠금이 아니라 서버가 막는다는 것**. 둘 중 하나라도 새면 근거 없는 값이
기준개수·수원·관경까지 그대로 흘러간다.
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
from core.design import session as S  # noqa: E402
from core.design.schema import BuildingDraft  # noqa: E402


def _raw_building(*, floor_height_mm=None):
    return {
        "schema": "fncadnet.building/1",
        "source": {"dxf": "1F.dxf",
                   "floors": [{"label": "1F", "dxf": "1F.dxf", "height_mm": floor_height_mm}]},
        "building": {"floors_total": None, "structure": None, "use": None},
        "rooms": [{
            "id": "R-1F-012", "floor": "1F", "area_m2": 82.4,
            "name": "사무실", "use": "업무시설",
            "ceiling": {"has_finish": None},
            "confidence": {"use": 0.62}, "provenance": {"use": "C180"},
        }, {
            # 벽 틈으로 새어 나온 잡 face. 사람이 지우는 대상이다.
            "id": "R-1F-013", "floor": "1F", "area_m2": 0.9,
        }],
        "cores": [{"id": "SH-01", "kind": "SHAFT", "area_m2": 2.1, "confidence": 0.4}],
        "obstacles": {"status": None},
    }


def _draft(**kw) -> BuildingDraft:
    return BuildingDraft.from_dict(_raw_building(**kw))


# ── 결손 판정 ───────────────────────────────────────────────────────────

def test_인식기가_채운_값은_확정이_아니다():
    """실명 귀속이 용도를 제안했어도 사람이 확정하기 전까지는 결손이다(§3.6)."""
    draft = _draft()
    assert draft.room("R-1F-012").use == "업무시설"
    assert "R-1F-012.use" in G.unresolved(draft)


def test_사람이_확정하면_결손에서_빠진다():
    draft = _draft()
    G.apply_values(draft, {"R-1F-012": {"use": "판매시설"}})
    assert "R-1F-012.use" not in G.unresolved(draft)
    assert draft.room("R-1F-012").provenance["use"] == G.PROV_GATE


def test_반자가_없으면_반자고를_묻지_않는다():
    draft = _draft()
    G.apply_values(draft, {"R-1F-012": {"ceiling.has_finish": False}})
    assert "R-1F-012.ceiling.finish_height_mm" not in G.unresolved(draft)
    G.apply_values(draft, {"R-1F-012": {"ceiling.has_finish": True}})
    assert "R-1F-012.ceiling.finish_height_mm" in G.unresolved(draft)


def test_건물_사실과_코어_확정도_결손이다():
    missing = G.unresolved(_draft())
    assert {"building.floors_total", "building.structure", "building.use",
            "obstacles.status", "SH-01.confirmed"} <= set(missing)


def test_장애물_미확보도_확정이다():
    """'none' 은 정보가 없다는 뜻이지 확정하지 않았다는 뜻이 아니다(§4.2)."""
    draft = _draft()
    G.apply_values(draft, {"obstacles.status": {"obstacles.status": "none"}})
    assert "obstacles.status" not in G.unresolved(draft)


def _by_field(applied):
    return {a["field"]: a for a in applied}


def test_천장고_기본값은_층고가_있을_때만_채운다():
    """근거 없는 채움은 G9 위반이다. 층고를 모르면 결손으로 남긴다."""
    assert "ceiling.slab_height_mm" not in _by_field(G.apply_defaults(_draft()))
    draft = _draft(floor_height_mm=3200)
    applied = _by_field(G.apply_defaults(draft))
    assert applied["ceiling.slab_height_mm"]["value"] == 3200
    assert draft.room("R-1F-012").provenance["ceiling.slab_height_mm"] == G.PROV_DEFAULT
    assert "R-1F-012.ceiling.slab_height_mm" not in G.unresolved(draft)


def test_주위온도는_상온_대푯값을_근거와_함께_채운다():
    """[문서정합] §4 표는 '용도별 기본값'이라 적었으나 소방 용도 분류에 열원 정보가
    없다. 표시온도는 39℃ 미만이 한 구간이라 상온 대푯값 하나로 결과가 같고, 그
    구간을 벗어나는 실(보일러실 등)은 사람이 게이트에서 직접 올린다."""
    draft = _draft()
    applied = _by_field(G.apply_defaults(draft))["ambient_temp_max_c"]
    assert applied["value"] == G.AMBIENT_DEFAULT_C and applied["basis"]
    room = draft.room("R-1F-012")
    assert room.provenance["ambient_temp_max_c"] == G.PROV_DEFAULT

    # 사람이 올린 값은 기본값이 덮지 않는다.
    G.apply_values(draft, {"R-1F-012": {"ambient_temp_max_c": 45}})
    G.apply_defaults(draft)
    assert room.ambient_temp_max_c == 45


def test_결손_항목은_한_번에_묶어서_낸다():
    """파이프라인 중간에 사람을 계속 부르면 자동화가 아니라 방해다(§4.2)."""
    items = G.gate_items(_draft())
    by_field = {g["field"]: g for g in items["groups"]}
    assert by_field["use"]["suggestion"]["R-1F-012"] == {"value": "업무시설", "confidence": 0.62}
    assert "floor" in by_field["ceiling.has_finish"]["bulk_apply"]
    assert by_field["obstacles.status"]["options"] == list(G.OBSTACLE_STATUSES)
    assert items["total"] == len(G.unresolved(_draft()))


def test_항목_목록은_비어_있어도_낸다():
    """`groups` 는 지금 비어 있는 것만 담는다. 화면이 표의 열을 세우려면 항목 자체가
    필요하고, 무엇이 무엇에 걸리는지도 화면이 다시 구현하지 않고 여기서 받아야 한다."""
    items = G.gate_items(_draft())
    fields = {f["field"]: f for f in items["fields"]}
    assert set(fields) == {s.field for s in G.ROOM_FIELDS + G.BUILDING_FIELDS}
    assert fields["ceiling.finish_height_mm"]["applies_when"] == {
        "field": "ceiling.has_finish", "value": True}
    # 반자를 아직 아무도 확정하지 않아 반자고 그룹은 없다 — 그래도 항목에는 있다.
    assert "ceiling.finish_height_mm" not in {g["field"] for g in items["groups"]}


def test_모르는_항목은_거절한다():
    with pytest.raises(ValueError):
        G.apply_values(_draft(), {"R-1F-012": {"ceiling.color": "white"}})


# ── 서버 강제 ───────────────────────────────────────────────────────────

@pytest.fixture()
def client(tmp_path):
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=True)
    app.config["DESIGN_SESSION_DIR"] = tmp_path
    return app.test_client()


def _seed(client, **kw) -> str:
    sid = client.post("/api/design/session", json={"operator": "jinwon"}).get_json()["session_id"]
    root = client.application.config["DESIGN_SESSION_DIR"]
    S.DesignSession.open(root, sid).write("building.json", _raw_building(**kw))
    return sid


def test_플래그가_꺼져_있으면_라우트가_없다(tmp_path):
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=False)
    assert app.test_client().post("/api/design/session").status_code == 404


@pytest.mark.parametrize("path,method", [
    ("/api/design/c2/constraints", "post"), ("/api/design/c2b/esfr", "post"),
    ("/api/design/c3/valves", "post"), ("/api/design/c3/zones", "post"),
    ("/api/design/c4/heads", "post"), ("/api/design/c5/route", "post"),
    ("/api/design/emit", "post"),
])
def test_게이트_전에는_c2_이후가_전부_409(client, path, method):
    sid = _seed(client)
    res = getattr(client, method)(path, json={"session_id": sid})
    assert res.status_code == 409
    body = res.get_json()
    assert body["code"] == "GATE_NOT_PASSED" and body["unresolved"]


def test_검사_조회도_게이트_뒤에_있다(client):
    sid = _seed(client)
    assert client.get(f"/api/design/checks/{sid}").status_code == 409


def test_결손이_남으면_통과시키지_않는다(client):
    sid = _seed(client)
    res = client.post("/api/design/gate/confirm",
                      json={"session_id": sid, "operator": "jinwon",
                            "values": {"R-1F-012": {"use": "업무시설"}}})
    assert res.status_code == 422
    assert res.get_json()["code"] == "GATE_INCOMPLETE"


def _confirm_all(client, sid, **extra):
    return client.post("/api/design/gate/confirm", json={
        "session_id": sid, "operator": "jinwon",
        "values": {
            "R-1F-012": {"use": "판매시설", "ceiling.has_finish": False},
            "SH-01": {"confirmed": True},
            "building.floors_total": {"building.floors_total": 12},
            "building.structure": {"building.structure": "내화구조"},
            "building.use": {"building.use": "판매시설"},
            "obstacles.status": {"obstacles.status": "partial"},
        },
        # 지운 실은 결손에서도 사라져야 한다 — 편집이 결손 판정보다 먼저 돈다.
        "edits": [{"op": "delete", "room": "R-1F-013"}],
        **extra,
    })


def test_전부_확정하면_게이트가_열리고_기록이_남는다(client):
    sid = _seed(client, floor_height_mm=3200)
    res = _confirm_all(client, sid)
    assert res.status_code == 200 and res.get_json()["passed"] is True

    root = client.application.config["DESIGN_SESSION_DIR"]
    sess = S.DesignSession.open(root, sid)
    events = [e["event"] for e in sess.audit_entries()]
    assert "use_override" in events        # 제안 0.62 → 판매시설. 감리 질의 1순위
    assert "default_applied" in events     # 천장고를 층고에서 가져온 근거
    assert "room_edit" in events           # C1 개선의 학습 데이터
    assert "passed" in events
    assert sess.status()["meta"]["gate_passed"] is True

    res = client.post("/api/design/c5/route", json={"session_id": sid})
    assert res.status_code == 501         # 게이트는 열렸고 C5 는 아직 없다


def test_다른_탭이_먼저_저장했으면_덮어쓰지_않는다(client):
    sid = _seed(client, floor_height_mm=3200)
    res = _confirm_all(client, sid, if_version=0)
    assert res.status_code == 409
    body = res.get_json()
    assert body["code"] == "VERSION_CONFLICT" and body["current_version"] == 1


def test_c1_전에는_확정할_대상이_없다(client):
    sid = client.post("/api/design/session").get_json()["session_id"]
    assert client.get(f"/api/design/c1/gate_items/{sid}").get_json()["code"] == "C1_NOT_DONE"
