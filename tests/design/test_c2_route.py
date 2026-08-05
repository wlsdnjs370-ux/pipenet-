# -*- coding: utf-8 -*-
"""지시서 §5·§6 — C2 결정론 기준과 C2B 103B 활성.

여기서 지키는 것은 세 가지다. **기준은 게이트를 통과해야만 나온다**, **결손은
추측하지 않고 게이트로 돌려보낸다**, **103B 는 사람이 켠다**. 마지막이 특히
중요한데, 103B 가 켜지면 NFTC 103 수평거리가 통째로 무효가 되므로 코드가 도면
용도만 보고 켜면 그 뒤 헤드 배치 전부가 근거를 잃는다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

import routes.r30_design as r30_design  # noqa: E402
from core.design import session as S  # noqa: E402
from core.design.deterministic import esfr as ESFR  # noqa: E402

_ESFR_YES = {"owner_adopts_esfr": True, "site_conditions_ok": True,
             "commodity_restricted": False}


def _building(*, use="판매시설", floors_total=12):
    return {
        "schema": "fncadnet.building/1",
        "source": {"dxf": "1F.dxf", "floors": [{"label": "1F", "dxf": "1F.dxf"}]},
        "building": {"floors_total": floors_total, "structure": "내화구조", "use": use},
        "rooms": [{
            "id": "R-1F-012", "floor": "1F", "area_m2": 82.4, "use": use,
            "ceiling": {"has_finish": False, "slab_height_mm": 3200},
            "ambient_temp_max_c": 29.0,
            "provenance": {"use": "GATE", "ceiling.has_finish": "GATE",
                           "ceiling.slab_height_mm": "default",
                           "ambient_temp_max_c": "default"},
        }],
        "cores": [{"id": "SH-01", "kind": "SHAFT", "area_m2": 2.1,
                   "confirmed": True, "provenance": {"confirmed": "GATE"}}],
        "obstacles": {"status": "partial", "provenance": {"status": "GATE"}},
        "gate": {"passed": True, "operator": "jinwon", "unresolved": []},
    }


@pytest.fixture()
def client(tmp_path):
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=True)
    app.config["DESIGN_SESSION_DIR"] = tmp_path
    return app.test_client()


def _seed(client, **kw) -> str:
    sid = client.post("/api/design/session",
                      json={"operator": "jinwon"}).get_json()["session_id"]
    root = client.application.config["DESIGN_SESSION_DIR"]
    S.DesignSession.open(root, sid).write("building.json", _building(**kw))
    return sid


def _sess(client, sid):
    return S.DesignSession.open(client.application.config["DESIGN_SESSION_DIR"], sid)


# ── C2 ──────────────────────────────────────────────────────────────────

def test_기준을_구우면_조문_출처가_함께_남는다(client):
    sid = _seed(client)
    res = client.post("/api/design/c2/constraints", json={"session_id": sid})
    assert res.status_code == 200
    body = res.get_json()
    data = body["constraints"]
    assert body["artifact"] == "constraints.json"
    # 판매시설 12층 → 기준개수 30, 내화구조 R 2.3
    assert data["scenario_head_count"] == 30
    assert data["horizontal_distance_m"] == 2.3
    assert {t["field"] for t in data["trace"]} >= {"scenario_head_count",
                                                   "horizontal_distance_m"}
    assert all(t["text_hash"] for t in data["trace"])

    sess = _sess(client, sid)
    assert sess.status()["meta"]["stage"] == "c3"
    assert "constraints_built" in [e["event"] for e in sess.audit_entries()]


def test_기준은_브라우저가_읽는_JSON_으로_나간다(client):
    """보 이격표 마지막 구간의 상한은 `inf` 다.

    파이썬 json 은 그것을 `Infinity` 로 적지만 JSON 에는 없는 낱말이라, 브라우저는
    응답 전체를 못 읽는다 — 서버는 200 인데 화면만 조용히 비는 자리였다.
    """
    sid = _seed(client)
    res = client.post("/api/design/c2/constraints", json={"session_id": sid})
    body = json.loads(res.get_data(as_text=True),
                      parse_constant=lambda name: pytest.fail(f"JSON 밖의 값: {name}"))
    last = body["constraints"]["beam_clearance_table"][-1]
    # `null` 은 이 코드베이스에서 "모른다" 다. 상한 없는 구간은 조문대로 하한으로.
    assert "horizontal_lt_m" not in last and last["horizontal_gte_m"] == 1.5


def test_확정되지_않은_사실은_추측하지_않고_게이트로_돌려보낸다(client):
    sid = _seed(client)
    raw, _v = _sess(client, sid).read("building.json")
    raw["rooms"][0]["ceiling"] = {"has_finish": False}   # 천장고가 없다
    _sess(client, sid).write("building.json", raw)

    res = client.post("/api/design/c2/constraints", json={"session_id": sid})
    assert res.status_code == 422
    body = res.get_json()
    assert body["code"] == "MISSING_BUILDING_FACT"
    assert body["field"] == "rooms[].ceiling"


def test_게이트를_통과하지_않으면_기준을_굽지_않는다(client):
    sid = _seed(client)
    raw, _v = _sess(client, sid).read("building.json")
    raw["gate"] = {"passed": False, "unresolved": ["R-1F-012.use"]}
    _sess(client, sid).write("building.json", raw)

    res = client.post("/api/design/c2/constraints", json={"session_id": sid})
    assert res.status_code == 409
    assert res.get_json()["code"] == "GATE_NOT_PASSED"


def test_같은_기준을_다시_굽지_않는다(client):
    """새로고침이 세대를 늘리면 감사 기록에서 진짜 재검토가 묻힌다."""
    sid = _seed(client)
    client.post("/api/design/c2/constraints", json={"session_id": sid})
    res = client.post("/api/design/c2/constraints", json={"session_id": sid}).get_json()
    assert res["artifact"] == "constraints.json" and res["rebaked"] is False
    assert not _sess(client, sid).path("constraints.v2.json").exists()


def test_달라진_기준은_옛_기준을_덮지_않는다(client):
    """§16.3 — 어떤 기준으로 설계했는지가 감사 대상이라 옆에 새로 쓴다."""
    sid = _seed(client)
    client.post("/api/design/c2/constraints", json={"session_id": sid})

    sess = _sess(client, sid)
    raw, version = sess.read("building.json")
    raw["building"]["structure"] = "비내화구조"     # R 2.3 → 2.1
    sess.write("building.json", raw, if_version=version)

    res = client.post("/api/design/c2/constraints", json={"session_id": sid}).get_json()
    assert res["artifact"] == "constraints.v2.json" and res["rebaked"] is True
    assert res["constraints"]["horizontal_distance_m"] == 2.1

    sess = _sess(client, sid)
    assert sess.read("constraints.json")[0]["horizontal_distance_m"] == 2.3
    assert sess.status()["meta"]["constraints"] == "constraints.v2.json"


# ── C2B ─────────────────────────────────────────────────────────────────

def test_랙크식_창고라고_103B_가_켜지지_않는다(client):
    """G2. 기본 트랙은 NFTC 103 2.7.2 이고 103B 는 별도 설비 종류다."""
    sid = _seed(client, use="랙크식창고")
    res = client.post("/api/design/c2/constraints", json={"session_id": sid})
    data = res.get_json()["constraints"]
    assert data.get("esfr") is None
    assert data["horizontal_distance_m"] == 2.5      # 랙크식 R, 103B 아님


def test_103B_는_사람이_켜고_그_순간_수평거리가_무효가_된다(client):
    sid = _seed(client, use="랙크식창고")
    client.post("/api/design/c2/constraints", json={"session_id": sid})

    res = client.post("/api/design/c2b/esfr",
                      json={"session_id": sid, "operator": "jinwon", **_ESFR_YES})
    assert res.status_code == 200
    body = res.get_json()
    assert body["esfr"]["active"] is True and body["esfr"]["operator"] == "jinwon"
    # 기준도 같이 다시 구워야 두 산출물이 어긋나지 않는다.
    data = body["constraints"]
    assert body["artifact"] == "constraints.v2.json"
    for field in ESFR.VOID_FIELDS:
        assert data[field] is None
    assert {t["field"] for t in data["trace"]}.isdisjoint(ESFR.VOID_FIELDS)


def test_103B_를_끄면_NFTC_103_값이_그대로_남는다(client):
    sid = _seed(client)
    res = client.post("/api/design/c2b/esfr",
                      json={"session_id": sid, "operator": "jinwon",
                            **{**_ESFR_YES, "owner_adopts_esfr": False}})
    assert res.get_json()["esfr"]["active"] is False
    # 기준을 굽기 전이라 다시 구울 것도 없다.
    assert "constraints" not in res.get_json()

    data = client.post("/api/design/c2/constraints",
                       json={"session_id": sid}).get_json()["constraints"]
    assert data["esfr"]["active"] is False
    assert data["horizontal_distance_m"] == 2.3


def test_확정자_없이는_103B_를_판단하지_않는다(client):
    sid = _seed(client)
    res = client.post("/api/design/c2b/esfr", json={"session_id": sid, **_ESFR_YES})
    assert res.status_code == 400
    assert res.get_json()["code"] == "OPERATOR_REQUIRED"


def test_묻지_않은_조건을_아니라고_읽지_않는다(client):
    """빠진 값을 False 로 흘리면 '설치장소 조건 미충족'이 조용히 확정된다."""
    sid = _seed(client)
    res = client.post("/api/design/c2b/esfr",
                      json={"session_id": sid, "operator": "jinwon",
                            "owner_adopts_esfr": True})
    assert res.status_code == 400
    body = res.get_json()
    assert body["code"] == "ESFR_INPUT_REQUIRED"
    assert set(body["fields"]) == {"site_conditions_ok", "commodity_restricted"}
