# -*- coding: utf-8 -*-
"""지시서 §11·§13.6·§15(PR-10) — 하류 체이닝.

지키는 것은 셋이다. **막는 문제를 안은 망은 방출하지 않는다**(단계는 서버가
정하고 여기서는 그 값만 믿는다), **세 포맷이 다 나온다**(SDF·KFP·HAS. 하나가
빠졌는데 성공으로 읽히면 그대로 솔버에 들어간다), **낸 것이 살아 돌아온다**
(방출한 SDF 를 다시 읽어 노드·배관·헤드·연장·관경을 맞대어 본다).

도면은 test_c5_route 와 같다. 모듈 A 방출기(`remote30_full_network`)가 평면
import 라 `core` 도 sys.path 에 올린다 — 서버(`대조 서버.py`)가 이미 지는 부담이다.
"""
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_ROOT / "core"))

import routes.r30_design as r30_design  # noqa: E402
from core.design import session as S  # noqa: E402
from test_c5_route import _REQ_OK, _building  # noqa: E402


@pytest.fixture()
def client(tmp_path):
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=True)
    app.config["DESIGN_SESSION_DIR"] = tmp_path
    return app.test_client()


def _sess(client, sid):
    return S.DesignSession.open(client.application.config["DESIGN_SESSION_DIR"], sid)


def _routed(client) -> str:
    """C5 라우팅까지 끝난 세션. 여기가 방출의 출발점이다."""
    sid = client.post("/api/design/session",
                      json={"operator": "jinwon"}).get_json()["session_id"]
    _sess(client, sid).write("building.json", _building())
    assert client.post("/api/design/c2/constraints",
                       json={"session_id": sid}).status_code == 200
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
    assert client.post("/api/design/c5/route",
                       json={"session_id": sid,
                             "operator": "jinwon"}).status_code == 200
    return sid


def _emit(client, sid):
    return client.post("/api/design/emit",
                       json={"session_id": sid, "operator": "jinwon"})


# ── 전제 ────────────────────────────────────────────────────────────────

def test_배관망이_없으면_방출하지_않는다(client):
    sid = client.post("/api/design/session",
                      json={"operator": "jinwon"}).get_json()["session_id"]
    _sess(client, sid).write("building.json", _building())
    res = _emit(client, sid)
    assert res.status_code == 409
    assert res.get_json()["code"] == "NETWORK_REQUIRED"


def test_앞_단계가_막고_있으면_방출하지_않는다(client):
    """배관망 파일이 있다는 것과 그 망이 내보내도 되는 망이라는 것은 다르다."""
    sid = _routed(client)
    _sess(client, sid).update_meta(stage="c3")
    res = _emit(client, sid)
    assert res.status_code == 409
    body = res.get_json()
    assert body["code"] == "EMIT_NOT_READY" and body["stage"] == "c3"


# ── 산출물 ──────────────────────────────────────────────────────────────

def test_망마다_네_파일이_나온다(client):
    """SDF·KFP·HAS 3종에 SLF 를 더한 넷. PIPENET 은 .sdf 옆의 .slf 로 내경을 찾는다."""
    sid = _routed(client)
    body = _emit(client, sid).get_json()
    (bundle,) = body["bundles"]
    suffixes = sorted(Path(f["name"]).suffix for f in bundle["files"])
    assert suffixes == [".has", ".kfp", ".sdf", ".slf"]
    assert all(f["bytes"] > 0 for f in bundle["files"])

    emit_dir = _sess(client, sid).path("emit")
    assert all((emit_dir / f["name"]).is_file() for f in bundle["files"])


def test_왕복_검증이_일치한다(client):
    """§13.6 — 낸 그래프가 SDF 를 돌아 그대로 살아 오는지. 신축배관은 모듈 A 가
    실배관 한 토막으로 펴므로 그만큼 느는 것이 정상이다."""
    sid = _routed(client)
    (bundle,) = _emit(client, sid).get_json()["bundles"]
    rt = bundle["round_trip"]
    assert rt["ok"] is True
    assert rt["nodes"][1] == rt["nodes"][0] + rt["flex"]
    assert rt["pipes"][1] == rt["pipes"][0] + rt["flex"]
    assert rt["heads"][0] == rt["heads"][1] > 0
    assert rt["dia_match"][0] == rt["dia_match"][1] > 0
    assert rt["length_m"][1] == pytest.approx(
        rt["length_m"][0] + rt["flex_length_m"], abs=0.01)


def test_배관_한_본이_사라지면_불일치로_잡는다():
    """대조기 자체를 재 본다. 실제 방출로는 이 상태를 만들 수 없다 — 모듈 A 가
    라벨 중복을 거절하고, 필드 하나를 흔들면 SDF 도 같이 흔들려 양쪽이 함께 움직인다.
    그래서 없어진 배관 한 본을 손으로 만들어 대조기가 그것을 보는지 확인한다."""
    combined = SimpleNamespace(
        nodes=[{}, {}], nozzles=[{}], equipment=[],
        pipes=[{"label": "1", "dia": 25, "length": 2.0},
               {"label": "2", "dia": 25, "length": 3.0}])
    parsed = SimpleNamespace(
        nodes={"10": SimpleNamespace(kind="head"), "11": SimpleNamespace(kind="tee")},
        pipes={"P1": SimpleNamespace(id="P1", length_m=2.0, nominal_mm=25)})

    rt = r30_design._round_trip(combined, parsed, [2.0])
    assert rt["ok"] is False
    assert rt["pipes"] == [2, 1] and rt["dia_match"] == [1, 2]


def test_어긋나면_플래그로_말한다(client, monkeypatch):
    """왕복이 어긋난 파일을 조용히 내주면 그대로 솔버에 들어간다."""
    sid = _routed(client)
    assert _emit(client, sid).get_json()["flags"] == []

    monkeypatch.setattr(r30_design, "_round_trip",
                        lambda combined, net, sdf_lengths: {"ok": False})
    body = _emit(client, sid).get_json()
    assert [f["code"] for f in body["flags"]] == ["ROUND_TRIP_MISMATCH"]
    assert body["bundles"][0]["round_trip"]["ok"] is False
    stored, _v = _sess(client, sid).read("emit.json")
    assert stored["flags"] == body["flags"]


# ── 내려받기 ────────────────────────────────────────────────────────────

def test_방출한_파일만_내려받는다(client):
    sid = _routed(client)
    (bundle,) = _emit(client, sid).get_json()["bundles"]
    name = next(f["name"] for f in bundle["files"] if f["name"].endswith(".sdf"))

    res = client.get(f"/api/design/emit/{sid}/{name}")
    assert res.status_code == 200 and res.data.lstrip().startswith(b"<")

    assert client.get(f"/api/design/emit/{sid}/meta.json").status_code == 404


def test_다시_방출하면_옛_파일이_남지_않는다(client):
    """망을 다시 그렸는데 옛 파일이 내려가면 화면과 손에 쥔 파일이 갈린다."""
    sid = _routed(client)
    _emit(client, sid)
    emit_dir = _sess(client, sid).path("emit")
    stale = emit_dir / "net99_OLD.sdf"
    stale.write_text("옛 방출물", encoding="utf-8")

    _emit(client, sid)
    assert not stale.exists()
    assert client.get(f"/api/design/emit/{sid}/{stale.name}").status_code == 404


def test_방출_기록이_감사로그에_남는다(client):
    sid = _routed(client)
    _emit(client, sid)
    (entry,) = [e for e in _sess(client, sid).audit_entries()
                if e["event"] == "files_emitted"]
    assert entry["stage"] == "EMIT"
    assert entry["detail"]["networks"] == 1
    assert entry["detail"]["round_trip_ok"] is True
    assert len(entry["detail"]["files"]) == 4


def test_산출물_목록에_방출이_실린다(client):
    sid = _routed(client)
    _emit(client, sid)
    status = client.get(f"/api/design/session/{sid}").get_json()
    assert status["artifacts"]["emit.json"] is not None
