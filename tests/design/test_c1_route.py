# -*- coding: utf-8 -*-
"""지시서 §11.1 / §11.2 — C1 인식 엔드포인트.

인식 알고리즘은 `test_pipeline.py` 와 도면 벤치마크가 본다. 여기서 지키는 것은
**라우트가 사슬의 결과를 왜곡하지 않는가** 다. 셋을 본다 — 막힌 인식을 성공으로
저장하지 않는가(막힌 채 `building.json` 을 쓰면 GATE 가 빈 도면을 확정 대상으로
삼는다), 캐시가 파라미터를 무시하지 않는가, 그리고 도면 토큰이 업로드 폴더를
벗어나지 않는가.
"""
from __future__ import annotations

import gzip
import hashlib
import json
import sys
from pathlib import Path

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

import routes.r30_design as r30_design  # noqa: E402
from core.design import session as S  # noqa: E402

_CACHE_VERSION = "v4"


def _room_dxf(layer="A-WALL", width=6000.0, height=5000.0, thickness=200.0):
    """두 겹 선으로 그린 방 하나 — C150 이 짝지을 수 있는 최소 도면."""
    t = thickness
    return [{"t": "L", "l": layer, "p": [x0, y0, x1, y1]} for x0, y0, x1, y1 in (
        (0, 0, width, 0), (0, height, width, height),
        (0, 0, 0, height), (width, 0, width, height),
        (t, t, width - t, t), (t, height - t, width - t, height - t),
        (t, t, t, height - t), (width - t, t, width - t, height - t))]


def _clutter_dxf(layer="A-FURN", n=12):
    """평행쌍은 있지만 벽으로 보기엔 너무 짧은 선들 — WALL 이 안 나오는 도면."""
    out = []
    for i in range(n):
        y = i * 900.0
        out.append({"t": "L", "l": layer, "p": [0.0, y, 400.0, y]})
        out.append({"t": "L", "l": layer, "p": [0.0, y + 100.0, 400.0, y + 100.0]})
    return out


@pytest.fixture()
def client(tmp_path):
    upload = tmp_path / "uploads"
    cache = upload / "_inspect_cache"
    cache.mkdir(parents=True)
    app = Flask(__name__)
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path / "sessions",
                        UPLOAD_DIR=upload, INSPECT_CACHE_DIR=cache,
                        INSPECT_CACHE_VERSION=_CACHE_VERSION, enabled=True)
    app.config["UPLOAD_DIR"] = upload
    app.config["INSPECT_CACHE_DIR"] = cache
    app.config["DESIGN_SESSION_DIR"] = tmp_path / "sessions"
    return app.test_client()


def _inspected(client, entities, *, name="plan.dxf") -> str:
    """도면 한 장을 올리고 inspect 렌더 캐시까지 남긴 상태를 만든다."""
    path = client.application.config["UPLOAD_DIR"] / name
    path.write_bytes(name.encode("utf-8"))
    key = f"{_CACHE_VERSION}_{hashlib.sha256(path.read_bytes()).hexdigest()}"
    cache = client.application.config["INSPECT_CACHE_DIR"]
    with gzip.open(cache / f"{key}.entities.ndjson.gz", "wt", encoding="utf-8") as gz:
        gz.write(json.dumps({"type": "progress", "entities": entities}) + "\n")
    (cache / f"{key}.meta.json").write_text(json.dumps({
        "bbox": {"x_min": -5000.0, "y_min": -5000.0, "x_max": 25000.0, "y_max": 25000.0},
        "counts": {"total_entities": len(entities)},
    }), encoding="utf-8")
    return name


def _session(client) -> str:
    return client.post("/api/design/session", json={}).get_json()["session_id"]


def _recognize(client, sid, token, **body) -> list[dict]:
    res = client.post("/api/design/c1/recognize",
                      json={"session_id": sid, "dxf_token": token, **body})
    assert res.status_code == 200, res.get_data(as_text=True)
    return [json.loads(line) for line in res.get_data(as_text=True).splitlines() if line]


# ── 정상 경로 ───────────────────────────────────────────────────────────

def test_스트림이_11_1_순서로_나온다(client):
    sid, token = _session(client), _inspected(client, _room_dxf())
    types = [m["type"] for m in _recognize(client, sid, token, wall_layers=["A-WALL"])]
    # `building` 은 사슬이 라우트에 넘기는 내부 메시지다. 밖으로 새면 클라이언트가
    # 세션에 저장되기도 전의 draft 를 최종본으로 읽는다.
    assert types == ["parse", "fingerprint", "phase", "centerlines", "virtual_edges",
                     "rooms", "cores", "result"]


def test_인식이_끝나면_building_json_과_게이트_주소가_생긴다(client):
    sid, token = _session(client), _inspected(client, _room_dxf())
    result = _recognize(client, sid, token, wall_layers=["A-WALL"])[-1]
    assert result["ok"] is True
    assert result["gate_items_url"] == f"/api/design/c1/gate_items/{sid}"
    assert result["counts"]["rooms"] == 1
    assert result["wall_source"] == "operator"

    sess = S.DesignSession.open(client.application.config["DESIGN_SESSION_DIR"], sid)
    draft, version = sess.read("building.json")
    assert version == result["version"]
    assert len(draft["rooms"]) == 1
    assert client.get(f"/api/design/c1/gate_items/{sid}").status_code == 200


def test_단계별_소요_시간이_result_에_실린다(client):
    """부록 C.1 예산을 어디서 넘겼는지 알 수 없으면 예산이 아니다."""
    sid, token = _session(client), _inspected(client, _room_dxf())
    result = _recognize(client, sid, token, wall_layers=["A-WALL"])[-1]
    assert [s["name"] for s in result["stages"]] == [
        "C130/C140", "C150", "C160", "C170", "C180", "C190"]


# ── 막혔을 때 (§3.2 판정표의 한계) ──────────────────────────────────────

def test_WALL_이_없으면_저장하지_않고_막혔다고_말한다(client):
    """빈 draft 를 저장하면 GATE 가 방 0개짜리 도면을 확정 대상으로 삼는다."""
    sid, token = _session(client), _inspected(client, _clutter_dxf())
    messages = _recognize(client, sid, token)
    result = messages[-1]
    assert result["ok"] is False
    assert result["blocked"] == "no_wall_layer"
    assert "gate_items_url" not in result

    phase = next(m for m in messages if m["type"] == "phase")
    assert [c["name"] for c in phase["candidates"]] == ["A-FURN"]
    assert client.get(f"/api/design/c1/gate_items/{sid}").status_code == 409


# ── 캐시 (§11.2) ────────────────────────────────────────────────────────

def test_같은_도면_같은_파라미터는_캐시를_쓴다(client):
    sid, token = _session(client), _inspected(client, _room_dxf())
    first = _recognize(client, sid, token, wall_layers=["A-WALL"])
    second = _recognize(client, sid, token, wall_layers=["A-WALL"])
    assert not any(m.get("cached") for m in first)
    assert next(m for m in second if m["type"] == "phase")["cached"] is True
    assert second[-1]["counts"] == first[-1]["counts"]
    assert "centerlines" not in [m["type"] for m in second]


def test_고른_WALL_레이어가_다르면_캐시가_갈린다(client):
    """파라미터가 키에 안 들어가면 레이어를 바꿔도 옛 결과가 나온다."""
    sid = _session(client)
    token = _inspected(client, _room_dxf() + _clutter_dxf())
    auto = _recognize(client, sid, token)[-1]
    picked = _recognize(client, sid, token, wall_layers=["A-FURN"])[-1]
    assert auto["wall_source"] == "C140" and auto["counts"]["rooms"] == 1
    assert picked["wall_layers"] == ["A-FURN"] and picked["counts"]["rooms"] == 0


# ── 거절 ────────────────────────────────────────────────────────────────

def test_세션이_없으면_404(client):
    token = _inspected(client, _room_dxf())
    res = client.post("/api/design/c1/recognize",
                      json={"session_id": "00000000-0000-4000-8000-000000000000",
                            "dxf_token": token})
    assert res.status_code == 404


@pytest.mark.parametrize("token", ["../secret.dxf", "plan.txt", "", "없는파일.dxf"])
def test_업로드_폴더_밖의_토큰은_거절한다(client, token):
    sid = _session(client)
    _inspected(client, _room_dxf())
    res = client.post("/api/design/c1/recognize",
                      json={"session_id": sid, "dxf_token": token})
    assert res.status_code == 400
    assert res.get_json()["code"] == "DXF_NOT_FOUND"


def test_화면에_올리지_않은_도면은_409로_멈춘다(client):
    """캐시가 없으면 DXF 재파싱으로 40초를 말없이 더 쓰는 대신 이유를 말한다."""
    sid = _session(client)
    (client.application.config["UPLOAD_DIR"] / "cold.dxf").write_bytes(b"cold")
    res = client.post("/api/design/c1/recognize",
                      json={"session_id": sid, "dxf_token": "cold.dxf"})
    assert res.status_code == 409
    assert res.get_json()["code"] == "DXF_NOT_INSPECTED"
