# -*- coding: utf-8 -*-
"""«이어서 열기» — 원본 DXF 가 정리된 저장본을 골랐을 때.

업로드 폴더는 24시간이 지나면 정리된다(`_sweep_old_upload_files`). 찍은스펙과
표시캐시는 그대로 남지만 **원본 없이는 못 연다** — 실측으로 표시캐시가 있는
키도 엔진이 `SystemExit: DXF를 못 찾음` 으로 죽는다.

종전에는 그 죽음이 잡 안에서 일어나, 사람은 진행바가 «실패» 로 바뀌고
`SystemExit: …` 이라는 문장만 보았다. 왜 못 여는지도, 어떻게 되살리는지도
없었다. 실측: 저장본 11개 중 8개가 이 상태였다.

이 시험이 지키는 것은 둘이다.
  ① 문 앞에서 막는다 — 잡을 띄우지 않고 그 자리에서 거절한다.
  ② 사유와 «되살리는 길» 을 문장에 담는다.
"""
from __future__ import annotations

import importlib
import json
import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _app():
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    for p in (_ROOT, os.path.join(_ROOT, "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    return srv.app


def _client(app):
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


@pytest.fixture()
def spec(tmp_path):
    """«원본이 사라진» 저장본 하나를 만든다 — 끝나면 걷어낸다.

    실제 작업폴더에 쓰는 이유: `_saved_keys()` 가 거기를 읽는다. 이름 앞에
    시험 표시를 달고 finally 로 반드시 지운다 — 데스크톱 G 와 같은 폴더라
    남기면 그쪽 목록에 낀다.
    """
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff
    key = "__시험_원본없음__"
    path = os.path.join(handoff.pick_out_dir(), f"{key}_찍은스펙.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"source_dxf": str(tmp_path / "없는도면.dxf")}, f,
                  ensure_ascii=False)
    try:
        yield key
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def test_원본이_없으면_목록이_그_사실을_싣는다(spec):
    c = _client(_app())
    items = (c.get("/api/module-f/saved").get_json() or {}).get("items") or []
    row = next((it for it in items if it["key"] == spec), None)
    assert row is not None, "만든 저장본이 목록에 없다"
    # ★목록에서 «빼지» 않는다 — 빠지면 「내가 찍어 둔 것이 사라졌다」 로 읽힌다.
    assert row["source_exists"] is False
    assert row["source_dxf"], "어디를 찾았는지가 없으면 되살릴 수가 없다"


def test_원본이_없으면_잡을_띄우지_않고_사유를_말한다(spec):
    c = _client(_app())
    rv = c.post("/api/module-f/reopen", json={"key": spec})
    assert rv.status_code == 409, rv.get_data(as_text=True)[:200]
    body = rv.get_json() or {}
    assert body.get("ok") is False
    msg = body.get("message") or ""
    # 왜 안 되는지 · 어떻게 되살리는지 — 둘 다 있어야 한다.
    assert "24시간" in msg, msg
    assert "다시 올리" in msg, msg
    # ★잡이 서면 안 된다. 서면 진행바가 돌다 «SystemExit» 으로 끝난다.
    assert "sid" not in body, "거절인데 세션을 만들었다"


def test_원본이_있는_저장본은_종전대로_받는다():
    """막는 자를 너무 넓게 잡지 않았나 — 성한 저장본은 그대로 열려야 한다."""
    c = _client(_app())
    items = (c.get("/api/module-f/saved").get_json() or {}).get("items") or []
    ok = next((it for it in items if it.get("source_exists")), None)
    if ok is None:
        pytest.skip("원본이 남아 있는 저장본이 없다")
    rv = c.post("/api/module-f/reopen", json={"key": ok["key"]})
    assert rv.status_code == 200, rv.get_data(as_text=True)[:200]
    assert (rv.get_json() or {}).get("sid")
