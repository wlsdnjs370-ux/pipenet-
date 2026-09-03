# -*- coding: utf-8 -*-
"""모듈 F — «살아 있는 세션 + 이상한 입력» 에서 500 이 나면 안 된다.

기존 안전망(`test_모든_라우트가_예외를_안_던진다`)은 세션 없이 두드리므로
대부분 410 에서 멈춘다 — **핸들러 본문을 안 지난다.** 그 한계는 그 파일이
스스로 적어 두었다. 여기서는 손질까지 간 진짜 세션에 이상한 값을 넣어
본문을 지나게 한다.

500 은 «사람이 읽을 수 없는 실패» 다. 이 저장소의 규약은 실패도 문장으로
말하는 것이므로 500 이 나오면 그 자리가 결함이다.

■ 이 시험이 처음 잡은 것
    POST /edit/click  {x: 1e308, y: -1e308}
      → 엔진이 점–선분 거리를 **제곱** 으로 재는데(user_net._pt_seg_d2)
        float 은 1.3e154 를 넘겨 제곱하면 OverflowError 를 던진다.
        그것이 그대로 «서버 오류: OverflowError» 로 나갔다.
    엔진은 안 고친다 — 데스크톱 G 는 그런 좌표를 만들 수 없고 웹은 아무
    문자열이나 들어온다. 그래서 문 앞에서 막는다(`common._check_xy`).
"""
from __future__ import annotations

import importlib
import json
import os
import sys
import time

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 무거운 잡을 띄우는 자리는 뺀다 — 본문 검증은 잡 «앞» 에서 끝나고,
# 여기서 태우면 시험이 몇 분짜리가 된다.
_HEAVY = {
    "/api/module-f/convert/run", "/api/module-f/design/build",
    "/api/module-f/auto/run", "/api/module-f/auto/network",
    "/api/module-f/auto/handoff", "/api/module-f/edit/autojoin/apply",
    "/api/module-f/edit/anchor-click", "/api/module-f/merge/build",
    "/api/module-f/merge/emit", "/api/module-f/pick/adopt",
    "/api/module-f/pick/suggest", "/api/module-f/open",
    "/api/module-f/slot/open", "/api/module-f/reopen",
    "/api/module-f/sub/extract", "/api/module-f/job/stream",
}

# «형식은 맞지만 뜻이 안 되는» 값들. 좌표는 셀 수 없는 수까지 넣는다.
_WEIRD = [
    {},
    {"x": "abc", "y": None, "max_d": []},
    {"x": 1e308, "y": -1e308, "max_d": 0},          # ★제곱에서 터지던 자리
    {"x": float("inf"), "y": 0, "max_d": 1},
    {"k": -5}, {"k": 10 ** 9}, {"k": "삼십"},
    {"zones": "not-a-list"}, {"zones": [[1, 2]]}, {"zones": [{}]},
    {"sheet": 99999}, {"source": "Z999"},
    {"eps_mm": -1}, {"eps_mm": "넓게"},
    {"mode": 123}, {"kind": []}, {"slot": {"a": 1}},
    {"heads": []}, {"heads": {"indices": ["x"]}}, {"heads": {"conf_min": "높음"}},
    {"dto": "문자열"}, {"outputs": "전부"},
    {"rows": "문자열"}, {"rows": [1, 2, 3]}, {"rows": [{"a": "x"}]},
    {"waypoints": "여기"}, {"waypoints": [[1]]},
    {"ceiling_m": "높이"}, {"snap_tolerance_mm": -3},
    {"layers": 5}, {"pump": "펌프"}, {"source_drop_m": "깊이"},
]


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


def _idle(c, sid, limit=2400):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


@pytest.fixture(scope="module")
def live():
    """손질까지 간 세션 하나 — 원본이 남은 저장본으로."""
    app = _app()
    c = _client(app)
    items = (c.get("/api/module-f/saved").get_json() or {}).get("items") or []
    ok = next((it for it in items if it.get("source_exists")), None)
    if ok is None:
        pytest.skip("원본이 남은 저장본이 없다 — 본문을 지나는 검사가 불가")
    rv = c.post("/api/module-f/reopen", json={"key": ok["key"]}).get_json()
    sid = rv["sid"]
    if _idle(c, sid).get("state") != "done":
        pytest.skip("저장본을 여는 데 실패")
    return app, c, sid


def test_이상한_입력에도_500_이_없다(live):
    app, c, sid = live
    rules = []
    for r in app.url_map.iter_rules():
        if "/api/module-f/" not in r.rule or "<" in r.rule:
            continue
        if r.rule in _HEAVY:
            continue
        if "POST" in r.methods:
            rules.append(("POST", r.rule))
        elif "GET" in r.methods:
            rules.append(("GET", r.rule))
    assert len(rules) >= 30, f"검사 대상이 갑자기 줄었다: {len(rules)}"

    crashes = []
    for meth, rule in sorted(rules):
        for w in _WEIRD:
            body = dict(w, sid=sid)
            try:
                if meth == "POST":
                    rv = c.post(rule, json=body)
                else:
                    rv = c.get(rule, query_string={
                        k: (json.dumps(v) if isinstance(v, (list, dict)) else v)
                        for k, v in body.items()})
            except Exception as exc:  # noqa: BLE001 — 던지는 것 자체가 결함
                crashes.append(f"{meth} {rule} {w} → "
                               f"{type(exc).__name__}: {exc}")
                continue
            if rv.status_code >= 500:
                msg = ((rv.get_json() or {}).get("message")
                       or rv.get_data(as_text=True))
                crashes.append(f"{meth} {rule} {w} → {rv.status_code} {msg[:90]}")
            if (c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
                    ).get("state") == "run":
                _idle(c, sid, 600)
    assert not crashes, "500·예외가 났다:\n  " + "\n  ".join(crashes[:10])


def test_셀_수_없는_좌표는_문장으로_거절한다(live):
    """★이 시험이 처음 잡은 결함 — 그 자리를 이름으로 못박는다."""
    _app_, c, sid = live
    for body in ({"x": 1e308, "y": -1e308, "max_d": 0},
                 {"x": float("nan"), "y": 0, "max_d": 1},
                 {"x": float("inf"), "y": 0, "max_d": 1}):
        rv = c.post("/api/module-f/edit/click", json=dict(body, sid=sid))
        assert rv.status_code == 400, f"{body} → {rv.status_code}"
        msg = (rv.get_json() or {}).get("message") or ""
        assert msg and "OverflowError" not in msg, msg


def test_성한_좌표는_종전대로_받는다(live):
    """막는 자를 너무 좁게 잡으면 진짜 클릭이 거절된다."""
    _app_, c, sid = live
    rv = c.post("/api/module-f/edit/click",
                json={"sid": sid, "x": 1000.0, "y": 2000.0, "max_d": 500.0})
    assert rv.status_code == 200, rv.get_data(as_text=True)[:200]
    assert (rv.get_json() or {}).get("ok") is True
