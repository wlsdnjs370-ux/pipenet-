# -*- coding: utf-8 -*-
"""세션 수명 — «도는 작업» 은 회수되지 않는다.

`_sweep` 은 두 갈래로 세션을 걷는다: 만료(TTL)와 수 상한(MAX_SESSIONS).
수 상한 쪽은 「도는 작업은 건드리지 않는다」고 명시하는데 **만료 쪽은
안 봐줬다** — 같은 함수 안에서 규칙이 갈려 있었다.

걷히면 워커는 계속 돌지만 결과를 받을 세션이 없다. 사람은 일이 끝난 뒤에야
「작업이 만료되었습니다. 도면을 다시 여세요」를 본다 — 무엇이 잘못됐는지도
모른 채 처음부터 다시 한다.

방아쇠는 진행 스트림이었다: 폴링은 요청마다 `_sess` 가 touched 를 갱신하는데
SSE 는 세션을 한 번 찾고 루프를 돌아, **지켜보는 동안 세션만 가만히 늙었다.**
"""
from __future__ import annotations

import importlib
import os
import sys
import time

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def test_도는_작업은_만료로_걷히지_않는다():
    from routes.module_f import jobs

    sess = jobs._new_session(slot="plan")
    sid = sess["id"]
    # 아주 오래 안 만진 것처럼 — 그러나 잡은 돌고 있다.
    sess["touched"] = time.time() - jobs.SESSION_TTL_SECONDS - 60
    sess["job"] = {"state": "run", "phase": "긴 작업", "started": time.time(),
                   "ended": None, "error": None, "result": None}

    jobs._sweep(force=True)
    assert jobs._sess(sid) is sess, "도는 작업의 세션이 만료로 걷혔다"


def test_끝난_작업은_만료로_걷힌다():
    """봐주는 자를 너무 넓게 잡지 않았나 — 안 도는 것은 종전대로 걷어야 한다."""
    import pytest
    from routes.module_f import jobs

    sess = jobs._new_session(slot="plan")
    sid = sess["id"]
    sess["touched"] = time.time() - jobs.SESSION_TTL_SECONDS - 60
    sess["job"] = {"state": "done", "phase": "끝난 작업", "started": 0.0,
                   "ended": 0.0, "error": None, "result": None}

    jobs._sweep(force=True)
    with pytest.raises(ValueError):
        jobs._sess(sid)


def test_잡이_없는_낡은_세션도_걷힌다():
    import pytest
    from routes.module_f import jobs

    sess = jobs._new_session(slot="plan")
    sid = sess["id"]
    sess["touched"] = time.time() - jobs.SESSION_TTL_SECONDS - 60
    jobs._sweep(force=True)
    with pytest.raises(ValueError):
        jobs._sess(sid)


def test_진행_스트림은_보는_동안_세션을_살려_둔다():
    """SSE 가 touched 를 갱신하지 않으면 지켜보는 세션만 늙는다."""
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    c = srv.app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True

    from routes.module_f import jobs
    sess = jobs._new_session(slot="plan")
    sid = sess["id"]
    old = time.time() - 9999
    sess["touched"] = old
    # 잡 없음 → 스트림은 idle 을 몇 박자 보다 스스로 끝난다(오래 안 걸린다).
    rv = c.get(f"/api/module-f/job/stream?sid={sid}")
    rv.get_data()          # 스트림을 끝까지 소비한다
    assert sess["touched"] > old, "스트림이 도는 동안 세션이 늙었다"
