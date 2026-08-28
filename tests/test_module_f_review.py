# -*- coding: utf-8 -*-
"""전체 로직 검토에서 나온 것들을 회귀로 못박는다.

검토가 찾아낸 다섯 가지:
  ① 손질 라우트가 워커 스레드와 같은 board 를 고쳤다 (가드 없음)
  ② 영역 개수에 상한이 없었다 (헤드 × 영역으로 늘어난다)
  ③ 세션 회수가 «새 도면 열 때» 에만 돌았고 수 상한도 없었다
  ④ 자동 경로에서 산출 저장 단추가 영영 잠겼다 (화면)
  ⑤ 자동 경로의 키가 `_check_key` 를 안 거쳤다

④는 화면이라 브라우저 검증기(_verify_module_f_ui.py)가 본다.
"""
from __future__ import annotations

import os
import sys
import time

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from routes.module_f import jobs  # noqa: E402
from routes.module_f.api_auto import HEAD_PREVIEW_CAP, MAX_ZONES  # noqa: E402
from routes.module_f.common import _check_key  # noqa: E402


@pytest.fixture(autouse=True)
def _clean_sessions():
    with jobs._SESSIONS_LOCK:
        jobs._SESSIONS.clear()
    jobs._last_sweep = 0.0
    yield
    with jobs._SESSIONS_LOCK:
        jobs._SESSIONS.clear()


@pytest.fixture()
def app_ctx():
    """`_fail` 은 jsonify 를 쓴다 — 앱 컨텍스트가 있어야 만들어진다."""
    from flask import Flask
    app = Flask(__name__)
    with app.app_context():
        yield


# ─────────────────────────────────────────── ① 손질 가드
def test_손질_헬퍼가_작업중을_거절한다(app_ctx):
    """헬퍼는 `route_session` 규약대로 «(sess, 사유)» 를 돌려준다.

    사유는 «(문장, 코드)» 다 — 자리를 데코레이터로 옮기면서도 옮기기 전의
    상태 코드를 그대로 지키기 위한 모양이다.
    """
    from routes.module_f.api_edit import _edit_session
    sess = jobs._new_session()
    sess["edit"] = object()
    _s, why = _edit_session({"sid": sess["id"]})
    assert why is None and sess["edit"] is not None

    sess["job"] = {"state": "run", "phase": "자동 이음", "started": 0.0,
                   "ended": None, "error": None, "result": None}
    _s, why = _edit_session({"sid": sess["id"]})
    assert why is not None, "작업이 도는데 board 를 내줬다"
    assert why[1] == 409


def test_손질_헬퍼가_세션없음도_가른다(app_ctx):
    from routes.module_f.api_edit import _edit_session
    sess = jobs._new_session()          # edit 없음
    _s, why = _edit_session({"sid": sess["id"]})
    assert why is not None and why[1] == 400


def test_모든_board_변경_라우트가_헬퍼를_탄다():
    """새 라우트가 늘어도 가드를 빠뜨리지 않게 — 소스로 확인한다.

    가드는 이제 함수 «앞» 의 `@route_session(_edit_session, …)` 에 있다.
    그래서 def 줄 앞뒤를 함께 본다 — 몸통만 보면 «있는데 없다» 고 나온다.
    """
    import inspect

    from routes.module_f import api_edit
    src = inspect.getsource(api_edit)
    # board 를 고치는 라우트들
    for fn in ("module_f_edit_click", "module_f_edit_kind",
               "module_f_edit_undo", "module_f_edit_mode",
               "module_f_edit_flow", "module_f_edit_worst",
               "module_f_edit_autojoin_scan",
               "module_f_edit_autojoin_apply"):
        i = src.index(f"def {fn}(")
        around = src[max(0, i - 200):i + 700]
        assert "_edit_session" in around, f"{fn} 이 가드를 안 탄다"
        assert "route_session(" in around, f"{fn} 이 가드를 안 탄다"


def test_설정을_바꾸는_라우트도_가드를_탄다():
    """2차 검토에서 나온 것 — design/emit · merge/mode 가 무가드였다.

    emit 은 «표 확정» 잡이 도는 동안 옛 표로 파일을 쓰고, merge/mode 는 결합
    잡이 돌면서 읽는 값(급수방식·낙차·펌프)을 바꾼다.
    """
    import inspect

    from routes.module_f import api_design, api_merge
    for mod, fn in ((api_design, "module_f_design_emit"),
                    (api_merge, "module_f_merge_mode")):
        src = inspect.getsource(mod)
        i = src.index(f"def {fn}(")
        assert "_job_running" in src[i:i + 900], f"{fn} 이 가드를 안 탄다"


# ─────────────────────────────────────────── ② 영역 상한
def test_영역_상한이_있다():
    assert 1 <= MAX_ZONES <= 1000
    assert HEAD_PREVIEW_CAP > 0


# ─────────────────────────────────────────── ③ 세션 회수·상한
def test_세션이_만료되면_회수된다():
    s = jobs._new_session()
    s["touched"] = time.time() - (jobs.SESSION_TTL_SECONDS + 10)
    jobs._sweep(force=True)
    assert s["id"] not in jobs._SESSIONS


def test_상한을_넘으면_오래된_것부터_버린다():
    # ★touched 는 «만료 안 된» 범위여야 한다 — 과거 절대값을 넣으면 TTL 에
    #   먼저 걸려 전부 사라지고, 상한 규칙을 시험하지 못한다(한 번 헛짚었다).
    now = time.time()
    made = []
    for i in range(jobs.MAX_SESSIONS + 5):
        s = jobs._new_session()
        s["touched"] = now - (jobs.MAX_SESSIONS + 5 - i)   # 앞이 더 오래됐다
        made.append(s)
    jobs._sweep(force=True)
    assert len(jobs._SESSIONS) <= jobs.MAX_SESSIONS
    assert made[-1]["id"] in jobs._SESSIONS, "가장 최근 것을 버렸다"
    assert made[0]["id"] not in jobs._SESSIONS, "가장 오래된 것이 남았다"


def test_도는_작업은_상한에도_안_버린다():
    """계산 중인 세션을 회수하면 그 작업이 통째로 사라진다."""
    now = time.time()
    busy = jobs._new_session()
    busy["touched"] = now - 3600         # 가장 오래됐지만 만료 전
    busy["job"] = {"state": "run", "phase": "자동 추출", "started": 0.0,
                   "ended": None, "error": None, "result": None}
    for i in range(jobs.MAX_SESSIONS + 5):
        jobs._new_session()["touched"] = now - i
    jobs._sweep(force=True)
    assert busy["id"] in jobs._SESSIONS, "도는 작업을 버렸다"


def test_sweep_은_스로틀된다():
    """_sess 마다 전량을 훑으면 세션이 많을 때 요청마다 비용이 붙는다."""
    s = jobs._new_session()
    s["touched"] = time.time() - (jobs.SESSION_TTL_SECONDS + 10)
    jobs._last_sweep = time.time()       # 방금 훑은 것으로 둔다
    jobs._sweep()                        # force 아님 → 건너뛴다
    assert s["id"] in jobs._SESSIONS
    jobs._sweep(force=True)
    assert s["id"] not in jobs._SESSIONS


# ─────────────────────────────────────────── ⑤ 키 검사
def test_두_길_다_같은_키_자를_쓴다():
    """읽기가 공통이 되면서 키 검사도 한 자리로 모였다 — 그 자리를 확인한다.

    키는 뒤에서 산출 파일 이름이 된다. 자동·수동 어느 길로 가든 `_open_job`
    (찍기판)이 도면을 열므로 거기 하나면 둘 다 덮인다.
    """
    import inspect

    from routes.module_f import api_open, api_slot
    assert "_check_key(" in inspect.getsource(api_open._open_job), \
        "공통 열기가 키 검사를 건너뛴다"
    # 자동은 보태기만 한다 — 키를 새로 정하지 않는다(덮으면 두 길이 갈린다).
    aug = inspect.getsource(api_slot._auto_augment_job)
    assert 'sess["key"] =' not in aug, "자동이 키를 덮어쓴다"


def test_키_자가_경로를_막는다():
    for bad in ("..", "a/b", "a\\b", "a:b", ""):
        with pytest.raises(ValueError):
            _check_key(bad)
    assert _check_key("B1F 평면도") == "B1F 평면도"
