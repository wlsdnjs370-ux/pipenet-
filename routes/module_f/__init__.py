# -*- coding: utf-8 -*-
"""모듈 F — CAD 임포트 웹 워크벤치. 모듈 E 의 소스를 브라우저에서 돌린다.

모듈 E(`cad_project_editor/`)는 PySide6 데스크톱 앱이라 `/module-e-cad-editor`
는 서버 PC 화면에 Qt 창을 띄울 뿐 브라우저 안에서는 아무것도 보이지 않는다.
그런데 E 는 `ui/` 밖이 전부 순수 Python 이다 — 찍기·손질·변환 세 파사드가
모두 "화면 없음" 을 계약으로 달고 있고, 판정은 전부 그 아래 board 에 있다.

그래서 이 패키지는 **E 의 소스를 한 줄도 고치지 않고** 그 파사드만 HTTP 로
열고, 캔버스·버튼은 브라우저에서 새로 그린다. Qt 는 import 하지 않는다
(`cad_project_editor/_smoke_headless.py` 로 3단 전부 Qt 없이 도는 것을 확인).

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 다른 도메인 라우트와 같은
패턴이라 엔드포인트명에 접두사가 붙지 않는다.

구성 — 아래로 갈수록 위를 쓴다(순환 없음)::

    common      경로·상수·부팅·입력 검사·실패 응답
    slots       도면 슬롯 세 칸의 상태 기계
    jobs        세션 · 무거운 작업(한 번에 하나) · `route_session` 라우트 앞머리
    world       찍기 도면 직렬화·저장 목록
    graph       덩이 통계·자동 이음(A 의 실측 + E 의 판정)
    remote30    최불리 K·도면 장 나누기·범위 제한·PIPENET
    recon       [F-8a] 모듈 A 인식 한 번 — 정찰
    adopt       [F-8b] 정찰 결과를 클릭 경로로 채택
    auto        [A 방식] 평면도 자동 추출 엔진 호출부
    merge       세 도면 결합
    subdrawing  계통도·기계실 추출
    emit        산출물 쓰기
    views       화면이 받는 상태 한 장(`keep` 규약)

    api_open · api_slot · api_sub · api_auto · api_pick ·
    api_edit · api_design · api_convert · api_merge     라우트 아홉 묶음

라우트 앞머리(요청 꺼내기 · 세션 찾기 · 실패 응답)는 `jobs.route_session`
하나로 모았다 — 예전에는 라우트 43곳이 저마다 네댓 줄을 손으로 적고 있었다.

`_` 로 시작하는 이름을 아래에서 다시 내보낸다. 진단 스크립트(`data/_probe_*.py`)
가 `import routes.module_f as mf` 로 붙잡아 쓰던 것들이라, 가르는 것 때문에
그 도구들이 깨지면 안 된다.
"""
from __future__ import annotations

from routes.module_f import (api_auto, api_convert, api_design, api_edit,
                             api_merge, api_open, api_pick, api_slot, api_sub)
from routes.module_f.common import (  # noqa: F401  (진단 스크립트가 쓴다)
    AUTOJOIN_ANG_TOL_DEG, AUTOJOIN_LADDER_MM, AUTOJOIN_MAX_PAIRS,
    AUTOJOIN_PLATEAU, DIAGRAMS, EDITOR_ROOT, GROUP_DIAGRAM, IMPORT_WORK_ROOT,
    LOG_TAIL, MAX_ARCS, MAX_CIRCLES, MAX_SEGS, REMOTE_K_DEFAULT,
    SESSION_TTL_SECONDS, _boot, _check_key, _fail, _layer_category, _r1)
from routes.module_f.graph import (  # noqa: F401
    _autojoin_apply, _autojoin_scan, _body_index, _body_stat)
from routes.module_f.jobs import (  # noqa: F401
    _HEAVY_LOCK, _SESSIONS, _job_running, _job_view, _new_session, _run_job,
    _sess, _sweep, route_session)
from routes.module_f.remote30 import (  # noqa: F401
    _emit_pipenet, _restrict_to_worst, _sheet_frames, _worst_k_heads,
    _worst_view)
from routes.module_f.slots import (  # noqa: F401
    SESSION_KEYS, SLOT_KINDS, SLOT_LABELS, _check_slot_kind, _slot_active,
    _slot_blank, _slot_capture, _slot_init, _slot_progress, _slot_restore,
    _slot_state, _slot_switch)
from routes.module_f.views import (  # noqa: F401
    _autojoin_view, _edit_state, _net_rev, _pick_state)
from routes.module_f.world import (  # noqa: F401
    _pts_bounds, _saved_keys, _world_payload)


def register(app, *, _save_upload, UPLOAD_DIR):
    """라우트 아홉 묶음(60개)을 한 앱에 붙인다.

    묶음마다 필요한 것만 받는다 — 업로드는 열기·슬롯에서만, 산출물 폴더는
    변환·설계·결합에서만 쓴다. 예전에는 한 `register()` 가 둘 다 들고
    스물여덟 라우트를 품었다.
    """
    api_open.register(app, _save_upload=_save_upload)
    api_slot.register(app, _save_upload=_save_upload)
    api_sub.register(app)
    api_auto.register(app)
    api_pick.register(app)
    api_edit.register(app)
    api_design.register(app, UPLOAD_DIR=UPLOAD_DIR)
    api_convert.register(app, UPLOAD_DIR=UPLOAD_DIR)
    api_merge.register(app, UPLOAD_DIR=UPLOAD_DIR)
