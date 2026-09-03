# -*- coding: utf-8 -*-
"""모듈 F 공통 바탕 — 경로·상수·부팅·입력 검사.

여기 있는 것은 «어느 단계에서나 참인 것» 뿐이다. 단계에 딸린 판단은
`world`(찍기) · `graph`(손질) · `remote30`(범위) 이 각자 갖는다.
"""
from __future__ import annotations

import math
import os
import sys
import threading
from pathlib import Path

from flask import jsonify

# [F-0 · D1] F 가 무는 엔진은 G 다. E(cad_project_editor)는 동결 레퍼런스로
# 은퇴했다 — G1~G18 이 만든 design/(최불리·관경·부속·SDF 방출)은 G 트리에만
# 있고, 두 트리는 services/domain 패키지 이름이 같아 한 프로세스에 둘 다
# 올리면 어느 쪽이 import 되는지가 순서 우연에 걸린다. 그래서 재지정은
# 절반이 아니라 전부다: 경로도, 작업폴더도, sys.path 도 G 하나만.
EDITOR_ROOT = Path(__file__).resolve().parents[2] / "cad_project_editor_g"
# 은퇴한 E 루트 — _boot() 가 «절대 path 에 없어야 한다» 를 검사할 때 쓴다.
RETIRED_E_ROOT = Path(__file__).resolve().parents[2] / "cad_project_editor"
# 찍은스펙·표시캐시·유저손질이 쌓이는 곳. 데스크톱 G 는 cwd 가 편집기 폴더라
# 상대경로 "docs/import" 로 여기를 가리킨다. 웹서버는 cwd 가 프로젝트 루트라
# 같은 상대경로가 엉뚱한 곳을 가리키므로, 부팅 때 절대경로로 고정한다.
# 같은 폴더를 쓰므로 데스크톱 G 에서 찍은 도면이 웹에서 그대로 이어진다.
# E 작업폴더에 있던 저장본은 scripts/_migrate_f_workdir.py 가 1회 복사했다.
IMPORT_WORK_ROOT = EDITOR_ROOT / "docs" / "import"

# 헤드 종류·수직 전개를 설명하는 그림. 모듈 E 의 대화상자가 쓰는 바로 그
# 파일이다(`ui/dialogs/dialog_kfp_convert.py` 의 _DIAGRAMS). E 의 대화상자는
# PySide6 를 끌고 오므로 import 하지 않고 파일 이름만 옮겨 적는다 — 그림
# 자체는 한 벌이라 E 에서 고치면 여기도 같이 바뀐다.
DIAGRAMS = {
    "branch": "_kfp_sample_가지.png",
    "upright": "_kfp_sample_상향식.png",
    "pendant": "_kfp_sample_하향식.png",
    "combo": "_kfp_sample_상하향식.png",
    "valve": "_kfp_sample_알람밸브.png",
}
# 변환 폼 묶음 제목 → 그림. 제목은 dto.py 의 칸 묶음과 1:1 이다.
# 칸 이름이 「① (m)」 뿐이라 그림 없이는 어느 토막인지 읽을 수 없다.
GROUP_DIAGRAM = {
    "메인 → 가지": "branch", "상향식": "upright", "하향식": "pendant",
    "상하향식": "combo", "알람밸브": "valve",
}

# 캔버스로 내려보내는 도형 상한. B1F 실도면이 선분 69,384 개라 이 위로는
# 브라우저가 아니라 JSON 직렬화에서 먼저 막힌다. 조용히 자르지 않고 몇 개를
# 뺐는지 응답에 실어 화면에 그대로 띄운다.
MAX_SEGS = 150_000
MAX_CIRCLES = 40_000
MAX_ARCS = 40_000

SESSION_TTL_SECONDS = 3 * 3600
LOG_TAIL = 40

# 모듈 A 에서 빌려오는 것 — 레이어 이름 사전(찍기 추천)과 Remote 30 개념.
# NFPC 103 이 요구하는 것은 "가장 불리한 헤드 30개" 다. 모듈 E 는 물 닿은
# 헤드를 전부 변환하므로(실측 264개) 법정 계산 대상보다 넓다.
REMOTE_K_DEFAULT = 30

# 자동 이음 — 모듈 A 가 실측으로 배운 것을 모듈 E 의 판정 위에 얹는다.
#
# A 의 교훈: 이음매 여유는 상수가 아니라 «그 도면이 배관선을 끊어놓은 부속 기호
# 획의 길이» 다(실측 — 대명동 70mm · 대구오페라 288mm · MF-304 240mm · B1F
# 141mm). E 의 stage1 은 30mm 고정이라, 그보다 넓게 끊긴 도면은 배관망이 조각난
# 채로 열린다(B1F 실측: 덩이 271개 · 급수원이 닿는 헤드 264/3163).
# 고정값을 키우는 것으로는 못 고친다 — 도면마다 다르기 때문이다.
#
# E 의 교훈: 그렇다고 알아서 이어버리면 안 된다. E 의 이음 규칙은 «사이 추정·
# 연장 금지» 이고, 붙일 수 있는 모양(일직선·T자·직각)만 붙인다.
#
# 그래서 F 는 «A 처럼 재고, E 처럼 붙인다» — 여유는 도면에서 재서 후보만 고르고
# (점선으로 그린다), 실제로 붙일지는 E 의 board.join 이 모양을 보고 정한다.
AUTOJOIN_LADDER_MM = (30.0, 50.0, 75.0, 100.0, 150.0, 200.0,
                      250.0, 300.0, 400.0, 500.0, 650.0, 800.0)
# 후보 상한 — 넘으면 조용히 자르지 않고 몇 개를 뺐는지 화면에 싣는다.
AUTOJOIN_MAX_PAIRS = 4000
# 틈을 «관이 이어지던 자리» 로 보는 각도 여유. 부속 기호가 관을 끊은 자리라면
# 틈은 그 관이 가던 방향으로 나 있다(일직선), 아니면 상대 관 축과 나란하다
# (직각·T자). 이 조건이 없으면 사다리가 반대 힘을 못 받아 늘 상한을 고른다
# (실측: B1F 에서 800mm 까지 올라가 후보 397곳 중 317곳이 E 에게 막혔다).
AUTOJOIN_ANG_TOL_DEG = 15.0
# 정점의 99% 에 처음 닿는 여유를 고른다 — 같은 결과라면 좁은 쪽이 안전하다.
AUTOJOIN_PLATEAU = 0.99

_boot_lock = threading.Lock()
_booted = False

# ★세션 저장소와 무거운 단계 자물쇠는 **여기 없다** — `routes.module_f.jobs`
#   가 갖는다. 종전에는 이 파일에도 같은 이름으로 `_SESSIONS`·`_SESSIONS_LOCK`·
#   `_HEAVY_LOCK` 이 있었는데 **아무도 안 썼다.** 이름이 같아서 더 나빴다:
#   `from .common import _HEAVY_LOCK` 으로 잡으면 진짜 자물쇠(jobs 쪽)와 다른
#   객체를 잡아, 직렬화한 줄 알고 겹쳐 돌게 된다. 죽은 코드가 아니라 «틀리게
#   쓰기 쉬운 함정» 이라 지웠다. 쓸 곳은 `jobs` 에서 가져다 쓴다.


def _fail(msg, code=400):
    """실패 응답 한 벌. 라우트마다 다시 쓰지 않는다."""
    return jsonify({"ok": False, "message": msg}), code


# ─────────────────────────────────────────────────────────── 부팅
def _boot() -> None:
    """편집기 소스를 import 가능하게 하고 쓰기 루트를 절대경로로 못박는다."""
    global _booted
    with _boot_lock:
        if _booted:
            return
        if not (EDITOR_ROOT / "main.py").exists():
            raise RuntimeError(f"모듈 G 엔진 소스를 찾을 수 없습니다: {EDITOR_ROOT}")
        # ★은퇴한 E 루트가 path 에 있으면 조용히 치우지 않고 세운다(D1).
        #   두 트리의 services 이름이 같아, E 가 먼저면 이후의 모든 import 가
        #   엉뚱한 엔진으로 간다 — 그 오염은 여기서 막는 것이 마지막 기회다.
        _e = str(RETIRED_E_ROOT)
        if _e in sys.path:
            raise RuntimeError(
                "은퇴한 모듈 E 루트가 sys.path 에 올라 있습니다 — 모듈 F 는 "
                f"G 엔진 하나만 물어야 합니다: {_e}")
        root = str(EDITOR_ROOT)
        # append 다 — insert(0) 로 앞에 두면 편집기의 services/domain 이 본
        # 프로젝트의 같은 이름 패키지를 가릴 수 있다. 지금은 겹치는 이름이
        # 없지만, 나중에 생겨도 본 서버가 먼저 이기게 둔다.
        if root not in sys.path:
            sys.path.append(root)
        import services as _svc
        # 이미 다른 데서 services 를 E 로 실어 놨어도 여기서 잡힌다.
        _svc_file = str(Path(getattr(_svc, "__file__", "") or "").resolve())
        if not _svc_file.startswith(str(EDITOR_ROOT.resolve())):
            raise RuntimeError(
                f"services 가 G 엔진이 아닌 곳에서 import 되었습니다: {_svc_file}")
        from services.cad_import.pipeline import disp_cache, handoff
        work = str(IMPORT_WORK_ROOT)
        handoff.import_write_root = lambda: work
        # OUT_DIR·_DISP_CACHE_DIR 은 import 때 이미 상대경로로 굳었다.
        # 함수만 갈아끼우면 이 둘은 안 따라오므로 직접 덮는다.
        handoff.OUT_DIR = handoff.pick_out_dir()
        disp_cache._DISP_CACHE_DIR = work
        os.makedirs(handoff.pick_out_dir(), exist_ok=True)
        os.makedirs(handoff.default_edits_dir(), exist_ok=True)
        _booted = True


# ─────────────────────────────────────────────────────────── 도형 직렬화
def _r1(v) -> float:
    return round(float(v), 1)


def _check_key(key: str) -> str:
    r"""도면 키는 «파일 이름» 이지 경로가 아니다 — 경로로 쓰이기 전에 막는다.

    모듈 E 는 키를 경로에 그대로 끼워 넣는다(`user_edits_path` ·
    `_disp_cache_path` · 찍은스펙). 정규화하는 것은 `handoff_path` 하나뿐이다.
    데스크톱 E 는 제 목록에서 고른 키만 넘기니 문제가 안 됐지만, 웹은 아무
    문자열이나 들어온다 — 실측으로 `..\..\..` 가 `docs/import` 밖으로 나갔다.
    E 소스는 안 고치는 것이 모듈 F 의 계약이므로 **들어오는 문 앞에서** 막는다.
    """
    k = str(key or "").strip()
    if not k:
        raise ValueError("도면 키가 비었습니다.")
    if len(k) > 160:
        raise ValueError("도면 키가 너무 깁니다.")
    if k != os.path.basename(k) or k in (".", ".."):
        raise ValueError(f"도면 키에 경로를 쓸 수 없습니다: {key!r}")
    if ".." in k or any(c in k for c in r'\/:*?"<>|'):
        raise ValueError(f"도면 키에 쓸 수 없는 문자가 있습니다: {key!r}")
    if any(ord(c) < 32 for c in k):
        raise ValueError("도면 키에 제어문자가 있습니다.")
    return k


# 도면 좌표로 받아들일 범위(mm). 1e9mm = 1,000km — 어떤 도면도 이 밖에 없다.
# ★상한이 필요한 이유는 «큰 도면» 때문이 아니라 **제곱** 때문이다. 엔진은
#   점–선분 거리를 제곱으로 잰다(`user_net._pt_seg_d2`). float 은 1.3e154 를
#   넘겨 제곱하면 OverflowError 를 던지고, 그것이 그대로 HTTP 500 이 된다
#   (실측: `edit/click` 에 x=1e308 → «서버 오류: OverflowError»).
#   엔진은 안 고친다 — 데스크톱 G 는 그런 좌표를 만들 수 없고, 웹은 아무
#   문자열이나 들어온다. 그래서 **들어오는 문 앞에서** 막는다(`_check_key`
#   와 같은 이유·같은 자리).
COORD_LIMIT_MM = 1e9


def _check_num(v, what: str, *, lo: float, hi: float,
               default=None) -> float:
    """숫자 하나 — 셀 수 있고 범위 안인가. NaN·무한대는 «숫자가 아니다».

    `default` 를 주면 값이 없을 때 그것을 쓴다(안 주면 «없음» 도 실패다).
    """
    if v is None or v == "":
        if default is None:
            raise ValueError(f"{what} 값이 필요합니다.")
        return float(default)
    try:
        f = float(v)
    except (TypeError, ValueError):
        raise ValueError(f"{what} 이(가) 숫자가 아닙니다: {v!r}") from None
    # ★`math.isfinite` 로 NaN·±inf 를 함께 거른다. NaN 은 어떤 비교와도
    #   거짓이라 «가장 가까운 것 고르기» 를 조용히 망가뜨린다 — 튀지도 않고
    #   틀린 답을 낸다. 튀는 쪽(무한대)보다 이쪽이 더 나쁘다.
    if not math.isfinite(f):
        raise ValueError(f"{what} 이(가) 셀 수 있는 수가 아닙니다: {v!r}")
    if not (lo <= f <= hi):
        raise ValueError(f"{what} 이(가) 범위를 벗어났습니다: {f:g} "
                         f"(허용 {lo:g} ~ {hi:g})")
    return f


def _check_xy(body, *, kx: str = "x", ky: str = "y", what: str = "좌표"):
    """도면 좌표 한 쌍. 네 곳(찍기·손질·원클릭·자동·계통도)이 같은 자를 쓴다."""
    return (_check_num(body.get(kx), f"{what} x",
                       lo=-COORD_LIMIT_MM, hi=COORD_LIMIT_MM),
            _check_num(body.get(ky), f"{what} y",
                       lo=-COORD_LIMIT_MM, hi=COORD_LIMIT_MM))


def _layer_category(name: str) -> str:
    """모듈 A 의 레이어 이름 사전으로 분류한다(PIPE/HEAD/ALARM/ARCH/…).

    **추천이지 결정이 아니다.** 사전은 이름만 보므로 실도면에서 절반 넘게
    OTHER 로 떨어진다(B1F 51묶음 중 35). 그래도 51개를 맨눈으로 훑는 것보다는
    출발점이 되고, 확정은 모듈 E 의 찍기(사람 클릭)가 그대로 맡는다.
    """
    try:
        from remote30_prototype import _categorize_layer
    except Exception:  # noqa: BLE001 — 모듈 A 를 못 불러도 찍기는 돌아야 한다
        return "OTHER"
    try:
        return _categorize_layer(str(name))
    except Exception:  # noqa: BLE001
        return "OTHER"
