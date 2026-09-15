# -*- coding: utf-8 -*-
"""[G14] `<User-lib file="…">` 의 **표기** — 한글 도면 이름으로도 라이브러리가 열리나.

■ 사용자 신고 (2026-09-14)

  「대명동 기준, 평면도 파트에서 수리계산 입력 변환으로 다운 받은 .sdf 파일에
   라이브러리(.slf)가 반영이 안된 것 같은데 조치좀 해주겠어?」

■ 원인

  PIPENET 은 이 속성을 **UTF-8 문자열로 안 읽는다.** 자기가 쓴 파일을 보면
  답이 나온다 — 한글 경로가 「CP949 바이트를 한 글자씩 latin-1 로 넓힌 뒤
  UTF-8 로 직렬화한」 형태다. 읽을 때는 그 반대로 latin-1 로 **좁혀** CP949
  바이트를 얻어 파일을 연다.

  우리는 진짜 UTF-8 로 썼다. 그러면 좁히기가 실패해(한글 코드포인트 > 255)
  라이브러리를 아예 못 연다 → `Type` 열은 차는데 **모든 `Diameter` 가 'Unset'**.

■ 이 결함이 오래 안 보인 이유

  저장소의 시험과 실측이 **전부 ASCII 이름**이었다. ASCII 는 두 표기가 같아
  아무 차이도 안 난다. 한글 도면 이름으로 PIPENET 을 돌려 본 적이 없었다.
  그래서 여기서는 **한글 이름으로** 못박는다.
"""
from __future__ import annotations

import os
import re
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core"),
           os.path.join(_ROOT, "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.sdf_post import (  # noqa: E402
    pipenet_lib_name, read_pipenet_lib_name)

_KO = "1. 입력도면 대명동 단위세대 평면도_수리계산입력.slf"


# ═══════════════════════════════════════════ ① 표기 규약
def test_ASCII_이름은_한_글자도_안_바뀐다():
    """ASCII 는 두 표기가 같다 — 그래서 옛 시험들이 이 결함을 못 봤다.
    바꾸지 않는다는 것 자체가 계약이다(기존 산출물 바이트 불변)."""
    for n in ("alpha.slf", "2. Pipenet_hand.slf", "tree25.slf",
              "combined_b477e6a2292b.slf", "module_f_merged_iso.slf"):
        assert pipenet_lib_name(n) == n


def test_한글_이름은_PIPENET_이_여는_표기가_된다():
    """쓴 값을 UTF-8 로 직렬화 → PIPENET 처럼 latin-1 로 좁힘 → CP949 해독."""
    w = pipenet_lib_name(_KO)
    assert w != _KO, "한글인데 표기가 그대로다 — PIPENET 이 못 읽는다"
    # PIPENET 이 하는 그대로
    got = w.encode("latin-1").decode("cp949")
    assert got == _KO, f"PIPENET 이 열 이름이 다르다: {got!r}"


def test_표기는_왕복한다():
    for n in (_KO, "alpha.slf", "B1F 현장조사 소화설비 평면도_수리계산입력.slf"):
        assert read_pipenet_lib_name(pipenet_lib_name(n)) == n


def test_CP949_밖의_글자는_지어내지_않는다():
    """옮길 수 없으면 원래 이름을 그대로 둔다 — 깨진 표기를 만드느니 낫다."""
    n = "도면🙂.slf"
    assert pipenet_lib_name(n) == n


# ═══════════════════════════════════════════ ② PIPENET 이 쓴 진짜 파일과 대조
_SAMPLES = [
    os.path.join(_ROOT, "assets",
                 "3-1형_자연낙차_LSP_4F_OA_지하층포함_120m~200m미만_6.6K로 감압_알람밸브.sdf"),
]


@pytest.mark.parametrize("path", _SAMPLES)
def test_PIPENET_이_쓴_파일이_그_표기를_쓴다(path):
    """★이 시험이 규약의 **근거**다 — 우리 추측이 아니라 그쪽 산출물이 증인이다.

    PIPENET 이 직접 쓴 SDF 의 `User-lib` 바이트를 우리 해독기로 풀면 사람이
    읽는 한글 경로가 나와야 한다. 안 나오면 우리가 규약을 잘못 읽은 것이다.
    """
    if not os.path.isfile(path):
        pytest.skip(f"레퍼런스 없음: {os.path.basename(path)}")
    raw = re.search(rb'<User-lib file="([^"]*)"',
                    open(path, "rb").read()).group(1)
    human = read_pipenet_lib_name(raw.decode("utf-8"))
    assert human.endswith(".slf"), human
    # 한글이 복원돼야 한다 — 이 레퍼런스는 한글 경로를 쓴다
    assert any("\uac00" <= ch <= "\ud7a3" for ch in human), \
        f"한글이 안 풀렸다: {human!r}"


# ═══════════════════════════════════════════ ③ 산출물에서 끝까지
def test_sanitize_가_한글_이름을_그_표기로_쓴다(tmp_path):
    """`sanitize_template` 이 실제로 그 표기를 박는지 — 파일 바이트로 본다."""
    from services.cad_import.design.sdf_post import sanitize_template

    sdf = tmp_path / "그림.sdf"
    sdf.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<Project version="1.8  (0)"><Network-spray>'
        '<Libraries><User-lib file="C:\\남의PC\\옛것.slf"/></Libraries>'
        "</Network-spray></Project>\n", encoding="utf-8")
    got = sanitize_template(sdf, _KO)
    assert got["user_lib"] == 1

    raw = re.search(rb'<User-lib file="([^"]*)"',
                    sdf.read_bytes()).group(1)
    # ★진짜 UTF-8 로 쓰면 이 좁히기가 터진다 — 그것이 바로 깨졌던 자리다
    narrowed = raw.decode("utf-8").encode("latin-1")
    assert narrowed.decode("cp949") == _KO
