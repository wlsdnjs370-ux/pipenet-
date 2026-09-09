# -*- coding: utf-8 -*-
"""[경로 늦은 바인딩] 부팅 뒤에 정한 쓰기 루트를 **먼저 import 된 모듈도** 따라오는가.

지시서 `ModuleF_경로상수_함수화_지시서.md` §3-1.

■ 무엇이 문제였나

  경로가 모듈 레벨 상수로 **import 시점에 굳었다**::

      pipeline/handoff.py:42   OUT_DIR = pick_out_dir()
      pick/io.py:13            NEW_DIR = handoff.OUT_DIR      ← 값을 복사한다
      pick/__init__.py:10      NEW_DIR 재수출                  ← 복사가 한 번 더
      pipeline/disp_cache.py:15 _DISP_CACHE_DIR = import_write_root()

  서버는 부팅 때 `_boot()` 에서 이 값들을 덮어썼지만, **이미 복사해 간 모듈은
  따라오지 않는다.** 한 프로세스 안에서 모듈마다 다른 경로를 붙들고, 오류는
  안 난다 — 파일이 조용히 엉뚱한 폴더에 생기거나 안 보인다.

  ★그래서 이 시험은 «부팅 전에 import 한 모듈» 을 일부러 만든다. 조치 전
    코드로 돌리면 실패해야 한다(그 실패가 이 시험의 존재 이유다).

■ 여기서 지키는 것

  ⑴ `set_write_root()` 하나가 주입점이다 — 함수를 갈아끼우지 않는다.
  ⑵ 먼저 import 한 모듈도 **부를 때** 읽으므로 새 루트를 따라온다.
  ⑶ 주입이 없으면 값이 예전과 **똑같다**(함수화만 했다).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _mods():
    """★부팅 «전» 에 import 한 시늉 — 실제로 그 순서로 들여온다."""
    import services.cad_import.pick.io as io
    import services.cad_import.pipeline.disp_cache as dc
    import services.cad_import.pipeline.handoff as handoff
    return io, dc, handoff


@pytest.fixture(autouse=True)
def _keep_write_root():
    """★쓰기 루트를 건드렸으면 **그때 가리키던 폴더로** 되돌린다.

    처음에 `set_write_root(None)` 으로 되돌렸다가 옆 시험 넷을 깼다. `_boot()`
    는 `_booted` 로 **한 번만** 도므로 None 으로 비워 두면 다시는 안 채워지고,
    그 뒤의 모든 모듈 F 시험이 cwd 상대경로 "docs/import" 를 본다 — 이 지시서가
    없애려던 바로 그 사고를, 그것을 지키는 시험이 저지른 셈이었다.

    실측(전체 스위트): 4 failed — 「만든 저장본이 목록에 없다」(찍은스펙은
    상대경로 폴더에 쓰고 `pick_store_dir()` 는 실작업 폴더에서 읽는다) 외.
    두 파일만 따로 돌리면 전부 통과해 «순서 의존» 으로만 보였다.
    """
    _io, _dc, handoff = _mods()
    saved = handoff.import_write_root()
    yield
    handoff.set_write_root(saved)


def _restore(handoff):
    """이 시험 안에서 «주입 없음» 상태를 보고 싶을 때만 — 뒷정리는 픽스처가 한다."""
    handoff.set_write_root(None)


def test_먼저_import_한_모듈도_새_루트를_따라온다(tmp_path):
    """★이 저장소가 값을 치른 자리 — 조치 전에는 여기서 빨간 줄이 났다."""
    io, dc, handoff = _mods()
    try:
        handoff.set_write_root(str(tmp_path))
        assert str(tmp_path) in io.new_dir(), io.new_dir()
        assert str(tmp_path) in dc._disp_cache_dir(), dc._disp_cache_dir()
        assert str(tmp_path) in handoff.pick_out_dir()
        assert str(tmp_path) in handoff.default_edits_dir()
    finally:
        _restore(handoff)


def test_표준샘플도_따라온다(tmp_path):
    """★`STD_DIR` 은 cwd 상대경로라 부팅 목록에서 아예 빠져 있었다.

    데스크톱 G 는 cwd 가 편집기 폴더라 "docs/import" 가 맞지만, 웹서버는 cwd 가
    프로젝트 루트라 같은 상대경로가 엉뚱한 곳을 가리킨다.
    """
    io, _dc, handoff = _mods()
    try:
        handoff.set_write_root(str(tmp_path))
        assert str(tmp_path) in io.std_dir(), io.std_dir()
        assert io.std_dir().endswith("0단계_표준샘플")
    finally:
        _restore(handoff)


def test_주입이_없으면_값이_예전과_같다():
    """함수화만 한다 — 같은 실행 조건에서 **같은 경로**가 나와야 한다."""
    import os
    io, dc, handoff = _mods()
    _restore(handoff)
    assert handoff.import_write_root() == os.path.join("docs", "import")
    assert handoff.pick_out_dir() == os.path.join("docs", "import",
                                                  "0단계_새찍기")
    assert io.new_dir() == handoff.pick_out_dir()
    assert dc._disp_cache_dir() == handoff.import_write_root()
    assert io.std_dir() == os.path.join("docs", "import", "0단계_표준샘플")


def test_주입점은_하나다():
    """함수를 람다로 갈아끼우면 누가 언제 바꿨는지 추적이 안 된다."""
    src = (_ROOT / "routes" / "module_f" / "common.py").read_text(
        encoding="utf-8")
    assert "handoff.import_write_root = lambda" not in src
    assert "handoff.OUT_DIR =" not in src
    assert "_DISP_CACHE_DIR =" not in src
    assert "handoff.set_write_root(" in src
    # ★E 루트 차단·`services` 검증은 **지우지 않는다** — 두 트리의 최상위
    #   패키지 이름이 둘 다 `services` 라 이것이 유일한 방어다(지시서 §5).
    assert "cad_project_editor" in src


def test_굳은_경로_상수가_안_남았다():
    """지시서 §3-2 — 남아도 되는 것은 `__file__` 기준 절대경로 둘뿐이다."""
    g = _ROOT / "cad_project_editor_g" / "services" / "cad_import"
    bad = []
    for py in g.rglob("*.py"):
        for no, line in enumerate(py.read_text(encoding="utf-8").splitlines(), 1):
            if line.startswith((" ", "\t", "#")) or "=" not in line:
                continue
            head, _, rhs = line.partition("=")
            if not head.strip().isidentifier():
                continue
            if any(w in rhs for w in ("import_write_root(", "pick_out_dir(",
                                      "default_edits_dir(", "OUT_DIR",
                                      '"docs"', "'docs'")):
                bad.append(f"{py.relative_to(g)}:{no} {line.strip()}")
    # stage1.py 의 둘(`_APP_ROOT` · `DWG_DIR`)은 `__file__` 기준 **절대경로**라
    # cwd·부팅 순서에 안 걸린다 — 그대로 둔다(지시서 §3-2).
    left = [b for b in bad if "stage1.py" not in b]
    assert left == [], left
