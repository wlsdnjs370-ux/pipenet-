# -*- coding: utf-8 -*-
"""헤드 종류의 원천은 하나다 — head_kinds(엔진)가 권위, disk_kinds(화면)는 파생.

■ 무엇이 갈라져 있었나 (2026-09-03 실측)

  분류가 레코드를 못 만든 헤드에 사람이 종류를 찍으면:
    · 화면(disk_kinds)  = 하향식   ← set_head_kind 의 «덧댐 블록» 이 칠했다
    · 엔진(head_kinds)  = 미지정   ← apply_kind_overrides 는 «갱신만» 한다
    · 재열기 후         = 미지정   ← io.py 가 apply(갱신만) → require(삽입) 순서
  즉 사람이 확정한 종류를 화면만 믿고, 수리계산·변환 게이트는 딴것을 읽었다.
  같은 순서 뒤집힘이 다섯 호출처(board·io·engine·preflight·planar·flow)에
  똑같이 서 있었다.

■ 무엇을 못박나

  «require(레코드 세우기) → apply(사람 결정 덮기)» 순서를 resolve_head_kinds
  한 함수에 넣고, 종류를 읽는 자리는 전부 그 함수를 거친다. disk_kinds 는
  언제나 disk_kind_list(disks, head_kinds) 와 같아야 한다 — 손으로 칠한
  화면은 거짓말이다.
"""
from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)

from services.cad_import.edit.board import EditBoard          # noqa: E402
from services.cad_import.edit.io import load_edits, write_edits  # noqa: E402
from services.cad_import.convert.preflight import (           # noqa: E402
    preflight_kfp_convert)
from services.cad_import.kinds import (                       # noqa: E402
    disk_kind_list, require_head_kinds, resolve_head_kinds)


def _bare_board(key="시험_종류SSOT"):
    """분류가 레코드를 하나도 못 만든 판 — head_kinds 가 빈 채로 시작한다."""
    pts = [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)]
    edges = [(0, 1), (1, 2)]
    disks = [(0.0, 0.0, 150.0), (6000.0, 0.0, 150.0)]
    return EditBoard(key, pts, edges, disks, head_kinds=[])


def _engine_kind_at(board, xy):
    """변환·수리계산이 읽는 그대로 — require 를 거친 head_kinds."""
    eng = require_head_kinds(board.disks, board.head_kinds)
    return {tuple(r["c"]): r["kind"] for r in eng}.get(xy)


def test_레코드_없던_헤드에_찍은_종류를_엔진도_읽는다():
    b = _bare_board()
    assert b.set_head_kind(b.disks[0], "하향식") == "하향식"
    assert b.disk_kinds[0] == "하향식", "화면이 안 바뀌었다"
    assert _engine_kind_at(b, (0.0, 0.0)) == "하향식", \
        "화면은 하향식인데 엔진은 딴것을 읽는다 — 바로 그 갈림이다"
    # 안 찍은 헤드는 여전히 미지정이어야 한다 — 상향식 가정 금지.
    assert _engine_kind_at(b, (6000.0, 0.0)) == "미지정"


def test_화면은_파생캐시다_직접_칠하지_않는다():
    """어느 시점에서든 disk_kinds == disk_kind_list(disks, head_kinds).

    종전 덧댐 블록은 파생이 안 먹힌 자리를 화면만 칠해 이 불변량을 깼다 —
    그 «편리» 가 갈림을 숨기는 장본인이라 지웠다. undo 뒤에도 성해야 한다.
    """
    b = _bare_board()
    def derived_ok():
        return b.disk_kinds == disk_kind_list(b.disks, b.head_kinds)
    assert derived_ok()
    b.set_head_kind(b.disks[0], "상하향식")
    assert derived_ok(), "찍은 직후 화면이 파생과 다르다 — 덧댐이 되살아났다"
    b.undo()
    assert derived_ok()
    assert b.disk_kinds[0] == "미지정", "undo 가 종류를 안 되돌렸다"


def test_저장_재열기에도_결정이_남는다(tmp_path):
    """★사람의 결정이 하룻밤을 못 넘기던 자리.

    종전 io.py 는 apply(갱신만) → require(미지정 삽입) 순서라, 레코드 없던
    헤드에 찍은 종류가 재열기마다 미지정으로 되돌아갔다 — 화면까지도.
    """
    b = _bare_board()
    b.set_head_kind(b.disks[0], "하향식")
    write_edits(b, str(tmp_path))

    b2 = _bare_board()                      # 같은 도면을 새로 연 상태
    assert load_edits(b2, str(tmp_path)) is True
    assert _engine_kind_at(b2, (0.0, 0.0)) == "하향식", \
        "재열기하니 사람의 결정이 미지정으로 되돌아갔다"
    assert b2.disk_kinds[0] == "하향식"
    assert b2.disk_kinds == disk_kind_list(b2.disks, b2.head_kinds)


def test_지워진_헤드의_override_는_유령을_만들지_않는다():
    """resolve 가 «삽입하는 apply» 가 아닌 이유.

    override 는 이제 없는 헤드를 가리킬 수 있다(찍고 나서 헤드를 지운 경우).
    그걸 레코드로 끼워 넣으면 종류 집계·미확정 게이트에 유령이 낀다 —
    살아 있는 디스크 목록은 require 만 알므로, 세우는 쪽도 require 다.
    """
    disks = [(0.0, 0.0, 150.0)]
    ghost = [{"c": [9999.0, 9999.0], "r": 150.0, "kind": "하향식"}]
    out = resolve_head_kinds(disks, [], ghost)
    assert len(out) == 1, f"유령 레코드가 끼었다: {out}"
    assert tuple(out[0]["c"]) == (0.0, 0.0)
    assert out[0]["kind"] == "미지정"


def test_확정한_헤드를_미확정이라며_막지_않는다():
    """변환 게이트에서의 증상 — 사람은 확정했는데 «미지정 헤드» 로 차단됐다."""
    got = preflight_kfp_convert({
        "hcov": [(100.0, 0.0, 5.0)],
        "head_kinds": [],                    # 분류 실패
        "kind_overrides": [{"c": [100.0, 0.0], "r": 5.0, "kind": "상향식"}],
    })
    assert got["ok"] is True, got["blockers"]
