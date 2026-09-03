# -*- coding: utf-8 -*-
"""헤드를 하나도 안 찍은 스펙 — «없음» 과 «비어 있음» 은 다르다 (BLOCKED §16).

■ 무엇이 죽었나

  자동 인식이 0 개인 도면(LH306 · 높음 띠 0/42)에서 찍기가 스펙을 쓰면
  `heads` 칸이 **아예 없었다** — `pick/board.spec()` 이 `if self.heads:` 로
  감싸고 있었다. 읽는 쪽(`flow.spots_body`)은 그 칸이 있다고 가정해
  `spec["heads"]` 로 곧장 읽었고, 결과는 `KeyError: 'heads'` 였다.

  사람이 본 것은 사유가 아니라 그 낱말 하나다. 「이 도면은 헤드를 못 찾았습니다」
  가 아니라 「KeyError: 'heads'」 였다.

■ 왜 오래 열려 있었나

  당시 지시서에서 `cad_project_editor_g/` 가 읽기 전용이라, F 의 문 앞에서
  막는 것으로 우회했다(D-F11-2 의 지배 띠 규칙이 그 뒤 대부분을 해소했다).
  게이트는 최후 방어로 남기고, 이제 엔진 쪽을 고친다 — **막을 것이 아니라
  그냥 지나가야 할 일**이기 때문이다. 헤드 0 은 오류가 아니다.

■ 고친 자리 둘

  쓰는 쪽 — 칸을 비워서라도 **쓴다**(안 찍은 것은 «비어 있음» 이다).
  읽는 쪽 — `spec.get("heads") or ()` (같은 파일 250 행이 이미 그렇게 읽었다).
  한쪽만 고치면 옛 스펙 파일이나 옛 찍기가 남아 다시 죽는다.
"""
from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)


def test_헤드를_안_찍어도_스펙에_칸이_있다():
    """쓰는 쪽 — `if self.heads:` 로 칸을 통째로 빼지 않는다."""
    from services.cad_import.pick.board import Board

    b = Board.__new__(Board)              # world 없이 spec() 만 본다
    b.mat = []
    b.heads = []
    b.kn = {"small_r": 300.0}

    class _W:
        arcs = ()
        arc_ang = ()
    b.w = _W()
    sp = b.spec()
    assert "heads" in sp, "헤드 0 이라고 칸을 빼면 읽는 쪽이 죽는다"
    assert sp["heads"] == []


def test_칸이_없는_옛_스펙도_안_죽인다():
    """읽는 쪽 — 옛 파일·옛 찍기가 남아 있어도 지나가야 한다.

    ★한쪽만 고치면 안 되는 이유가 이것이다. 쓰는 쪽을 고쳐도 이미 디스크에
      있는 스펙은 그대로다.
    """
    import inspect

    from services.cad_import.pipeline import flow

    src = inspect.getsource(flow.spots_body)
    assert 'spec["heads"]' not in src, \
        "칸을 직접 읽는다 — 옛 스펙에서 KeyError 로 죽는다"
    assert 'spec.get("heads")' in src


def test_헤드_0_은_오류가_아니다():
    """`require_head_kinds` 도 빈 목록을 그대로 받는다 — 여기서 막지 않는다."""
    from services.cad_import.kinds import disk_kind_list, require_head_kinds

    assert require_head_kinds([], []) == []
    assert require_head_kinds([], None) == []
    assert disk_kind_list([], []) == []
