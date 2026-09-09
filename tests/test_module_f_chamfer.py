# -*- coding: utf-8 -*-
"""[모서리] 45° 챔퍼와 «같은 자리의 두 점» — 평면 ↔ 아이소 위상 1:1.

■ 목적 (2026-09-09 오너)

  「평면도랑 아이소로 바꿀 때랑 위상이 일대일 대응. 누락되거나 휘는 값 없이.」

  등각에서 나올 수 있는 각도는 셋뿐이다 — 평면 가로 +30° · 평면 세로 +150° ·
  헤드 스텁 수직. 평면의 45° 조각은 그 격자를 벗어나 **휜다.**

■ 사용자가 배관 11개를 지목한 데서 시작했다

  노드로 이으면 한 줄기인데 관경이 65 → 40 → 65 로 튀었고, 담당 헤드 1개짜리에
  65A 가 붙어 있었다. 재 보니 뿌리는 «조각남» 이었다 (대명동 K=30)::

      배관 262 중 0.3 m 미만 조각 140 (53 %) · 그중 45° 대각 91
      조각 140 중 «양 끝 차수 2»(한 배관을 자른 것) 105
      → 조각마다 근처 치수 텍스트를 **제각기** 물어 관경이 튄다

  ★역참조(`edge_ref`)는 무죄였다 — 가리키는 자리와 실제 자리의 어긋남 0.

■ 그리고 조각남의 뿌리가 하나 더 있었다

      880(275401,-234671) 차수 3 – 1227 · 챔퍼 70 mm
          880–2243     0.0 mm     ← ★길이 0
          1227–2455    0.0 mm     ← ★길이 0

  같은 자리에 점이 둘씩 있고 길이 0 간선으로 이어져 있다(도면이 배관을 두 줄로
  그린 것). 그 탓에 차수가 3·4 로 부풀어 챔퍼를 **못 편다**(218곳이 막혔다).

■ 실측 — 셋을 나란히 (대명동 K=30 · `scripts/_probe_flex_fold.py`)

                        (가) 둘 다 끔   (나) 챔퍼만   (다) 둘 다
      등각 격자 벗어남      99            90           **29**
      평면 비직각          173           163           82
      물닿음 헤드          111           111          **111**
      배관 수              262           259          221
      표 전체 총연장       158,083      160,898      175,874
      평면 X 교차           12            14           18

  ★「형상은 안 바뀐다」는 **틀린 말이었다.** 좌표는 그대로지만 겹쳐 그린 두
    줄이 하나가 되면서 최불리 경로 선택이 달라진다 — 총연장 +9.4 %, X 교차 +4.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _ch():
    from routes.module_f import chamfer
    return chamfer


class _B:
    def __init__(self, pts, edges, sources=(), valves=()):
        self.pts = list(pts)
        self.edges = frozenset(tuple(sorted(e)) for e in edges)
        self.sources = list(sources)
        self.valves = list(valves)


# ─────────────────────────────── 같은 자리의 두 점
def test_길이_0_간선을_접는다():
    ch = _ch()
    b = _B([(0.0, 0.0), (100.0, 0.0), (100.0, 0.0), (100.0, 100.0)],
           [(0, 1), (1, 2), (2, 3)])
    got = ch.collapse_zero_edges(b)
    assert got["zero_edges"] == 1 and got["merged"] == 1
    # 1 과 2 가 하나가 되고, 길이 0 간선은 사라진다.
    assert b.edges == frozenset({(0, 1), (1, 3)}), b.edges


def test_겹쳐_그린_두_줄이_하나가_된다():
    """★도면이 배관을 두 줄로 그리면 중복 간선이 생긴다 — set 이 하나로 만든다."""
    ch = _ch()
    #  0–1 과 2–3 이 완전히 같은 자리에 겹쳐 있다.
    b = _B([(0.0, 0.0), (100.0, 0.0), (0.0, 0.0), (100.0, 0.0)],
           [(0, 1), (2, 3), (0, 2), (1, 3)])
    ch.collapse_zero_edges(b)
    assert b.edges == frozenset({(0, 1)}), b.edges


def test_급수원이_가리키던_점이_사라지면_따라간다():
    """★급수원·알람밸브는 **인덱스로** 산다 — 안 옮기면 가리키는 자리가 어긋난다."""
    ch = _ch()
    b = _B([(0.0, 0.0), (0.0, 0.0), (100.0, 0.0)],
           [(0, 1), (1, 2)], sources=[1], valves=[1])
    ch.collapse_zero_edges(b)
    assert b.sources == [0] and b.valves == [0], (b.sources, b.valves)


def test_떨어진_점은_안_합친다():
    ch = _ch()
    b = _B([(0.0, 0.0), (100.0, 0.0)], [(0, 1)])
    got = ch.collapse_zero_edges(b)
    assert got["zero_edges"] == 0 and b.edges == frozenset({(0, 1)})


# ─────────────────────────────── 챔퍼 → 모서리
def test_챔퍼를_모서리로_편다():
    """가로 – 45° – 세로 를 모서리 하나로. 그러면 등각 격자 위에만 남는다."""
    ch = _ch()
    #  0 ─가로─ 1 ─45°─ 2 ─세로─ 3
    b = _B([(0.0, 0.0), (100.0, 0.0), (200.0, 100.0), (200.0, 300.0)],
           [(0, 1), (1, 2), (2, 3)])
    got = ch.restore_corners(b)
    assert got["corners"] == 1, got
    # 챔퍼가 사라지고 1 이 모서리(200, 0)로 옮겨 간다.
    assert b.edges == frozenset({(0, 1), (1, 3)}), b.edges
    assert abs(b.pts[1][0] - 200.0) < 1e-6 and abs(b.pts[1][1]) < 1e-6, b.pts[1]


def test_모서리를_도는_만큼_길어진다():
    """★줄지 않는다 — 줄면 마찰손실이 과소 산정되고 그것은 비보수측이다."""
    ch = _ch()
    b = _B([(0.0, 0.0), (100.0, 0.0), (200.0, 100.0), (200.0, 300.0)],
           [(0, 1), (1, 2), (2, 3)])
    got = ch.restore_corners(b)
    assert got["grew_mm"] > 0, got


def test_실제_대각_주행은_안_건드린다():
    """★접기 지시서 §6 — 가지관의 2~7 m 45° 는 실제 주행이라 펴면 망가진다."""
    ch = _ch()
    #  0 ─가로─ 1 ─45°─ 2 ─**45° 대각 주행**─ 3
    b = _B([(0.0, 0.0), (100.0, 0.0), (200.0, 100.0), (5000.0, 4900.0)],
           [(0, 1), (1, 2), (2, 3)])
    got = ch.restore_corners(b)
    assert got["corners"] == 0, got
    assert "양옆이 직교가 아님 (실제 대각 주행)" in got["why"], got["why"]


def test_긴_45도는_챔퍼가_아니다():
    """챔퍼와 주행은 **길이로 갈린다** — 대명동 실측 챔퍼 71·141 mm."""
    ch = _ch()
    b = _B([(0.0, 0.0), (100.0, 0.0), (2100.0, 2000.0), (2100.0, 5000.0)],
           [(0, 1), (1, 2), (2, 3)])
    assert ch.restore_corners(b)["corners"] == 0


def test_분기가_걸린_조각은_안_건드린다():
    """지우면 그 가지가 떨어진다 — 실측으로 218곳이 이 사유였다."""
    ch = _ch()
    b = _B([(0.0, 0.0), (100.0, 0.0), (200.0, 100.0), (200.0, 300.0),
            (200.0, -300.0)],
           [(0, 1), (1, 2), (2, 3), (2, 4)])
    got = ch.restore_corners(b)
    assert got["corners"] == 0
    assert "양 끝 차수≠2 (분기가 걸려 있다)" in got["why"]


def test_급수원이_찍힌_점은_안_없앤다():
    ch = _ch()
    b = _B([(0.0, 0.0), (100.0, 0.0), (200.0, 100.0), (200.0, 300.0)],
           [(0, 1), (1, 2), (2, 3)], sources=[2])
    got = ch.restore_corners(b)
    assert got["corners"] == 0
    assert "급수원·알람밸브가 가리키는 점" in got["why"]


# ─────────────────────────────── 사실을 사실대로 적었는가
def test_형상이_안_바뀐다고_말하지_않는다():
    """★한 번 그렇게 적었다가 실측으로 틀렸다 — 총연장 +9.4 % · X 교차 +4."""
    src = (_ROOT / "routes" / "module_f" / "chamfer.py").read_text(
        encoding="utf-8")
    assert "형상은 한 톨도 안 변한다" not in src
    # ★줄바꿈에 안 걸리게 «공백을 접어» 찾는다. 문장이 두 줄에 걸쳐 있으면
    #   한 줄 문자열로는 못 찾는다 — 이 시험이 그렇게 한 번 헛돌았다.
    flat = " ".join(src.split())
    assert "좌표가 안 바뀌는 것과 망이 안 바뀌는 것은 다른 이야기다" in flat
    api = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    assert "형상은 안 바뀝니다" not in api


def test_계측_스위치가_있다():
    """어느 조작이 무엇을 바꿨는지 가르려면 하나씩 꺼 봐야 한다."""
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    assert "MF_NO_ZERO_COLLAPSE" in src and "MF_NO_CHAMFER" in src
