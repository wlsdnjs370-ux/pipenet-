# -*- coding: utf-8 -*-
"""[A 방식] 평면도 자동 추출 — 수동(E)과 같은 자리에 놓이는가.

같은 평면도에서 같은 것(최불리 헤드군)을 뽑는 길이 둘이고, 그 결과가 하류
(수리계산 표 · 통합 · 산출)에서 구분 없이 받아져야 한다. 여기서 못박는 것:

  ① 라벨 오프셋이 경로마다 갈린다 — A 는 이미 10, G 는 1부터라 +9
  ② 영역 없이는 자동 추출이 성립하지 않는다 (anchored 의 필수 인자)
  ③ 요약이 수동 경로와 같은 것을 말한다
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from routes.module_f.auto import (  # noqa: E402
    LABEL_OFFSET_FOR_AUTO, AutoError, head_region_of, preview_view, summarize)
from routes.module_f.merge import (  # noqa: E402
    ANCHOR_LABEL, LABEL_OFFSET, label_offset_for, to_head_tables)


# ─────────────────────────────────────────── ① 라벨 오프셋
def test_자동은_옮기지_않고_수동은_9만큼():
    """A 의 표는 처음부터 10 이다 — 또 +9 하면 기준점이 19 가 된다."""
    assert label_offset_for("auto") == 0 == LABEL_OFFSET_FOR_AUTO
    assert label_offset_for("manual") == LABEL_OFFSET == 9
    assert label_offset_for(None) == LABEL_OFFSET      # 모르면 수동으로 본다
    assert label_offset_for("AUTO") == 0               # 대소문자 무관


class _T:
    """A 의 `PipeTables` 모양 — 급수원이 이미 10."""

    def __init__(self):
        self.nodes = [
            {"label": "10", "x": 0, "y": 0, "elevation": 0.0, "io_node": "Input"},
            {"label": "11", "x": 100, "y": 0, "elevation": 0.0, "io_node": "No"},
        ]
        self.pipes = [{"label": "P1", "in": "10", "out": "11", "dia": 65,
                       "length": 1.0}]
        self.nozzles = [{"label": "1", "in": "11", "out": "@/1"}]
        self.fittings = []
        self.equipment = []
        self.meta = [("앵커 노드", "11")]


def test_자동_표는_그대로_통과한다():
    ht = to_head_tables(_T(), offset=label_offset_for("auto"))
    assert [n["label"] for n in ht.nodes] == ["10", "11"]
    anchor = next(n for n in ht.nodes if n["label"] == ANCHOR_LABEL)
    assert anchor["io_node"] == "Input"


def test_자동_표에_9를_먹이면_기준점이_사라진다():
    """왜 갈라야 하는지 — 잘못 먹이면 결합이 성립하지 않는다."""
    from routes.module_f.merge import MergeError
    with pytest.raises(MergeError, match="기준점"):
        to_head_tables(_T(), offset=LABEL_OFFSET)


# ─────────────────────────────────────────── ② 필수 입력
def test_영역이_없으면_올린다():
    with pytest.raises(AutoError, match="영역"):
        head_region_of(None)
    with pytest.raises(AutoError, match="영역"):
        head_region_of([])


def test_영역은_A의_HeadRegion이_된다():
    r = head_region_of([[0, 0, 100, 100]])
    assert r.contains((50, 50)) is True
    assert r.contains((150, 50)) is False


def test_영역_여러개는_합집합():
    r = head_region_of([[0, 0, 10, 10], [100, 100, 110, 110]])
    assert r.contains((5, 5)) and r.contains((105, 105))
    assert not r.contains((50, 50))


# ─────────────────────────────────────────── ③ 요약·미리보기
class _Sel:
    heads = [1, 2, 3]
    distances = [1000.0, 2500.0, 4000.0]
    edges = [("a", "b", 1.0)] * 5
    source_bridge_dist_mm = 0.0
    source_fallback = False


def test_요약이_최원_최근을_m로_낸다():
    s = summarize(_Sel(), _T())
    assert s["k"] == 3
    assert s["far_m"] == 4.0 and s["near_m"] == 1.0
    assert s["nodes"] == 2 and s["pipes"] == 1 and s["nozzles"] == 1
    assert s["anchor_label"] == "11"


def test_요약은_빈_선정도_받는다():
    class _Empty:
        heads = []
        distances = []
        edges = []
    s = summarize(_Empty(), _T())
    assert s["k"] == 0 and s["far_m"] == 0.0 and s["near_m"] == 0.0


def test_급수원_대체를_숨기지_않는다():
    """멀어서 최근접으로 갈아탄 것은 사람이 알아야 한다."""
    class _Fb(_Sel):
        source_fallback = True
        source_bridge_dist_mm = 1234.5
    s = summarize(_Fb(), _T())
    assert s["source_fallback"] is True
    assert s["source_bridge_mm"] == 1234.5


def test_미리보기가_헤드와_급수원을_표시한다():
    v = preview_view(_T())
    assert len(v["nodes"]) == 2 and len(v["pipes"]) == 1
    by = {n["label"]: n for n in v["nodes"]}
    assert by["10"].get("input") is True
    assert by["11"].get("head") is True
    assert by["10"].get("head") is None
