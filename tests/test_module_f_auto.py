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
    LABEL_OFFSET_FOR_AUTO, AutoError, head_region_of, preview_view,
    region_around, summarize)
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
        # ★모듈 A 의 표가 실제로 내는 키를 쓴다. 종전에는 여기에
        #   「앵커 노드」를 손으로 넣어 두어, `summarize` 가 그 키를 읽는 것이
        #   **실제로는 늘 None** 이라는 사실을 이 시험이 못 봤다(A 의 meta 에는
        #   그런 키가 아예 없다 — 그건 모듈 G 의 표가 내는 것이다).
        self.meta = [("알람밸브 좌표 (snap)", "(100.0, 0.0)")]


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
def test_영역이_없으면_검출에서_만든다():
    """영역은 «좁히는» 선택이지 시작 조건이 아니다 — 안 그리면 헤드 전부."""
    heads = [{"x": 0, "y": 0}, {"x": 1000, "y": 500}, {"x": -200, "y": 800}]
    r = region_around(heads, pad_mm=100.0)
    assert len(r) == 1
    x0, y0, x1, y1 = r[0]
    assert x0 == -300.0 and y0 == -100.0        # 최소 − 여유
    assert x1 == 1100.0 and y1 == 900.0         # 최대 + 여유
    # 만든 사각형은 모든 헤드를 담는다
    reg = head_region_of(r)
    assert all(reg.contains((h["x"], h["y"])) for h in heads)


def test_헤드가_없으면_범위를_못_만든다():
    """0개인데 조용히 넘어가면 «영역 밖» 으로 오해한다 — 도면 탓임을 말한다."""
    with pytest.raises(AutoError, match="헤드를 찾지 못"):
        region_around([])


# ── 한 파일에 도면 여러 장 (실측 B1F: 헤드 3,338 · 전부의 bbox 2,241 x 2,172 m
#    인데 실제 도면은 253 x 142 m — 도면 밖 이상점이 범위를 9배로 부풀렸다)
def _sheet(x0, y0, rows, cols=6, step=2000.0):
    """정상 간격(2 m)으로 놓인 헤드 격자 하나 — 도면 한 장 흉내."""
    return [{"x": x0 + c * step, "y": y0 + r * step}
            for r in range(rows) for c in range(cols)]


def _spans(rect):
    return rect[2] - rect[0], rect[3] - rect[1]


def test_여러_장이면_알람밸브가_놓인_장으로_좁힌다():
    """전부의 bbox 로 잡으면 다른 장의 헤드까지 최불리 후보가 된다."""
    near = _sheet(0, 0, rows=5)                    # 30개 · 10 x 8 m
    far = _sheet(500_000, 0, rows=5)               # 500 m 떨어진 다른 장
    r = region_around(near + far, (0.0, 0.0), pad_mm=0.0)[0]

    w, h = _spans(r)
    assert w < 20_000, f"장을 안 갈랐다 — 범위 폭 {w:,.0f} mm"
    reg = head_region_of([r])
    assert all(reg.contains((p["x"], p["y"])) for p in near)
    assert not any(reg.contains((p["x"], p["y"])) for p in far)


def test_알람밸브를_반대쪽_장에_찍으면_그쪽이_잡힌다():
    near = _sheet(0, 0, rows=5)
    far = _sheet(500_000, 0, rows=5)
    r = region_around(near + far, (500_000.0, 0.0), pad_mm=0.0)[0]

    reg = head_region_of([r])
    assert all(reg.contains((p["x"], p["y"])) for p in far)
    assert not any(reg.contains((p["x"], p["y"])) for p in near)


def test_알람밸브가_어느_장에도_없으면_헤드가_많은_장():
    """경계 밖에 찍혀도 «전부» 로 되돌아가지는 않는다."""
    big = _sheet(0, 0, rows=6)                     # 36개
    small = _sheet(500_000, 0, rows=3)             # 18개
    r = region_around(big + small, (250_000.0, 900_000.0), pad_mm=0.0)[0]

    reg = head_region_of([r])
    assert all(reg.contains((p["x"], p["y"])) for p in big)
    assert not any(reg.contains((p["x"], p["y"])) for p in small)


def test_한_장짜리는_헤드_전부를_그대로_쓴다():
    """멀쩡한 한 장을 괜히 쪼개면 설계면적이 잘려 나간다."""
    one = _sheet(0, 0, rows=6)
    r = region_around(one, (0.0, 0.0), pad_mm=0.0)[0]

    reg = head_region_of([r])
    assert all(reg.contains((p["x"], p["y"])) for p in one)


def test_자동_추출이_알람밸브를_범위_생성에_넘긴다():
    """안 넘기면 «헤드가 많은 장» 으로만 물러서, 사람이 찍은 곳이 무시된다."""
    import inspect

    from routes.module_f import auto
    src = inspect.getsource(auto.run_auto)
    assert "region_around(cand, (float(alarm_xy[0]), float(alarm_xy[1])))" in src


def test_빈_사각형은_여전히_올린다():
    with pytest.raises(AutoError, match="비었"):
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
    # 자동 차선이 «실제로 아는 것» — 알람밸브를 어디에 스냅했나.
    assert s["alarm_snap"] == "(100.0, 0.0)"
    assert "anchor_label" not in s, "늘 None 이던 칸이 되살아났다"


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
