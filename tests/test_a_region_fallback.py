# -*- coding: utf-8 -*-
"""모듈 A — 영역을 «안 그렸다» 는 이유로 옛 경로에 떨어지지 않는다.

A 화면은 `use_anchored = bool(alarm_xy) and bool(zones)` 로 두 알고리즘을
조용히 갈랐다. 영역을 안 그리면 비-anchored 로 떨어진다. 실측(B1F 110MB ·
scripts/_probe_a_fallback.py · 알람밸브 동일):

      앵커 경로   급수원 결합    14 mm · 영역 게이트 있음
      옛  경로   급수원 결합   152 mm · 영역 게이트 없음

헤드군은 둘 다 34.0 m 로 뭉쳤다 — 옛 경로가 «틀린» 것은 아니다. 다만 급수원을
10배 먼 데서 잡고, 영역 게이트가 없어 범례·다른 장의 문자까지 관경 판독에
섞이며, load_mode 가 조용히 무시된다. 사람은 그것을 화면에서 알 수 없다.

영역은 헤드군을 «좁히는» 선택이지 시작 조건이 아니므로, 알람밸브만 찍혀 있으면
검출한 헤드에서 범위를 만들어 앵커 경로로 간다. 알람밸브까지 없으면 앵커를
세울 수 없어 여전히 옛 경로이고, 그 사실을 응답이 말한다.
"""
from __future__ import annotations

import inspect
import os
import re
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from remote30_prototype import (  # noqa: E402
    head_bbox_for_region, sheet_frame_at)


def _sheet(x0, y0, rows, cols=6, step=2000.0):
    return [(x0 + c * step, y0 + r * step)
            for r in range(rows) for c in range(cols)]


# ─────────────────────────────── ① 공용 규칙 (A 화면과 F 자동이 같이 쓴다)
def test_여러_장이면_알람밸브가_놓인_장():
    near, far = _sheet(0, 0, 5), _sheet(500_000, 0, 5)
    f = sheet_frame_at(near + far, (0.0, 0.0))
    assert f is not None
    x0, _, x1, _ = f["bbox"]
    assert x0 <= 0 <= x1 and x1 < 500_000


def test_한_장짜리는_가르지_않는다():
    assert sheet_frame_at(_sheet(0, 0, 6), (0.0, 0.0)) is None


def test_범위는_다른_장을_빼고_만든다():
    near, far = _sheet(0, 0, 5), _sheet(500_000, 0, 5)
    (x0, y0, x1, y1), = head_bbox_for_region(near + far, (0.0, 0.0), pad_mm=0.0)
    assert x1 - x0 < 20_000, f"장을 안 갈랐다 — 폭 {x1 - x0:,.0f} mm"
    assert all(x0 <= p[0] <= x1 and y0 <= p[1] <= y1 for p in near)
    assert not any(x0 <= p[0] <= x1 for p in far)


def test_헤드가_없으면_빈_목록():
    """A 쪽은 올리지 않고 빈 목록으로 답한다 — 부르는 쪽이 종전대로 간다."""
    assert head_bbox_for_region([]) == []


def test_여유는_기호_반경을_감싸도록_바깥으로():
    """헤드 중심만으로 자르면 기호와 접속관 끝이 경계 밖으로 나간다."""
    (x0, y0, x1, y1), = head_bbox_for_region([(0.0, 0.0), (1000.0, 500.0)],
                                             pad_mm=100.0)
    assert (x0, y0, x1, y1) == (-100.0, -100.0, 1100.0, 600.0)


# ─────────────────────────────── ② A 화면의 배선
def _combined_src() -> str:
    path = os.path.join(_ROOT, "routes", "r30_combined.py")
    return open(path, encoding="utf-8").read()


def test_영역이_없으면_검출에서_만들어_앵커로_간다():
    src = _combined_src()
    i = src.index("use_anchored = bool(alarm_xy) and bool(zones)")
    before = src[max(0, i - 1600):i]
    assert "if alarm_xy and not zones:" in before, \
        "영역이 없을 때 만들어 쓰는 자리가 없다 — 옛 경로로 떨어진다"
    assert "head_bbox_for_region(" in before, "공용 규칙을 안 쓴다"
    assert "region_auto = True" in before


def test_만들기가_실패해도_추출은_돌아간다():
    """범위를 못 만드는 도면에서 500 을 내면 안 된다 — 종전대로 간다."""
    src = _combined_src()
    i = src.index("if alarm_xy and not zones:")
    body = src[i:i + 900]
    assert "except Exception" in body and "_auto = []" in body


def test_응답이_어느_경로였는지_말한다():
    """조용히 갈리면 사람이 «이상하다» 고만 느낀다."""
    src = _combined_src()
    assert '"anchored": use_anchored,' in src
    assert '"region_auto": region_auto,' in src


def test_사람이_그린_영역은_그대로_쓴다():
    """자동 생성은 «안 그렸을 때» 만 — 그린 것을 덮어쓰면 안 된다."""
    src = _combined_src()
    i = src.index("if alarm_xy and not zones:")
    j = src.index("use_anchored = bool(alarm_xy) and bool(zones)", i)
    assert re.search(r"\bzones = _auto\b", src[i:j])
    # 조건 밖에서 zones 를 갈아엎는 곳이 없다
    assert src.count("zones = _auto") == 1


# ─────────────────────────────── ③ F 자동도 같은 규칙을 쓴다
def test_F_자동은_A의_규칙을_그대로_부른다():
    from routes.module_f import auto
    assert "head_bbox_for_region" in inspect.getsource(auto.region_around)
    assert "sheet_frame_at" in inspect.getsource(auto.sheet_of)
