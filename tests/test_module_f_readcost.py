# -*- coding: utf-8 -*-
"""「읽기」 뒤의 기다림 — 헛일 두 가지를 못박는다.

  ① 자동은 A 의 «캐시» 파서를 써야 한다. 안 쓰면 같은 도면을 열 때마다
     통째로 다시 읽는다(실측 LH306 16MB: 5.02s vs 캐시 0.06s — 86배).
  ② 수동은 이미 받아 둔 도면을 다시 받지 않아야 한다. 「불러오기」가 방금
     내려준 것을 「수동으로 읽기」가 또 받으면 그만큼이 통째로 헛일이다
     (실측 1.44 MB).
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


# ─────────────────────────────────────────── ① 캐시 파서
def test_자동은_캐시_파서를_쓴다():
    from routes.module_f import auto
    src = inspect.getsource(auto.parse_plan)
    assert "parse_dxf_bundle_cached" in src, "캐시 없는 파서를 쓰고 있다"


def test_캐시_파서가_실제로_있다():
    """A 쪽에서 이름이 바뀌면 여기서 먼저 깨진다."""
    import remote30_prototype as A
    assert callable(getattr(A, "parse_dxf_bundle_cached", None))


def test_캐시는_내용_해시로_건다():
    """mtime 으로 걸면 같은 파일을 다시 올리기만 해도 캐시가 죽는다 —
    handoff 가 그 함정에 빠졌던 적이 있다."""
    import remote30_prototype as A
    src = inspect.getsource(A.parse_dxf_bundle_cached)
    assert "_file_content_key" in src
    assert "mtime" not in src


# ─────────────────────────────────────────── ② 재다운로드
def _script() -> str:
    """화면 JS 본문 — 인라인이든 정적 파일이든 같은 것을 돌려준다.

    자산을 `static/module_f.js` 로 떼어낸 뒤로 템플릿엔 <script src> 만 남는다.
    이 시험이 보는 것은 «코드가 무엇을 하는가» 지 «어디에 적혀 있는가» 가
    아니므로, 출처를 가리지 않고 읽는다.
    """
    path = os.path.join(_ROOT, "templates", "module_f.html")
    html = open(path, encoding="utf-8").read()
    bodies = [b for b in re.findall(r"<script[^>]*>(.*?)</script>", html, re.S)
              if b.strip()]
    if bodies:
        return max(bodies, key=len)
    return open(os.path.join(_ROOT, "static", "module_f.js"),
                encoding="utf-8").read()


def test_수동은_받아_둔_도면을_다시_받지_않는다():
    src = _script()
    i = src.index("async function loadWorld(")
    body = src[i:i + 400]
    # reuse 를 받고, 그때는 다시 받지 않는다
    assert "loadWorld(reuse)" in body, "재사용 인자가 없다"
    assert "if (!reuse || !S.world) await loadWorldRaw();" in body, \
        "재사용 여부와 무관하게 다시 받는다"


def test_읽기_단계가_재사용으로_부른다():
    src = _script()
    i = src.index("async function readSlot(")
    body = src[i:i + 900]
    assert "loadWorld(true)" in body, "읽기 단계가 도면을 또 받는다"


def test_불러오기는_처음이라_그대로_받는다():
    """열 때는 재사용할 것이 없다 — 그때까지 아낀다고 건너뛰면 화면이 빈다.

    ★«앞에서 1400자» 로 자르지 않는다. 그 창은 «도면을 받나» 가 아니라
      «그 줄이 얼마나 가까이 붙어 있나» 를 재는 자였다 — 실제로 주석이
      늘자(이른 그리기 설명) 동작은 그대로인데 빨간불이 켜졌다. 핸들러의
      **끝**은 다음 최상위 선언이 정한다.
    """
    src = _script()
    i = src.index('$("btn-open").onclick')
    j = src.index("async function readSlot(", i)   # 핸들러 바로 다음 선언
    body = src[i:j]
    assert "loadWorldRaw()" in body, "불러오기가 도면을 안 받는다"
