# -*- coding: utf-8 -*-
"""지시서 §12 — 설계 워크벤치 페이지.

화면은 눈으로 보면 되니 테스트할 게 없어 보이지만, **조용히 사라지는 것**이 둘 있다.
가상 폐합선의 점선·경고색(§12.5)과 CDN 금지(외부 AI 격리 증빙)다. 둘 다 지워져도
화면은 멀쩡히 뜨기 때문에 사람 눈으로는 회귀를 못 잡는다.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest
from flask import Flask

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

import routes.r30_design as r30_design  # noqa: E402

_HTML = (_ROOT / "templates" / "design_workbench.html").read_text(encoding="utf-8")


@pytest.fixture()
def client(tmp_path):
    app = Flask(__name__, template_folder=str(_ROOT / "templates"))
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=True)
    return app.test_client()


def test_플래그가_꺼져_있으면_페이지도_없다(tmp_path):
    """페이지만 열리고 API 가 404 면 화면은 이유를 알 수 없는 고장으로 보인다."""
    app = Flask(__name__, template_folder=str(_ROOT / "templates"))
    r30_design.register(app, DESIGN_SESSION_DIR=tmp_path, enabled=False)
    assert app.test_client().get("/design-workbench").status_code == 404


def test_페이지가_뜨고_캐시되지_않는다(client):
    res = client.get("/design-workbench")
    assert res.status_code == 200
    assert "no-store" in res.headers["Cache-Control"]


@pytest.mark.parametrize("dom_id", [
    "dw-canvas", "dw-dxf", "dw-stepper", "dw-layer-list", "dw-design-list",
])
def test_지시서가_정한_dom_id_를_쓴다(dom_id):
    """§12.2 — 뒤 PR 들이 이 id 로 붙는다. 이름이 흔들리면 그때 조용히 null 이 된다."""
    assert f'id="{dom_id}"' in _HTML


def test_원본_워크벤치의_id_를_물려받지_않았다():
    """복제본이라 `wb-*` 가 남기 쉽다. 남으면 두 화면이 같은 id 를 놓고 헷갈린다."""
    assert not re.search(r'id="wb-', _HTML)


def test_가상_폐합선은_점선_경고색이다():
    """추정 연결을 실측 벽과 같은 실선으로 그리면 검수자가 구분할 방법이 없다."""
    block = _HTML.split("virtualEdges && state.virtualEdges.length")[1][:400]
    assert "setLineDash([6, 4])" in block
    assert "DESIGN_COLORS.virtualEdge" in block
    assert 'virtualEdge: "#f87171"' in _HTML


def test_건축_12종이_전부_색과_렌더_순서를_갖는다():
    cats = ["WALL", "DOOR", "WINDOW", "COLUMN", "STAIR", "SHAFT",
            "ROOM_TEXT", "DIM", "FURNITURE", "GRID", "BEAM", "OTHER"]
    order = _HTML.split("const ARCH_ORDER = [")[1].split("];")[0]
    colors = _HTML.split("const ARCH_COLORS = {")[1].split("};")[0]
    for cat in cats:
        assert f'"{cat}"' in order, f"{cat} 이 렌더 순서에 없다"
        assert f"{cat}:" in colors, f"{cat} 에 색이 없다"


def test_기존_6종_자동분류를_12종으로_옮기지_않는다():
    """`auto_category` 는 별개 체계다(§3.2). 그대로 옮기면 근거 없는 주장이 된다."""
    assert "category: UNCLASSIFIED" in _HTML
    assert "category: layer.auto_category" not in _HTML


def test_외부_cdn_을_부르지_않는다():
    """LH 외부AI 격리 증빙 — 이 화면은 망 밖으로 아무것도 요청하지 않는다."""
    hits = re.findall(r'(?:src|href)\s*=\s*["\']https?://[^"\']+', _HTML)
    assert not hits, f"외부 리소스 참조 {hits}"
