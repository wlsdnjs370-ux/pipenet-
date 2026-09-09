# -*- coding: utf-8 -*-
"""[헤드배관 꼬임] 신축배관·X교차·등각 겹침 — 지시서 `ModuleF_헤드배관_꼬임_수정지시서.md`.

■ 무엇이 문제였나 (대명동 단위세대 평면도 · 배관 211개)

  등각에서 나올 수 있는 각도는 셋뿐이다 — 평면 가로 +30° · 평면 세로 +150° ·
  헤드 스텁 수직. 그런데 **31개가 그 격자를 벗어나** 있었다. 평면으로 되돌리면
  전부 «모서리를 자른 141 mm 45° 챔퍼» 이고, 66%가 헤드에서 2홉 자리다.

  정체는 **신축배관(SP 후렉시블)이 배관망 재료로 들어온 것**이다. 색 6 이
  가지관(`-소화(SP가지관)`)과 겹쳐, 가지관을 찍으면 신축배관이 같이 딸려온다.
  실측: 재료 9묶음 중 «SP 후렉시블» 색 6·7 (선분 320개).

■ 여기서 지키는 것

  ⑴ 레이어를 «접기» 로 지정할 수 있다 — 자동 사전은 **추천만**, 확정은 사람이.
  ⑵ 접은 것을 조용히 넘기지 않는다(세션·로그에 남는다 · S340).
  ⑶ X자 교차는 **세기만** 한다 — 자동으로 쪼개면 없던 분기가 생겨 유량이 갈린다.
  ⑷ 등각에서 겹쳐 보이는 접속관은 «위상 문제가 아니다» 라고 화면이 말한다.

■ ★C-0 계측 결과 → 그리고 그 다음에 일어난 일

  신축배관을 빼면 대명동에서 물닿음 헤드가 **111 → 5** 로 떨어진다.
  신축배관이 헤드를 가지관에 잇는 유일한 경로이기 때문이다.

  그래서 「빼기」는 **철회됐다**(`ModuleF_신축배관_접기_지시서.md`). 대신
  **접기** 다 — 한 가닥을 직선 하나로 펴되 길이와 연결은 그대로 둔다.
  `exclude-layer` 는 아예 없앴다: 남겨 두면 언젠가 누가 누른다.

  이 파일의 ⑴⑵ 는 그 «접기» 를 지킨다. ⑶⑷(X교차·등각 겹침)는 그대로다.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


# ─────────────────────────────── ⑶ X자 교차 검출
def _find():
    from services.cad_import.convert.planar import _find_x_crossings
    return _find_x_crossings


def test_X자로_만나면_찾는다():
    pos = {1: (0.0, 0.0), 2: (100.0, 0.0), 3: (50.0, -50.0), 4: (50.0, 50.0)}
    got = _find()(pos, {(1, 2), (3, 4)})
    assert len(got) == 1, got
    assert abs(got[0]["at"][0] - 50.0) < 1e-6
    assert abs(got[0]["at"][1] - 0.0) < 1e-6


def test_노드를_공유하면_교차가_아니다():
    """★T·꺾임은 이미 노드로 만나 있다 — 그것까지 세면 온 도면이 «교차» 다.

    처음에 «끝점 좌표가 선 위에 얹힌 것» 으로 시험을 썼다가 걸렸다. 그 판정은
    좌표가 아니라 **노드 번호 공유**다 — 좌표만 얹힌 것은 티 겹침 정규화가
    보는 다른 자리(관통 T)이고, 그쪽은 이미 쪼개고 있다.
    """
    pos = {1: (0.0, 0.0), 2: (100.0, 0.0), 3: (50.0, 50.0)}
    assert _find()(pos, {(1, 2), (2, 3)}) == []


def test_스치기만_하면_교차가_아니다():
    pos = {1: (0.0, 0.0), 2: (100.0, 0.0), 3: (200.0, -50.0), 4: (200.0, 50.0)}
    assert _find()(pos, {(1, 2), (3, 4)}) == []


def test_평행선은_교차가_아니다():
    pos = {1: (0.0, 0.0), 2: (100.0, 0.0), 3: (0.0, 10.0), 4: (100.0, 10.0)}
    assert _find()(pos, {(1, 2), (3, 4)}) == []


def test_자동으로_쪼개지_않는다():
    """★실제 티일 수도, 다른 높이의 배관이 겹쳐 보이는 것일 수도 있다.

    코드가 구분할 수 없으므로 임의로 노드를 만들면 없던 분기가 생긴다.
    검출은 «세기» 이고 간선 집합을 건드리지 않는다.
    """
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "convert" / "planar.py").read_text(encoding="utf-8")
    i = src.index("x_crossings = _find_x_crossings(")
    seg = src[i:i + 700]
    assert "snap_edges =" not in seg, "검출이 간선을 바꾼다"
    assert "자동으로 쪼개지 않는다" in seg
    assert '"x_crossings": x_crossings' in src, "결과를 안 실어 보낸다"


# ─────────────────────────────── ⑷ 등각 겹침 세기
def test_같은_수직선의_접속관_겹침을_센다():
    from services.cad_import.design.sdf_post import _count_screen_crossings
    nodes = [
        {"label": "p1", "x": 0.0, "y": 0.0},
        {"label": "h1", "x": 0.0, "y": 30.0},     # p1 위로 30
        {"label": "p2", "x": 0.0, "y": 10.0},
        {"label": "h2", "x": 0.0, "y": 40.0},     # 구간이 겹친다
    ]
    got = _count_screen_crossings(nodes, {"h1", "h2"},
                                  {"h1": "p1", "h2": "p2"})
    assert got["total"] == 1 and got["stub_vs_pipe"] == 1


def test_다른_수직선이면_안_겹친다():
    from services.cad_import.design.sdf_post import _count_screen_crossings
    nodes = [
        {"label": "p1", "x": 0.0, "y": 0.0},
        {"label": "h1", "x": 0.0, "y": 30.0},
        {"label": "p2", "x": 99.0, "y": 10.0},
        {"label": "h2", "x": 99.0, "y": 40.0},
    ]
    got = _count_screen_crossings(nodes, {"h1", "h2"},
                                  {"h1": "p1", "h2": "p2"})
    assert got["total"] == 0


def test_세기가_좌표를_안_바꾼다():
    """지시서 §8 — bake_isometric 의 좌표 계산식은 한 글자도 건드리지 않는다."""
    from services.cad_import.design.sdf_post import _count_screen_crossings
    nodes = [{"label": "p1", "x": 0.0, "y": 0.0},
             {"label": "h1", "x": 0.0, "y": 30.0}]
    before = [dict(n) for n in nodes]
    _count_screen_crossings(nodes, {"h1"}, {"h1": "p1"})
    assert nodes == before


def test_구운_결과에_겹침_셈이_실린다():
    from services.cad_import.design.sdf_post import bake_isometric

    class _T:
        nodes = [{"label": "1", "x": 0.0, "y": 0.0, "elevation": 0.0},
                 {"label": "2", "x": 100.0, "y": 0.0, "elevation": 0.0},
                 {"label": "3", "x": 100.0, "y": 0.0, "elevation": -0.3}]
    got = bake_isometric(_T(), units_per_m=100.0,
                         head_nodes=["3"], head_parent={"3": "2"})
    assert "crossings" in got and got["crossings"]["total"] == 0


def test_로그가_위상_문제가_아니라고_말한다():
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "design" / "emit.py").read_text(encoding="utf-8")
    # ★첫 «[G15]» 는 헬퍼 docstring 이다 — 로그 문장은 `msg = (f"[G15]` 다.
    i = src.index('msg = (f"[G15]')
    seg = src[i:i + 900]
    assert "위상 문제 아님" in seg, seg[:300]


def test_화면도_같은_말을_한다():
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    i = js.index("function renderIsoNote()")
    seg = js[i:i + 800]
    assert "위상 문제가 아닙니다" in seg
    assert "층고를 입력하면" in seg
    html = (_ROOT / "templates" / "module_f.html").read_text(encoding="utf-8")
    assert 'id="dg-iso-note"' in html


# ─────────────────────────────── ⑴⑵ 레이어 제외
def _client():
    import importlib
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    c = srv.app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def test_자동_사전은_추천만_한다():
    """★도면마다 관례가 다르다 — 사전은 반드시 놓친다. 확정은 사람이 한다."""
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    i = src.index("FLEX_WORDS = (")
    seg = src[max(0, i - 900):i + 400]
    assert "추천만" in seg and "확정은 사람" in seg, seg[-400:]
    # 자동으로 접는 코드가 없어야 한다 — 라우트는 사람이 부를 때만 돈다.
    assert "def module_f_pick_fold_layer" in src
    j = src.index("def module_f_pick_adopt")
    assert "_is_flex" not in src[j:j + 2000], "채택이 몰래 접고 있다"


def test_빼는_길은_없앴다():
    """★빼면 물닿음 헤드가 111 → 5 다(실측). 남겨 두면 언젠가 누가 누른다."""
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    assert "def module_f_pick_exclude_layer" not in src
    assert '"/api/module-f/pick/exclude-layer"' not in src
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    assert "excludeMatLayer" not in js
    assert "pick/exclude-layer" not in js


def test_접은_것을_기록에_남긴다():
    """조용히 접으면 형상이 바뀐 것을 아무도 모른다(S340)."""
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    i = src.index("def module_f_pick_fold_layer")
    seg = src[i:i + 2600]
    assert "pick_layer_folded" in seg, "접기 이력이 세션에 안 남는다"
    assert "print(" in seg, "로그에도 안 남는다"
    # ★찍기 기록에는 넣지 않는다 — 그 목록 항목은 좌표를 가진 «클릭» 이고,
    #   `highlight_geom()` 이 마지막 항목의 x 를 읽는다(실측: 화면 500).
    assert "clicks.append" not in seg, "좌표 없는 항목을 클릭 기록에 끼운다"


def test_화면이_무엇이_일어나는지_말한다():
    """★막는 게 아니라 «무엇이 되는지» 를 알려 준다 — 길이와 연결은 그대로다."""
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    i = js.index("async function foldMatLayer(")
    seg = js[i:i + 1200]
    assert "직선으로 접습니다" in seg, seg[:500]
    assert "길이" in seg and "연결은 그대로" in seg, seg[:500]
    html = (_ROOT / "templates" / "module_f.html").read_text(encoding="utf-8")
    assert "길이와 연결은 그대로" in html, "찍기 화면이 무엇이 되는지 안 말한다"


def test_재료_목록이_레이어별로_묶여_나온다():
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    i = src.index("def module_f_pick_materials")
    seg = src[i:i + 1200]
    for key in ('"layer"', '"colors"', '"segs"', '"flex"'):
        assert key in seg, key


def test_접은_레이어도_목록에_남는다():
    """★한 번 지정하면 그 줄이 사라져 되돌릴 길이 없었다 — 실측으로 막혔다."""
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    i = src.index("def module_f_pick_materials")
    seg = src[i:i + 2200]
    assert "pick_layer_folded" in seg, "접기 이력을 안 본다"
    assert '"fold"' in seg or 'row["fold"]' in seg, "접는지 아닌지를 안 말한다"
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    j = js.index("async function loadPickMaterials()")
    body = js[j:j + 1800]
    assert "r.fold" in body, "화면이 «접기» 상태를 안 그린다"
