# -*- coding: utf-8 -*-
"""[계통도·기계실] 배관 레이어 고르기 화면 — 색으로 보이고, 섞인 것을 말한다.

■ 사용자 지적 둘

  「펌프 찍고 자동으로 알람밸브 추적해 빨간선으로 표시하는 기능은 아주 좋은데,
   캐드도면의 색깔별로 보이게 해서(기계실도 같은 논리로) 구분이 쉽게.」
  「계통도 경로 추출이 그 레이어 안에서 색깔별로 최소 길이가 나와야 되는데,
   뭔가 경로가 꼬여서 추적이 된다.」

■ 그래서 이 화면이 지킬 것

  ⑴ 목록의 색이 **도면에 그려진 그 색**이다(맞대 볼 수 있어야 한다).
  ⑵ 여러 계통을 섞어 뽑고 있으면 그 사실과 대가를 말한다.
  ⑶ 자동으로 한 계통으로 좁혔으면 무엇을·왜 골랐는지 말한다.
  ⑷ 사람이 직접 고를 길이 늘 있다.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

_STUB = r"""
const BOX = {innerHTML: "", nodes: []};
const CALLS = {graph: []};
const S = {sid: "x", subGraph: GRAPH};
const esc = (v) => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const $ = (id) => (id === "sub-layers" ? {
  set innerHTML(v) { BOX.innerHTML = v; },
  get innerHTML() { return BOX.innerHTML; },
  querySelectorAll: () => [],
} : null);
const loadSubGraph = (l) => { CALLS.graph.push(l || null); };
"""


def _js() -> str:
    return open(os.path.join(_ROOT, "static", "module_f.js"),
                encoding="utf-8").read()


def _render(graph):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = _js()
    i = js.index("  function renderSubLayers()")
    src = js[i:js.index("\n  }\n", i) + 4]
    prog = "\n".join([f"const GRAPH = {json.dumps(graph)};", _STUB, src,
                      "renderSubLayers();",
                      "console.log(JSON.stringify({html: BOX.innerHTML,"
                      " calls: CALLS}));"])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-900:]
    return json.loads(out.stdout)


def _g(**kw):
    base = {"nodes": [[0, 0], [1, 1]], "edges": [[0, 1, 100, 0]],
            "forced": 0, "components": 1, "auto_layers": [],
            "chosen": None, "chosen_auto": False, "narrowed": None,
            "layers": [{"layer": "HSP", "n": 40, "cat": "PIPE",
                        "color": 1, "css": "#c04040"},
                       {"layer": "LSP", "n": 60, "cat": "PIPE",
                        "color": 3, "css": "#40c040"}]}
    base.update(kw)
    return base


def test_레이어마다_도면_색을_찍는다():
    got = _render(_g())
    assert "#c04040" in got["html"] and "#40c040" in got["html"], got["html"]
    assert got["html"].count('class="sw"') == 2


def test_색이_없는_레이어도_목록에서_빠지지_않는다():
    """색을 못 읽었다고 고를 거리를 지우면 사람이 손을 못 댄다."""
    got = _render(_g(layers=[{"layer": "SP", "n": 9, "cat": "OTHER"}]))
    assert "SP" in got["html"]
    assert 'class="sw"' not in got["html"]


def test_여러_계통을_섞고_있으면_그_사실을_말한다():
    """★«꼬임» 의 정체다 — 고층·저층이 한 그래프에 있으면 경로가 오간다."""
    got = _render(_g(auto_layers=["HSP", "LSP"]))
    assert "섞어" in got["html"], got["html"]
    assert "HSP" in got["html"] and "LSP" in got["html"]


def test_하나만_고르는_단추를_준다():
    got = _render(_g(auto_layers=["HSP", "LSP"]))
    assert got["html"].count("data-only=") == 2


def test_좁히려다_못_좁혔으면_그_사유도_말한다():
    """★실도면에서 자주 나는 쪽이다 — 두 점이 서로 다른 계통에 붙을 때.

    그때 조용히 섞은 채로 뽑으면 경로가 계통 사이를 오간다. 사람이 두 점을
    다시 찍을지 레이어를 직접 고를지 정할 수 있게, 사유를 그 자리에 적는다.
    """
    got = _render(_g(auto_layers=["HSP", "LSP"], chosen=None,
                     chosen_auto=False,
                     narrowed={"layer": None,
                               "reason": "찍은 두 점이 서로 다른 계통에 붙습니다"}))
    assert "스스로 좁히지 못했습니다" in got["html"], got["html"]
    assert "서로 다른 계통" in got["html"]


def test_한_장뿐이면_섞였다고_겁주지_않는다():
    got = _render(_g(auto_layers=["HSP"]))
    assert "섞어" not in got["html"]


def test_자동으로_좁혔으면_무엇을_왜_골랐는지_말한다():
    got = _render(_g(chosen=["HSP"], chosen_auto=True,
                     auto_layers=["HSP", "LSP"],
                     narrowed={"layer": "HSP",
                               "reason": "찍은 두 점이 «HSP» 배관에 가장 붙어 있습니다"}))
    assert "HSP" in got["html"]
    assert "가장 붙어" in got["html"], got["html"]
    # 사람이 뒤집을 수 있다는 것도 말해야 한다.
    assert "직접 고르면" in got["html"]


def test_좁힌_뒤에는_섞였다는_경고를_겹쳐_띄우지_않는다():
    got = _render(_g(chosen=["HSP"], chosen_auto=True,
                     auto_layers=["HSP", "LSP"], narrowed={"layer": "HSP"}))
    assert "섞어" not in got["html"]


def test_자동으로_고른_레이어에는_표가_붙는다():
    got = _render(_g(auto_layers=["HSP"]))
    assert "자동" in got["html"]


def test_추측_연결과_조각을_숨기지_않는다():
    got = _render(_g(forced=7, components=3))
    assert "추측 연결 7" in got["html"] and "조각 3" in got["html"]


def test_그래프가_없으면_조용히_비운다():
    got = _render(None)
    assert got["html"] == ""


# ─────────────────────────────── 두 점을 서버로 보내는가
def test_두_점을_찍으면_그래프를_다시_받는다():
    """서버가 «그 두 점이 있는 계통» 으로 좁히려면 점을 받아야 한다."""
    js = _js()
    i = js.index("  async function loadSubGraph(layers)")
    src = js[i:js.index("\n  }\n", i) + 4]
    assert "body.a" in src and "body.b" in src, src
    i2 = js.index("  function subClick(x, y)")
    src2 = js[i2:js.index("\n  }\n", i2) + 4]
    assert "narrowSubLayer" in src2, "두 점을 찍어도 다시 안 받는다"


def test_그래프를_갈아_끼우면_옛_미리보기_경로를_버린다():
    """★실측으로 콘솔 오류 2건이 여기서 났다.

    미리보기 경로는 그래프의 «절점 번호» 열이다. 좁히면 절점이 255 → 76 으로
    줄어드는데, 옛 번호를 그대로 두면 없는 절점을 그리다 터진다. 그리기 한
    줄이 죽으면 도면이 통째로 사라진다.
    """
    js = _js()
    i = js.index("  async function loadSubGraph(layers)")
    src = js[i:js.index("\n  }\n", i) + 4]
    assert "S.sub.preview = null" in src, src
    j = js.index("  function drawSubPreview()")
    draw = js[j:js.index("\n  }\n", j) + 4]
    assert "if (!a || !b) continue;" in draw, "없는 절점을 그대로 그린다"


def test_좁힌_결과를_화면_상태에_담는다():
    """★골라 담는 자리가 늘 사고를 낸다(forced_penalty_mm 를 그렇게 잃었다)."""
    js = _js()
    i = js.index("  async function loadSubGraph(layers)")
    src = js[i:js.index("\n  }\n", i) + 4]
    for key in ("auto_layers", "chosen_auto", "narrowed", "chosen"):
        assert key in src, f"{key} 를 안 담으면 화면이 그것을 말할 수 없다"
