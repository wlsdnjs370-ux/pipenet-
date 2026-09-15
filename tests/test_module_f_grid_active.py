# -*- coding: utf-8 -*-
"""[통합 격자·활성·위상 §4 M1·M2] 격자와 활성 표시 — 화면 규약을 못박는다.

지시서 `ModuleF_통합_격자활성_위상수정_지시서.md` §3-1 · §3-2 · D1~D4.

화면 코드라 «보기 좋은가» 는 사람이 본다. 여기서 잡는 것은 **말이 되는가** 다:

  M1 격자 — 한 칸이 정확히 1 m 인가 · 줌아웃하면 단계가 올라가는가(D1) ·
            원점이 기준점(통합)/접속점(04)을 지나는가(D2)
  M2 활성 — 고른 것만 1.0 이고 나머지는 0.35 인가 · 흐리기가 **곱해지는가**
            (캔버스 알파는 덮어쓰기라, 제 알파를 세우는 자리가 겉을 지운다)
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_JS = os.path.join(_ROOT, "static", "module_f.js")


def _src():
    return open(_JS, encoding="utf-8").read()


def _fn(head):
    """`  function 이름(` 부터 그 함수 끝(`\\n  }\\n`)까지."""
    js = _src()
    i = js.index(head)
    return js[i:js.index("\n  }\n", i) + 4]


def _node(prog):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-900:]
    return json.loads(out.stdout)


# ─────────────────────────────────────────────── M1 격자
_GRID_STUB = """
const S = {view: {scale: SCALE}};
const GRID_STEPS = [1, 5, 25];
const GRID_MIN_PX = 8;
"""


def _step(scale_per_m, view_scale):
    prog = "\n".join([
        f"const SCALE = {view_scale};", _GRID_STUB,
        _fn("  function gridStepM("),
        f"console.log(JSON.stringify(gridStepM({scale_per_m})));"])
    return _node(prog)


def test_한_칸이_8px_보다_촘촘하면_다음_단계로():
    """D1 — 고정 1 m 만 두면 줌아웃했을 때 격자가 화면을 회색으로 덮는다."""
    # 1 m = 1000 단위 · 화면 배율 0.01 → 한 칸 10 px → 1 m 그대로
    assert _step(1000, 0.01) == 1
    # 배율을 1/4 로 → 한 칸 2.5 px → 5 m 단계(12.5 px)
    assert _step(1000, 0.0025) == 5
    # 더 줄이면 25 m
    assert _step(1000, 0.0004) == 25
    # 25 m 조차 8 px 을 못 채우면 **안 그린다** — 회색으로 덮지 않는다
    assert _step(1000, 0.00001) is None


def test_배율이_없으면_격자를_안_깐다():
    """지어내지 않는다 — 1 m 가 몇 단위인지 모르면 자가 아니다."""
    assert _step(0, 1.0) is None
    assert _step(-5, 1.0) is None


def test_통합의_한_칸은_1000_단위이고_원점은_기준점이다():
    """D2 · 1-B ① — 결합 좌표는 mm 라 유추할 것이 없다."""
    src = _fn("  function gridSpecFor(")
    assert "s: 1000" in src, src
    assert "n.anchor" in src, src          # 교점 하나가 기준점을 지난다


def test_04_의_한_칸은_1000곱하기_underlay_k_다():
    """§3-1-2 — 04 는 정규화 좌표라 1 m 가 1000·k 단위다."""
    src = _fn("  function gridSpecFor(")
    assert "1000 * u.k" in src, src
    assert "n.input" in src, src          # 원점은 접속점(Input)


def test_아이소를_눕혀도_한_변은_1m_다():
    """1-B ⑤ — (1,0)·(0,1) 을 30° 등각으로 보내도 길이가 1 이다."""
    c, s = 0.8660254037844387, 0.5
    for vx, vy in ((1.0, 0.0), (0.0, 1.0)):
        x, y = (vx - vy) * c, (vx + vy) * s
        assert (x * x + y * y) ** 0.5 == pytest.approx(1.0, abs=1e-12)


def test_격자는_망보다_먼저_그린다():
    """뒤에 그리면 배관을 덮는다 — 순서가 곧 규약이다."""
    js = _src()
    i_grid = js.index('paintGrid("design"')
    i_net = js.index("withDim(drawDesign)")
    assert i_grid < i_net, (i_grid, i_net)
    j_grid = js.index('paintGrid("merge"')
    j_net = js.index("withDim(drawMerged)")
    assert j_grid < j_net, (j_grid, j_net)


# ─────────────────────────────────────────────── M2 활성 표시
def test_흐리기는_덮어쓰지_않고_곱한다():
    """★캔버스의 `globalAlpha` 는 대입이다.

    결합의 평면 배관은 담당 헤드 수로 제 알파(0.55~1.0)를 세운다. 그 자리에서
    그냥 대입하면 겉의 0.35 가 **통째로 지워져** 「나머지를 흐리게」가 안 먹는다.
    그래서 `dimA()` 한 자리로 모은다.
    """
    js = _src()
    assert "function dimA(a) { ctx.globalAlpha = a * dimK; }" in js
    # 결합의 두 자리가 그 자를 쓴다
    assert "dimA(Math.min(1, 0.55 + 0.45 * t));" in js
    assert "\n    dimA(1);\n" in js


def test_고른_것은_빨강_1_0_이고_나머지는_0_35_다():
    js = _src()
    assert 'const SEL_RED = "#ff3b3b";' in js
    assert "const SEL_DIM = 0.35;" in js
    sel = _fn("  function paintSel(")
    assert "ctx.globalAlpha = 1;" in sel          # 고른 것은 온전히
    assert "SEL_RED" in sel
    assert "12, 0, Math.PI * 2" in sel            # 노드는 반지름 12 고리


def test_두_화면이_같은_규약을_부른다():
    """§5 금지 — 선택 강조를 화면마다 각자 쓰지 않는다."""
    js = _src()
    for fn in ("function drawInspectSel(", "function drawMergeSel("):
        i = js.index(fn)
        body = js[i:js.index("\n  }\n", i)]
        assert "paintSel(" in body, fn


def test_표에서_켠_주황_강조는_흐리기에서_뺀다():
    """D4 — 그것도 «지금 보는 것» 이다."""
    body = _fn("  function drawDesign(")
    assert "ctx.globalAlpha = hot ? 1 : dimK;" in body, body[:400]


def test_카드를_닫으면_원상으로_돌아온다():
    """§3-2 3 — 선택이 풀리면 `hasSelection()` 이 거짓이라 흐리기도 없다."""
    body = _fn("  function insClose(")
    assert "S.design.sel = null" in body
    assert "S.mergeSel = null" in body
    assert "S.opArm = null" in body              # 무장도 함께 풀린다


# ─────────────────────────────────────────────── §3-4 밑그림
def test_밑그림은_한_함수를_두_화면이_쓴다():
    """지시서 §3-4 — 「변환을 인자로 받도록 한 줄만 바꾼다」."""
    js = _src()
    assert "function drawUnderlay(u) {" in js
    assert "drawUnderlay(((S.design || {}).view || {}).underlay)" in js
    assert "drawUnderlay((S.mergeView || {}).underlay)" in js


def test_통합_밑그림_변환은_서버가_만든다():
    """§5 금지 — 좌표식을 화면에서 새로 만들지 않는다."""
    py = open(os.path.join(_ROOT, "routes", "module_f", "api_merge.py"),
              encoding="utf-8").read()
    assert "def _merge_underlay(" in py
    assert '"underlay": under' in py
    # 1-B ④ — 결합 프레임은 정규화가 없다. k 는 1 이다.
    m = re.search(r'return \{\n\s+"k": ([0-9.]+),', py)
    assert m and float(m.group(1)) == 1.0, py[py.index("def _merge_underlay"):][:900]
    # 평행이동은 board 원점을 표 원점으로 — `xf_mm_to_m` 의 mm 판.
    assert '"tx": 1000.0 - float(origin[0])' in py
    assert '"ty": 1000.0 - float(origin[1])' in py


def test_재료가_없으면_밑그림을_안_깐다():
    """어림값으로 깔면 「그럴듯하게 어긋난 그림」이 된다(F-10e)."""
    body = _fn("  function drawUnderlay(")
    assert "if (!u || !S.edit || !S.edit.body_groups) return;" in body
    js = _src()
    assert "밑그림 변환을 받지 못했습니다 — 결합을 다시 해 주세요." in js


# ─────────────────────────────────────────────── §3-5 카드 단추
def test_단추는_템플릿에_있다():
    """지어 넣으면 「JS 가 찾는 id 는 템플릿에 있다」는 그물을 빠져나간다."""
    html = open(os.path.join(_ROOT, "templates", "module_f.html"),
                encoding="utf-8").read()
    for wid in ("dg-ins-ops", "dg-ins-noops", "op-why", "op-node",
                "op-equip", "op-lib", "op-del", "op-hint"):
        assert f'id="{wid}"' in html, wid


def test_헤드를_지우기_전에_K_가_준다고_먼저_말한다():
    """D6 — 지운 뒤에 알면 늦다."""
    body = _fn("  function renderInsOps(")
    assert "기준개수 K" in body, body[-700:]
    assert "designK()" in body


def test_사유가_없으면_보내지_않는다():
    """왜 고쳤는지가 안 남으면 그 계산서를 다음 사람이 못 믿는다."""
    body = _fn("  async function opPost(")
    assert "사유를 적어 주세요" in body
    assert 'if (!why) {' in body


def test_회랑을_고치면_두_배너를_함께_세운다():
    """§3-3-1 — 「표 확정」 → 「결합」 둘 다 눌러야 두 산출이 같은 말을 한다."""
    body = _fn("  function markOpsDirty(")
    assert '"dg-stale"' in body and '"mg-stale"' in body
    assert "다시 결합하세요" in body


def test_손질_화면에도_격자를_두되_기본은_끔이다():
    """D5 — 거기는 DXF 도면이 통째로 깔려 있어 격자가 도면선과 섞인다."""
    html = open(os.path.join(_ROOT, "templates", "module_f.html"),
                encoding="utf-8").read()
    i = html.index('id="ed-gridline"')
    # 앞뒤 한 줄 안에 `checked` 가 없어야 «기본 끔» 이다.
    around = html[max(0, i - 200):i + 120]
    assert "checked" not in around.split("ed-gridline")[1][:80], around
    src = _fn("  function gridSpecFor(")
    assert 'stage === "edit"' in src, src
    assert "손질 좌표(board mm)" in src
