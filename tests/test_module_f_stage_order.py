# -*- coding: utf-8 -*-
"""[회로 순서] 재료가 먼저 나고, 그 뒤에 파일이 나온다.

■ 사용자 지적

  「4 변환에 필요한 변환 실행 파트를 위한 키가 5 수리계산에 있는 게 순서가
  좀 안 맞지 않아? 회로 순서에 대대적인 조치가 필요하겠는데.
  그리고 6.통합(및 계통도·기계실의 3.통합)은 어차피 같은 목적지니까 따로
  빼서 우측 상단에 하나의 버튼으로 통일.」

■ 무엇이 어긋나 있었나 — 서버 코드가 증거였다

  `convert/run`(4단계)이 최불리 산출물을 요구받으면 이렇게 답했다:
  「수리계산 패널에서 「표 확정」을 먼저 눌러 주세요」 — **뒤 단계(5)의
  산출을 앞 단계(4)가 요구**하는 회로다. 게다가

    · 기준개수 K 입력칸이 손질(ed-k)과 수리계산(dg-k) **두 곳**에 있었고,
      코드가 최불리를 뽑을 때마다 값을 몰래 맞춰 주고 있었다. 맞춰 줘야
      한다는 것 자체가 키가 잘못된 자리에 있다는 증거다.
    · 「.sdf+.slf 저장」이 수리계산과 변환 두 곳에 있었는데 **같은 함수**
      (`emit_design_files`)를 불렀다 — 같은 파일이 두 자리에서 났다.
    · 통합이 슬롯마다 단계바 끝에 붙어 세 번 되풀이됐다.

■ 여기서 못박는 것

  ⑴ 평면도 흐름은 열기 → 찍기 → 손질 → **수리계산** → **변환** 이다.
  ⑵ 통합은 단계바에 없다 — 머리말 단추 하나가 그 자리다.
  ⑶ K 입력칸은 하나뿐이다(앞 단계). 화면은 그 값을 «보여만» 준다.
  ⑷ 뒤로 가는 단추는 바로 앞 단계를 가리킨다.
  ⑸ 재료가 없으면 잠그지 말고 말한다.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _js() -> str:
    return open(os.path.join(_ROOT, "static", "module_f.js"),
                encoding="utf-8").read()


def _html() -> str:
    return open(os.path.join(_ROOT, "templates", "module_f.html"),
                encoding="utf-8").read()


def _fn(js: str, head: str) -> str:
    i = js.index(head)
    return js[i:js.index("\n  }\n", i) + 4]


# ─────────────────────────────── ⑴ 순서
def test_평면도는_수리계산_다음에_변환이다():
    js = _js()
    i = js.index("const STAGE_FLOW = {")
    src = js[i:js.index("};", i)]
    plan = re.search(r'plan:\s*\[([^\]]+)\]', src).group(1)
    got = [v.strip().strip('"') for v in plan.split(",")]
    assert got == ["open", "pick", "edit", "design", "conv"], got


def test_손질의_다음_단추는_수리계산으로_간다():
    """★변환으로 곧장 보내면 재료가 없는 화면에 떨어진다."""
    js = _js()
    i = js.index('$("ed-next").onclick')
    src = js[i:js.index(";\n", i)]
    assert '"design"' in src, src
    assert '"conv"' not in src, src


# ─────────────────────────────── ⑵ 통합 = 머리말 단추 하나
def test_통합은_단계바에_붙지_않는다():
    js = _js()
    src = _fn(js, "  function stageFlow()")
    assert 'concat(["merge"])' not in src, "슬롯마다 통합이 되풀이된다"


def test_머리말에_통합_단추가_하나_있다():
    html = _html()
    assert html.count('id="btn-merge"') == 1
    head = html[html.index("<header>"):html.index("</header>")]
    assert 'id="btn-merge"' in head, "머리말(우측 상단)에 있어야 한다"


def test_통합_단추는_잠그지_않고_말한다():
    """잠긴 채 침묵하면 «고장» 으로 읽힌다 — 결합 단추에서 이미 겪었다."""
    js = _js()
    i = js.index('$("btn-merge").onclick')
    src = js[i:js.index("\n  };\n", i)]
    assert "disabled" not in src, src
    assert "say(" in src and "loadMerge" in src


def test_지금_통합에_있으면_단추가_그것을_보인다():
    js = _js()
    src = _fn(js, "  function renderMergeTab()")
    assert '"merge"' in src and "classList.toggle" in src


# ─────────────────────────────── ⑶ K 는 한 곳에서만
def test_수리계산에는_K_입력칸이_없다():
    """★칸이 둘이면 손질에서 본 30개와 표에 실린 K 가 갈린다."""
    html = _html()
    assert 'id="dg-k"' not in html, "수리계산에 K 입력칸이 남아 있다"
    assert 'id="ed-k"' in html, "손질의 K 칸까지 사라지면 정할 자리가 없다"
    assert 'id="dg-k-note"' in html, "무엇으로 도는지 보여 줄 자리가 없다"


def test_설정의_K_는_앞_단계_칸에서_온다():
    js = _js()
    src = _fn(js, "  function designK()")
    assert '"au-k"' in src and '"ed-k"' in src, src
    st = _fn(js, "  function designSettings()")
    assert "k: designK()" in st, st


def test_다시_계산도_같은_K_를_쓴다():
    js = _js()
    i = js.index('$("dg-recalc").onclick')
    src = js[i:js.index("\n  };\n", i)]
    assert "designK()" in src, src


def test_서버도_최불리_K_를_설계설정에_맞춘다():
    """화면이 옛 클라이언트여도 두 곳이 다른 말을 못 하게."""
    py = open(os.path.join(_ROOT, "routes", "module_f", "api_edit.py"),
              encoding="utf-8").read()
    i = py.index('sess["worst_k"] = k')
    seg = py[i:i + 700]
    assert 'design_settings' in seg and 'ds["k"] = k' in seg, seg[:300]


# ─────────────────────────────── ⑷ 앞뒤 단추
def test_뒤로_가는_단추는_바로_앞_단계를_가리킨다():
    js = _js()
    i = js.index('$("dg-back").onclick')
    assert '"edit"' in js[i:js.index(";\n", i)], "수리계산의 «뒤로» 가 손질이 아니다"
    j = js.index('$("btn-back-design").onclick')
    assert '"design"' in js[j:js.index(";\n", j)], "변환의 «뒤로» 가 수리계산이 아니다"
    html = _html()
    assert 'id="btn-back-design">← 수리계산' in html
    assert 'id="dg-back">← 손질' in html
    assert 'id="dg-to-conv">수리계산 입력 변환 →' in html


def test_5단계_이름은_무엇을_무엇으로_옮기는지_말한다():
    """「변환」만으로는 무엇을 무엇으로 옮기는지가 빠진다(사용자 지적)."""
    js = _js()
    i = js.index("const STAGE_LABEL = {")
    src = js[i:js.index("};", i)]
    assert 'conv: "수리계산 입력 변환"' in src, src


def test_옛_이름의_단추는_남아_있지_않다():
    """옛 단추가 남으면 사람이 옛 회로로 되돌아간다."""
    html, js = _html(), _js()
    for old in ("btn-to-design", "dg-to-merge", "btn-back-edit"):
        assert old not in html, f"{old} 가 화면에 남아 있다"
        assert old not in js, f"{old} 가 코드에 남아 있다"


# ─────────────────────────────── ⑸ 파일은 한 자리에서
def test_수동_경로의_저장은_변환_한_곳에서만():
    """같은 파일(emit_design_files)을 두 자리에서 내지 않는다."""
    js = _js()
    src = _fn(js, "  function syncDesignForMethod()")
    assert '"dg-emit-row"' in src and "!auto" in src, src
    assert '"dg-to-conv"' in src, src


def test_변환은_없는_재료를_이름으로_말한다():
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = _js()
    src = "\n".join([_fn(js, "  function convMissing()"),
                     _fn(js, "  function renderConvWhy()")])
    stub = """
const BOX = {innerHTML: "", cls: {}};
const S = {edit: EDIT, design: DESIGN};
const CHK = CHECKED;
const $ = (id) => (id === "cv-why" ? {
  set innerHTML(v) { BOX.innerHTML = v; },
  get innerHTML() { return BOX.innerHTML; },
  classList: {toggle(k, on) { BOX.cls[k] = !!on; }},
} : {checked: !!CHK[id]});
"""

    def run(edit, design, checked):
        prog = "\n".join([
            f"const EDIT = {json.dumps(edit)};",
            f"const DESIGN = {json.dumps(design)};",
            f"const CHECKED = {json.dumps(checked)};",
            stub, src, "renderConvWhy();",
            "console.log(JSON.stringify({miss: convMissing(), box: BOX}));"])
        out = subprocess.run([node, "-e", prog], capture_output=True,
                             text=True, encoding="utf-8", errors="replace")
        assert out.returncode == 0, out.stderr[-800:]
        return json.loads(out.stdout)

    on = {"cv-full-kfp": True, "cv-worst-kfp": True, "cv-worst-sdf": True}
    got = run({"worst": None}, None, on)
    assert len(got["miss"]) == 2, got
    assert got["box"]["cls"]["warn"] is True

    got = run({"worst": {"k": 30}}, {"tables": {}}, on)
    assert got["miss"] == [], got
    assert "모두 있습니다" in got["box"]["innerHTML"]

    # ★표가 «없어도» S.design 은 채워진다(미리보기가 view:null 로 200 을 준다).
    #   그래서 재료 여부는 S.design 이 아니라 **표** 로 본다.
    got = run({"worst": {"k": 30}}, {"view": None, "tables": None}, on)
    assert any("표 확정" in m for m in got["miss"]), got

    # 안 고른 산출물 때문에 겁주지 않는다.
    got = run({"worst": None}, None, {"cv-full-kfp": True})
    assert got["miss"] == [], got


def test_변환_실행_전에도_같은_말을_한다():
    js = _js()
    i = js.index('$("btn-convert").onclick')
    src = js[i:js.index("\n  };\n", i)]
    assert "convMissing()" in src, "재료를 안 보고 서버로 보낸다"
    assert src.index("convMissing()") < src.index("convert/run"), src[:200]
