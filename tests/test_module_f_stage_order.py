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
def test_파일이_나는_자리는_하나다():
    """같은 파일(emit_design_files)을 두 자리에서 내지 않는다.

    수리계산에 있던 「.sdf+.slf 저장 · 내려받기」는 걷어냈다 — 자동 경로에도
    남기지 않는다(그 표는 방출기가 받지 못한다 · `_probe_auto_emit.py`).
    """
    html, js = _html(), _js()
    for gone in ('id="dg-emit"', 'id="dg-download"'):
        assert gone not in html, f"{gone} 가 화면에 남아 있다"
    assert 'dg-emit' not in js and 'dg-download' not in js
    # 파일을 내는 자리로 가는 문은 수리계산의 «수리계산 입력 변환 →» 하나다.
    assert html.count('id="dg-to-conv"') == 1
    # 주석에 이름이 남는 것은 괜찮다 — 막을 것은 «부르는» 자리다.
    assert 'post("/api/module-f/design/emit"' not in js, \
        "화면이 방출 라우트를 직접 부르는 곳이 남았다"


def test_자동은_짧은_지선이다():
    """★사용자 반려 [2026-09-08]: 「두 길의 화면 차례를 같게 만들면 지선이
    본선과 대등해 보인다. D-F10-2 가 지선을 접힌 «고급» 으로 밀어넣은 방향과
    반대이고, 직렬 지향과도 어긋난다.」

    한 번 그렇게 폈다가 되돌린 자리다 — 자동 단계바를 본선과 같은 다섯 칸으로
    세웠었다. 지선은 짧게 두고, 끝내려면 «손질로 이어받기» 로 합류한다.
    """
    js = _js()
    i = js.index("const STAGE_FLOW = {")
    src = js[i:js.index("};", i)]
    auto = re.search(r'plan_auto:\s*\[([^\]]+)\]', src).group(1)
    got = [v.strip().strip('"') for v in auto.split(",")]
    assert got == ["open", "auto"], got
    plan = re.search(r'plan:\s*\[([^\]]+)\]', src).group(1)
    manual = [v.strip().strip('"') for v in plan.split(",")]
    assert len(got) < len(manual), (got, manual)
    assert got[-1:] != manual[-1:], "지선 끝이 본선 끝과 같으면 대등해 보인다"


def test_자동_화면에서_나가는_문은_이어받기_하나다():
    """그 화면은 A 결과를 보는 자리다 — 본선 화면으로 건너뛰지 않는다."""
    html, js = _html(), _js()
    assert "au-to-design" not in html and "au-to-design" not in js, \
        "지선에서 본선 화면으로 건너뛰는 단추가 남아 있다"
    assert 'id="au-handoff"' in html
    assert 'id="dg-emit"' not in html and 'id="dg-download"' not in html


def test_자동_화면이_다음_걸음을_알려준다():
    """막는 것이 아니라 «파일을 내려면 이어받기» 라고 알려 주는 문구다."""
    html = _html()
    i = html.index('id="au-handoff"')
    seg = html[i:i + 600]
    assert "파일을 내려면" in seg and "이어받기" in seg, seg[:300]


def test_이어받을_때_기준개수를_들고_간다():
    """자동에서 20 을 골라 뽑고 이어받았는데 30 으로 되돌아가면 안 된다."""
    js = _js()
    i = js.index('$("au-handoff").onclick')
    src = js[i:js.index("\n  };\n", i)]
    assert '"au-k"' in src and '"ed-k"' in src, src


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


# ─────────────────────────────── ⑹ 수리계산 화면은 «평면부터»
def test_수리계산에_들어오면_평면부터_보인다():
    """★사용자 지적: 「손질까지 끝내고 수리계산 → 을 누르니 화면에 아무것도
    안 나온다. 원래면 아이소매트릭 망이 나와야 되는 거 아니야?」

    그럴 수밖에 없었다 — 이 화면이 그리는 것은 «설계 좌표» 이고 그 좌표는
    「표 확정」이 만든다. 확정 전에는 그릴 것이 없어 캔버스가 검게 빈다.
    빈 화면은 «고장» 으로 읽히므로, 들어올 때는 손질한 평면을 보여 준다.
    """
    html, js = _html(), _js()
    i = html.index('id="dg-plan"')
    assert "checked" in html[i:i + 60], "평면 보기가 기본이 아니다"
    src = _fn(js, "  async function enterDesign()")
    assert 'edit/state' in src, "손질 상태 없이 평면을 그릴 수 없다"
    assert '$("dg-plan").checked = true' in src, src


def test_보기_때문에_저장되는_좌표를_바꾸지_않는다():
    """★`dg-iso` 는 화면 전환이 아니라 **투영 설정** 이다.

    `_view_opts` 를 거쳐 `emit_design_sdf` 로 간다 — 보기 편하자고 끄면
    저장되는 .sdf 좌표가 조용히 바뀐다. 화면을 가르는 스위치는 «평면에서
    보기» 하나여야 한다.
    """
    js = _js()
    src = _fn(js, "  async function enterDesign()")
    assert '$("dg-iso").checked =' not in src, "보기 때문에 투영 설정을 건드린다"
    py = open(os.path.join(_ROOT, "routes", "module_f", "api_design.py"),
              encoding="utf-8").read()
    i = py.index("def _view_opts(")
    assert '"iso"' in py[i:i + 400], "투영 설정이 산출로 간다는 전제가 깨졌다"


def test_들어올_때_시점을_평면에_맞춘다():
    """미리보기가 설계 좌표로 시점을 끌고 가면 도면이 사라진 것처럼 보인다."""
    js = _js()
    src = _fn(js, "  async function designPreview()")
    assert "fitDesignView()" in src, src
    assert "const xs = d.view.nodes.map" not in src, "설계 좌표로 시점을 끈다"


def test_평면_보기_중_클릭은_고르지_않으면_읽기다():
    """★평면 보기가 기본이 된 뒤로는 손질에서 쓰던 모드가 남아 있을 수 있다.

    그때 무심코 찍으면 알람밸브가 놓인다 — 보기 화면에서 일어나면 안 된다.
    «이음·삭제» 를 고른 동안에만 고친다.
    """
    js = _js()
    i = js.index('S.stage === "design" && planUnderlayOn()')
    seg = js[i:i + 300]
    assert '"이음"' in seg and '"삭제"' in seg, seg
