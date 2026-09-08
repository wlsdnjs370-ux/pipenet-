# -*- coding: utf-8 -*-
"""[통합] 결합 단추는 **잠그지 않고 말한다**.

■ 사용자 지적

  「평면도·계통도·기계실 도면을 전부 통합하고 가압송수방식을 고르고 결합을
  하면 아이소메트릭이 나와야 되는데, 애초에 결합 버튼이 활성화 되지 않고
  있어.」

■ 재현으로 잡은 것

  전 공정을 화면에서 태워 보니(`scripts/_verify_module_f_pipeline.py`) 결합은
  **돈다** — 절점 308 · 배관 307. 단추가 안 켜지던 것은 재료(평면도 수리계산
  표·급수방식)가 덜 갖춰진 상태였고, 그때 화면이 하는 일이 «disabled» 하나
  뿐이라 사람에게는 «고장» 으로 보였다.

  같은 실수를 최불리 단추에서 이미 한 번 했다. 규칙은 하나다:
  **잠그고 침묵하지 말고, 켜 두고 왜 안 되는지 말한다.**

■ 여기서 지키는 셋

  ⑴ 무엇이 남았는지 **이름으로** 말한다.
  ⑵ 단추를 `disabled` 로 잠그지 않는다.
  ⑶ 눌렀을 때도 같은 말을 한다(빈 요청을 서버로 보내지 않는다).
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

_STUB = r"""
const BOX = {innerHTML: "", cls: {}};
const CALLS = {say: [], post: []};
const S = {sid: "x", merge: MERGE};
const $ = (id) => (id === "mg-why" ? {
  set innerHTML(v) { BOX.innerHTML = v; },
  get innerHTML() { return BOX.innerHTML; },
  classList: {toggle(k, on) { BOX.cls[k] = !!on; }},
} : {textContent: "", innerHTML: "", classList: {toggle(){}}});
const say = (m, k) => CALLS.say.push([String(m), k || null]);
const post = async (p, b) => { CALLS.post.push([p, b]); return {ok: true}; };
"""


def _js() -> str:
    return open(os.path.join(_ROOT, "static", "module_f.js"),
                encoding="utf-8").read()


def _fn(js: str, head: str) -> str:
    i = js.index(head)
    return js[i:js.index("\n  }\n", i) + 4]


def _run(merge, body):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = _js()
    src = "\n".join([_fn(js, "  function mergeMissing(d)"),
                     _fn(js, "  function renderMergeWhy(d)")])
    prog = "\n".join([f"const MERGE = {json.dumps(merge)};", _STUB, src, body])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-900:]
    return json.loads(out.stdout)


_FULL = {"ready": {"plan": True, "system": True, "machineroom": True},
         "mode": "pump", "can_build": True}


def test_다_갖추면_남은_것이_없다():
    got = _run(_FULL, """renderMergeWhy(S.merge);
console.log(JSON.stringify({miss: mergeMissing(S.merge), box: BOX}));""")
    assert got["miss"] == []
    assert "준비" in got["box"]["innerHTML"]
    assert got["box"]["cls"]["warn"] is False


def test_급수방식이_없으면_그것을_이름으로_말한다():
    got = _run({**_FULL, "mode": None, "can_build": False},
               """renderMergeWhy(S.merge);
console.log(JSON.stringify({miss: mergeMissing(S.merge), box: BOX}));""")
    assert got["miss"] == ["급수방식 고르기"]
    assert "급수방식" in got["box"]["innerHTML"]
    assert got["box"]["cls"]["warn"] is True


def test_평면도_표가_없으면_그것을_이름으로_말한다():
    got = _run({"ready": {"plan": False}, "mode": "pump"},
               """console.log(JSON.stringify(mergeMissing(S.merge)));""")
    assert got == ["평면도의 «수리계산 → 표 확정»"]


def test_둘_다_없으면_둘_다_말한다():
    got = _run({"ready": {"plan": False}, "mode": None},
               """console.log(JSON.stringify(mergeMissing(S.merge)));""")
    assert len(got) == 2


def test_상태를_아직_못_받았어도_터지지_않는다():
    """열자마자 그리는 자리다 — null 하나로 화면이 죽으면 안 된다."""
    got = _run(None, """renderMergeWhy(S.merge);
console.log(JSON.stringify({miss: mergeMissing(S.merge), box: BOX}));""")
    assert len(got["miss"]) == 2


# ─────────────────────────────── 소스 규약 — 잠그지 않는다
def test_결합_단추를_disabled_로_잠그지_않는다():
    """★최불리에서 한 번 한 실수다 — 잠긴 채 침묵하면 «고장» 으로 읽힌다."""
    js = _js()
    for ln in js.splitlines():
        if "mg-build" in ln and "disabled" in ln:
            raise AssertionError(f"결합 단추를 잠그는 줄이 남았다: {ln.strip()}")


def test_누르면_같은_말을_하고_서버로_안_보낸다():
    """단추를 눌렀을 때도 «왜 안 되는지» 를 말한다 — 조용히 실패하지 않는다."""
    js = _js()
    i = js.index('$("mg-build").onclick')
    src = js[i:js.index("\n  };\n", i)]
    assert "mergeMissing" in src, "누른 자리에서 무엇이 남았는지 안 본다"
    assert "renderMergeWhy" in src
    assert src.index("mergeMissing") < src.index("post("), \
        "재료를 보기 전에 서버로 보낸다"


def test_서버의_can_build_와_같은_조건을_본다():
    """두 곳이 다른 말을 하면 «켜졌는데 400» 이 난다.

    서버: can_build = 평면도 재료 · 급수방식.  화면도 그 둘만 본다.
    """
    py = open(os.path.join(_ROOT, "routes", "module_f", "api_merge.py"),
              encoding="utf-8").read()
    i = py.index('"can_build"')
    line = py[i:py.index("\n", i)]
    assert 'mats["plan"]' in line and "supply_mode" in line, line
    js = _js()
    fn = _fn(js, "  function mergeMissing(d)")
    assert ".plan" in fn and "d.mode" in fn, fn
    for kind in ("system", "machineroom"):
        assert f'"{kind}"' not in fn, \
            f"서버는 {kind} 를 안 보는데 화면만 막으면 못 넘어간다"
