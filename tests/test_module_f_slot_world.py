# -*- coding: utf-8 -*-
"""[슬롯] 슬롯을 바꾸면 **도면도 바뀐다** — 남의 도면이 남지 않는다.

■ 사용자 지적

  「계통도·기계실 업로드 후 다시 평면도로 돌아갔다가 계통도·기계실로 돌아가면
  그 도면이 아예 스크린에서 사라진다.」

■ 실측으로 잡은 것

  슬롯 탭을 오갈 때 `switchSlot` 이 도면(`S.world`)을 갈래마다 따로 불렀는데,
  **손질(edit) 갈래가 안 불렀다.** 그래서 평면도로 돌아오면 `S.world` 가
  직전 슬롯 것 그대로였다 — 실측: 평면도인데 묶음이 계통도 24개 · 기계실 19개
  (제 값은 30개).

  그 상태가 왜 «사라짐» 으로 보이나: 슬롯마다 **좌표계가 아예 다르다**
  (계통도 x≈-660,214 · 평면도 x≈248,153). 남의 도면이 남아 있는 채로 무엇이든
  `fit(S.world.bounds)` 를 부르면 시점이 남의 좌표로 튀고, 화면은 텅 빈다.

■ 그래서 여기서 지키는 둘

  ⑴ 슬롯을 바꾸면 그 슬롯의 도면을 **한 곳에서** 갈아 끼운다(갈래마다 따로
     부르면 또 한 갈래가 빠진다).
  ⑵ 못 읽으면 «남의 도면» 을 남기느니 **비우고 말한다.**
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

_STUB = r"""
const CALLS = {loaded: [], say: [], stage: [], fit: 0};
let SERVER_SLOT = "system";
let RAW_FAILS = false;
const S = {sid: "x", slot: FROM, method: METHOD, zones: [], undo: [],
           world: {tag: "system", bundles: [1, 2], bounds: {}},
           edit: {tag: "system"}, key: "sys", sub: {}, subGraph: {tag: "sys"},
           view: {}};
const $ = () => ({textContent: "", title: "", value: "",
                  classList: {toggle(){}, add(){}, remove(){}}});
const busy = () => {};
const say = (m, k) => CALLS.say.push([String(m), k]);
const setStage = (n) => CALLS.stage.push(n);
const fit = () => { CALLS.fit += 1; };
const draw = () => {};
const renderSlots = () => {};
const post = async (_p, body) => {
  SERVER_SLOT = body.kind;
  return {slots: SLOTS.map((s) => ({...s, active: s.kind === body.kind}))};
};
const api = async (p) => {
  if (p.indexOf("/auto/state") >= 0) return {method: METHOD, dxf_name: "x"};
  return {};
};
async function loadWorldRaw() {
  CALLS.loaded.push(SERVER_SLOT);
  if (RAW_FAILS) throw new Error("못 읽음");
  S.world = {tag: SERVER_SLOT, bundles: [1], bounds: {}};
  return S.world;
}
async function loadSub() { CALLS.loaded.push("sub"); }
async function loadEdit() { CALLS.loaded.push("edit"); }
async function loadAuto() { CALLS.loaded.push("auto"); }
async function loadWorld(reuse) { CALLS.loaded.push("world:" + !!reuse); }
async function loadRecon() {}
async function autoStart() {}
"""


def _run(slots, body, frm="system", method=None):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  async function switchSlot(kind)")
    src = js[i:js.index("\n  }\n", i) + 4]
    prog = "\n".join([f"const SLOTS = {json.dumps(slots)};",
                      f"const FROM = {json.dumps(frm)};",
                      f"const METHOD = {json.dumps(method)};",
                      _STUB, src, body])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-800:]
    return json.loads(out.stdout)


_SLOTS = [
    {"kind": "plan", "label": "평면도", "opened": True, "stage": "edit",
     "key": "p"},
    {"kind": "system", "label": "계통도", "opened": True, "stage": "sub",
     "key": "s"},
    {"kind": "machineroom", "label": "기계실", "opened": True,
     "stage": "sub", "key": "m"},
]


def test_손질_슬롯으로_돌아와도_제_도면을_받는다():
    """★이것이 실제 결함이었다 — edit 갈래만 도면을 안 갈아 끼웠다."""
    got = _run(_SLOTS, """
switchSlot("plan").then(() => console.log(JSON.stringify(
  {world: S.world && S.world.tag, calls: CALLS})));""",
               frm="machineroom", method="manual")
    assert got["world"] == "plan", "평면도인데 남의 도면이 남았다"
    assert "edit" in got["calls"]["loaded"], got["calls"]["loaded"]


def test_어느_갈래로_가든_도면을_갈아_끼운다():
    """갈래마다 따로 부르면 또 한 갈래가 빠진다 — 한 곳에서 한 번만."""
    for kind, stage, frm in (("plan", "edit", "system"),
                             ("plan", "pick", "system"),
                             ("system", "sub", "plan"),
                             ("machineroom", "sub", "plan")):
        slots = [dict(s, stage=stage if s["kind"] == kind else s["stage"])
                 for s in _SLOTS]
        got = _run(slots, f"""
switchSlot({json.dumps(kind)}).then(() => console.log(JSON.stringify(
  {{world: S.world && S.world.tag, calls: CALLS}})));""",
                   frm=frm, method="manual")
        assert got["world"] == kind, (kind, stage, got)
        assert got["calls"]["loaded"][0] == kind, got["calls"]["loaded"]


def test_못_읽으면_남의_도면을_남기지_않고_말한다():
    """★남겨 두는 것이 가장 나쁘다 — 시점이 남의 좌표로 튀어 화면이 빈다."""
    got = _run(_SLOTS, """
RAW_FAILS = true;
switchSlot("plan").then(() => console.log(JSON.stringify(
  {world: S.world, say: CALLS.say})));""", frm="system", method="manual")
    assert got["world"] is None, "못 읽었는데 남의 도면이 남았다"
    assert any(k == "err" for _m, k in got["say"]), got["say"]
    assert any("못 읽" in m for m, _k in got["say"]), got["say"]


def test_안_연_슬롯으로_가면_비운다():
    slots = [dict(s, opened=(s["kind"] != "machineroom")) for s in _SLOTS]
    got = _run(slots, """
switchSlot("machineroom").then(() => console.log(JSON.stringify(
  {world: S.world, edit: S.edit, stage: CALLS.stage,
   loaded: CALLS.loaded})));""", frm="plan", method="manual")
    assert got["world"] is None and got["edit"] is None
    assert got["stage"] == ["open"]
    assert got["loaded"] == [], "도면이 없는데 읽으러 갔다"


def test_슬롯이_바뀌면_경로_그래프도_버린다():
    """남의 도면 그래프 위에서 선이 따라오면 그 길로 추출까지 간다."""
    got = _run(_SLOTS, """
switchSlot("plan").then(() => console.log(JSON.stringify(
  {graph: S.subGraph, zones: S.zones.length, undo: S.undo.length})));""",
               frm="system", method="manual")
    assert got["graph"] is None
    assert got["zones"] == 0 and got["undo"] == 0


# ── 찍기의 «찍힌 헤드» 수 (2026-09-08) ─────────────────────────────
#
# 사용자 지적: 「평면도에서 최불리 선정 누르니까 또 작동 안 한다」.
#
# 재 보니 단추는 멀쩡했다 — 찍기에서 클릭 한 번이 헤드 «칸(부류)» 을 통째로
# 끄는 바람에 헤드가 111개에서 5개로 떨어져 있었고, 최불리는 「5개뿐」이라고
# 정당하게 거절한 것이었다. 문제는 **그 사실이 찍기 화면에 안 보였다는 것**
# 이다. 착지가 찍기로 바뀌면서 사람이 그 화면에서 손을 대게 됐으니 더 그렇다.


def _pick_state(**kw):
    base = {"n_heads": 3, "n_head_circles": 111, "has_tri_heads": False,
            "materials": [1, 2, 3], "mode": "헤드", "armed": True,
            "mat_done": True, "head_label": "상향하향"}
    base.update(kw)
    return base


def _run_pick(state, k=30):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  function renderPickCount(p)")
    src = js[i:js.index("\n  }\n", i) + 4]
    prog = "\n".join([
        f"const P = {json.dumps(state)}; const K = {k};",
        "let CLS = new Set(); let HTML = '';",
        "const EL = {classList: {toggle: (c, v) => { if (v) CLS.add(c);"
        "                        else CLS.delete(c); }}};",
        "Object.defineProperty(EL, 'innerHTML', {set: (v) => { HTML = v; }});",
        "const $ = () => EL;",
        "const edK = () => K;",
        "const kv = (a, b) => `${a}=${b};`;",
        src,
        "renderPickCount(P);",
        "console.log(JSON.stringify({html: HTML, warn: CLS.has('warn')}));",
    ])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-600:]
    return json.loads(out.stdout)


def test_찍기가_헤드_개수를_센다():
    """★칸 수를 세면 안 된다 — 실측 대명동은 칸 3개에 헤드 111개다."""
    got = _run_pick(_pick_state())
    assert "111" in got["html"], got["html"]
    assert not got["warn"], "111개인데 경고를 띄웠다"


def test_기준개수보다_적으면_찍기에서_경고한다():
    """★이것이 없어서 111 → 5 가 조립 뒤에야 보였다."""
    got = _run_pick(_pick_state(n_head_circles=5, n_heads=1))
    assert "5" in got["html"] and "30" in got["html"], got["html"]
    assert got["warn"], "모자란데 경고가 없다"


def test_칸_수를_기준개수와_견주지_않는다():
    """칸(3)을 K(30)와 견주면 늘 «모자란다» 는 엉터리 경고가 된다."""
    got = _run_pick(_pick_state(n_heads=3, n_head_circles=111))
    assert not got["warn"], got["html"]


def test_삼각형_헤드가_있으면_단정하지_않는다():
    """삼각형 헤드는 이 수에 안 들어간다 — 없는 수로 겁주지 않는다."""
    got = _run_pick(_pick_state(n_head_circles=5, has_tri_heads=True))
    assert not got["warn"], "셀 수 없는 것을 두고 모자란다고 했다"
    assert "삼각형" in got["html"], got["html"]
