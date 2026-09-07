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
