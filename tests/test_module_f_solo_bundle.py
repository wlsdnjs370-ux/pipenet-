# -*- coding: utf-8 -*-
"""[§27] «이 묶음만 크게» — 판정을 늘리는 대신 사람이 보게 한다.

■ 왜 판정이 아니라 «보기» 인가

  이름 사전은 도면에 따라 **정반대로** 읽는다. 대명동 계통도에서 `SP` 는
  사전이 「스프링클러 배관」으로 읽지만 실제 내용은 **헤드 기호 918개** 이고,
  진짜 배관은 사전이 OTHER 로 떨어뜨린 `0` 레이어에 있다. 여기에 규칙을 하나
  더 얹으면 «틀린 확신» 만 늘어난다 — 선분 길이와 접점, 두 지표를 실제로 재
  보고 둘 다 기각했다(평면도의 진짜 배관은 긴 선분이 10~18%, 계통도의 층
  구획선은 100%다. 그럴듯한 규칙이 정확히 거꾸로 작동한다).

  그래서 이 기능은 아무 것도 «말하지» 않는다. 묶음 하나만 남기고 그 범위로
  확대할 뿐이다. 판정을 안 하므로 틀릴 수가 없다.

■ 이 시험이 소스를 «읽지» 않고 «돌리는» 이유

  이 저장소에서 소스 문자열 검사는 여섯 번 깨졌다. 그리고 화면 코드의 사고는
  구문이 아니라 **상태**에서 난다: 실제로 처음 판에는 목록을 다시 그릴 때
  체크박스를 `true` 로 고정해서, 그림은 한 묶음인데 목록은 전부 켜진 얼굴을
  하고 있었다. `node --check` 도 브라우저 콘솔도 조용하다. 그러니 JS 를 그대로
  꺼내 node 로 돌려 상태를 본다.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 화면이 기대하는 최소한의 DOM·이웃 함수. 진짜 브라우저 검증은 따로 있고
# (`scripts/_verify_module_f_solo.py`) 여기서 보는 것은 «상태 전이» 다.
_STUB = r"""
const CALLS = {fit: [], draw: 0, say: []};
function mkEl(tag) {
  const e = {tag, _cls: new Set(), children: [], dataset: {}, style: {},
             checked: false, textContent: "", title: "", type: ""};
  e.classList = {
    toggle: (c, v) => { if (v) e._cls.add(c); else e._cls.delete(c); },
    contains: (c) => e._cls.has(c),
  };
  Object.defineProperty(e, "className", {
    get: () => [...e._cls].join(" "),
    set: (v) => { e._cls = new Set(String(v).split(/\s+/).filter(Boolean)); },
  });
  e.append = (...xs) => { e.children.push(...xs); };
  e.appendChild = (x) => { e.children.push(x); };
  e.querySelectorAll = (sel) => {
    const want = sel.replace(".", "");
    const out = [];
    const walk = (n) => {
      for (const c of n.children) {
        if (sel === "input" ? c.tag === "input" : c._cls.has(want)) out.push(c);
        walk(c);
      }
    };
    walk(e);
    return out;
  };
  return e;
}
const document = {createElement: mkEl};
const BOX = mkEl("div");
Object.defineProperty(BOX, "innerHTML", {set: (_v) => { BOX.children = []; }});
// id 마다 다른 것을 준다 — 하나로 뭉치면 `ly-all`·`ly-none` 의 onclick 이
// 서로를 덮어써 «둘 다 도는지» 를 볼 수 없다.
const ELS = {layers: BOX};
const $ = (id) => (ELS[id] || (ELS[id] = mkEl("div")));
const num = (v) => Number(v) || 0;
const draw = () => { CALLS.draw += 1; };
const fit = (b) => { CALLS.fit.push(b); };
const say = (m, _k) => { CALLS.say.push(m); };
const S = {hidden: new Set(), world: {bounds: {minx: 0, miny: 0,
                                              maxx: 999, maxy: 999},
                                     bundles: []}};
"""


def _fn(js: str, head: str) -> str:
    i = js.index(head)
    return js[i:js.index("\n  }\n", i) + 4]


def _chunk(js: str, head: str, end: str) -> str:
    i = js.index(head)
    return js[i:js.index(end, i) + len(end)]


def _run(body: str):
    """화면 코드를 그대로 꺼내 node 로 돌리고 마지막 상태를 돌려받는다."""
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    src = "\n".join([
        _STUB,
        "let _soloId = null;",
        _fn(js, "function buildLayers()"),
        _fn(js, "function markSolo()"),
        _fn(js, "function bundleBounds(id)"),
        _fn(js, "function soloBundle(id)"),
        # 「모두 켜기·끄기」 단추도 화면의 것을 그대로 돌린다.
        _chunk(js, '$("ly-all").onclick', "function box_all(v) {\n"
                                         "    for (const cb of "
                                         '$("layers").querySelectorAll'
                                         '("input")) cb.checked = v;\n  }'),
        body,
    ])
    # ★인코딩을 안 박으면 윈도우 기본(cp949)으로 읽다 한글 메시지에서
    #   읽기 스레드가 죽고, 결과가 조용히 None 이 된다.
    out = subprocess.run([node, "-e", src], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-800:]
    return json.loads(out.stdout)


# 세 묶음 — 둘은 좌표가 있고 하나는 «잴 것이 없다»(세그먼트가 안 실려 온 묶음).
_SETUP = r"""
S.world.bundles = [
  {id: "b1", layer: "0", name: "LINE", cat: "OTHER", css: "#888",
   len_m: 10, len_mid: 5, n_all: 3, n_seg: 3, n_circle: 0,
   segs: [0, 0, 100, 0, 100, 0, 100, 50]},
  {id: "b2", layer: "SP", name: "INSERT", cat: "HEAD", css: "#fa0",
   len_m: 0, len_mid: 0, n_all: 918, n_seg: 0, n_circle_all: 918,
   segs: [], circles: [[900, 900], [950, 980]]},
  {id: "b3", layer: "TEXT", name: "TEXT", cat: "OTHER", css: "#555",
   len_m: 0, len_mid: 0, n_all: 7, n_seg: 0, n_circle: 0, segs: []},
];
const rows = () => BOX.children.map((lb) => {
  const kids = lb.children;
  return {
    checked: kids.find((k) => k.tag === "input").checked,
    solo_on: kids.some((k) => k._cls.has("solo") && k._cls.has("on")),
  };
});
const dump = (extra) => console.log(JSON.stringify(Object.assign(
  {rows: rows(), hidden: [...S.hidden].sort(), calls: CALLS,
   soloId: _soloId}, extra || {})));
buildLayers();
"""


def test_묶음마다_단추가_하나씩_붙는다():
    got = _run(_SETUP + r"""
console.log(JSON.stringify({
  n: BOX.children.length,
  solo: BOX.children.filter((lb) =>
    lb.children.some((k) => k._cls.has("solo"))).length,
}));""")
    assert got["n"] == 3
    assert got["solo"] == 3, "단추가 빠진 묶음이 있다"


def test_누르면_그_묶음만_남고_목록도_같은_말을_한다():
    """★처음 판의 실제 버그 — 그림은 하나인데 목록은 전부 켜진 얼굴이었다.

    `buildLayers` 가 체크박스를 `true` 로 고정했기 때문이다. 화면과 목록이
    다른 말을 하면 사람은 목록을 믿는다.
    """
    got = _run(_SETUP + 'soloBundle("b2"); dump();')
    assert got["hidden"] == ["b1", "b3"]
    assert [r["checked"] for r in got["rows"]] == [False, True, False]
    assert [r["solo_on"] for r in got["rows"]] == [False, True, False]
    assert got["soloId"] == "b2"


def test_그_묶음_범위로_확대한다():
    """세계 전체가 아니라 «그 묶음» 이 화면을 채워야 볼 수 있다."""
    got = _run(_SETUP + 'soloBundle("b2"); dump();')
    assert got["calls"]["fit"] == [{"minx": 900, "miny": 900,
                                   "maxx": 950, "maxy": 980}]
    assert got["calls"]["draw"] == 1


def test_잴_것이_없으면_화면을_안_옮긴다():
    """★엉뚱한 데로 튀느니 그대로 둔다 — 끄기는 하되 시점은 건드리지 않는다."""
    got = _run(_SETUP + 'soloBundle("b3"); dump();')
    assert got["calls"]["fit"] == [], "범위를 못 재는데 화면을 옮겼다"
    assert got["hidden"] == ["b1", "b2"], "끄는 것까지 그만두면 안 된다"
    assert got["calls"]["draw"] == 1


def test_같은_단추를_다시_누르면_되돌아온다():
    got = _run(_SETUP + 'soloBundle("b2"); soloBundle("b2"); dump();')
    assert got["hidden"] == []
    assert [r["checked"] for r in got["rows"]] == [True, True, True]
    assert got["soloId"] is None
    assert got["calls"]["fit"][-1] == {
        "minx": 0, "miny": 0, "maxx": 999, "maxy": 999}


def test_다른_단추를_누르면_그리로_옮겨간다():
    """되돌리기를 거치지 않아도 된다 — 묶음을 훑어볼 때 이 편이 자연스럽다."""
    got = _run(_SETUP + 'soloBundle("b2"); soloBundle("b1"); dump();')
    assert got["hidden"] == ["b2", "b3"]
    assert got["soloId"] == "b1"
    assert [r["solo_on"] for r in got["rows"]] == [True, False, False]


def test_모두_켜기_끄기도_그_상태를_끝낸다():
    """전부 켠(또는 전부 끈) 화면에서 단추 하나만 «보는 중» 으로 남으면 거짓말이다.

    화면의 그 단추를 그대로 꺼내 누른다 — 여기서 손으로 흉내 내면 실제
    핸들러가 안 고쳐져도 시험만 통과한다.
    """
    on = _run(_SETUP + 'soloBundle("b2"); $("ly-all").onclick(); dump();')
    assert not any(r["solo_on"] for r in on["rows"]), "모두 켜기 뒤 표시가 남았다"
    assert on["hidden"] == []
    assert [r["checked"] for r in on["rows"]] == [True, True, True]

    off = _run(_SETUP + 'soloBundle("b2"); $("ly-none").onclick(); dump();')
    assert not any(r["solo_on"] for r in off["rows"]), "모두 끄기 뒤 표시가 남았다"
    assert off["hidden"] == ["b1", "b2", "b3"]


def test_사람이_직접_켜고_끄면_그_상태는_끝난다():
    """안 풀면 같은 단추가 «되돌리기» 로 남아, 눌러도 아무 일이 없어 보인다."""
    got = _run(_SETUP + r"""
soloBundle("b2");
const cb = BOX.children[0].children.find((k) => k.tag === "input");
cb.checked = true; cb.onchange();
const after_manual = {soloId: _soloId, marked: rows().map((r) => r.solo_on)};
soloBundle("b2");
dump({after_manual: after_manual});""")
    assert got["after_manual"]["soloId"] is None
    assert got["after_manual"]["marked"] == [False, False, False], \
        "단추에 «보고 있는 중» 표시가 남았다"
    # 그 다음 같은 단추를 누르면 되돌리기가 아니라 «b2 만 보기» 여야 한다.
    assert got["hidden"] == ["b1", "b3"]
