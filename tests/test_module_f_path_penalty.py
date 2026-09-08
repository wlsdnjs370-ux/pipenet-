# -*- coding: utf-8 -*-
"""경로 미리보기와 추출이 **같은 비용**으로 고른다.

■ 무엇이 어긋나 있었나 (2026-09-03 실측)

  계통도·기계실의 «선이 따라오는» 미리보기는 브라우저가 직접 최단경로를 푼다.
  그때 추측연결(force_connect 가 이은 직선)에 벌점을 더하는데, 그 값이
  **서버 1e9 · 화면 1e6 으로 1000배** 달랐다. 그래프는 같은 것을 쓰는데
  비용 함수가 달랐던 셈이다.

  도면 단위가 mm 라 1e6 은 1 km 다. 큰 도면에서 실배관 우회가 그 값에 닿으면
  화면은 추측 직선을, 서버는 실배관을 고른다 — 「미리보기와 추출이 같은
  그래프를 쓴다」는 약속이 그 자리에서 깨진다.

■ 이 시험이 지키는 것

  값을 **화면에도 서버에도 적지 않는다.** 추출이 실제로 쓰는 엔진 기본값을
  읽어 payload 에 실어 보내고, 화면은 그것만 쓴다. 못 받으면 미리보기를
  접는다 — 추출과 다른 길을 그리느니 안 그리는 편이 낫다.
"""
from __future__ import annotations

import inspect
import os
import re
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def test_서버가_보내는_벌점이_추출이_쓰는_그_값이다():
    from remote30_graph import _shortest_path
    from routes.module_f.subdrawing import _forced_penalty_mm

    engine = inspect.signature(_shortest_path).parameters["penalty_mm"].default
    assert _forced_penalty_mm() == float(engine), (
        f"화면에 보낼 값 {_forced_penalty_mm()} 이 추출이 쓰는 {engine} 과 다르다")


def test_그래프_payload_에_벌점이_실린다():
    """화면이 읽을 자리에 실제로 들어가는가 — 없으면 화면은 미리보기를 접는다."""
    from routes.module_f.subdrawing import graph_payload

    # 선분 두 개짜리 최소 도면. 엔티티 규약은 `path_graph` 가 받는 그대로.
    ents = [
        {"layer": "PIPE", "type": "LINE", "points": [(0.0, 0.0), (1000.0, 0.0)]},
        {"layer": "PIPE", "type": "LINE",
         "points": [(1000.0, 0.0), (2000.0, 0.0)]},
    ]
    got = graph_payload(ents)
    assert "forced_penalty_mm" in got, sorted(got)
    assert got["forced_penalty_mm"] > 0


def test_화면은_벌점을_스스로_적지_않는다():
    """★JS 에 숫자를 손으로 적으면 이 어긋남이 되살아난다.

    소스에서 낱말을 찾는 검사이지만, 여기서 막으려는 것이 바로 «리터럴이
    다시 들어오는 것» 이라 이 방식이 맞다. 다만 «있어야 할 문자열» 이 아니라
    «있으면 안 되는 리터럴» 을 본다 — 이름을 바꿔도 부러지지 않는다.
    """
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("function subPath(")
    body = js[i:js.index("\n  }\n", i)]
    # 벌점 자리에 «숫자 리터럴» 이 곱해지거나 더해지면 안 된다.
    bad = re.findall(r"forced\s*\?\s*([0-9][0-9eE_.+]*)\s*:", body)
    assert not bad, f"화면이 벌점을 스스로 적었다: {bad}"
    assert "forced_penalty_mm" in body, "서버가 보낸 값을 안 쓴다"


def test_못_받으면_미리보기를_접는다():
    """어림값으로 그리지 않는다 — 어긋난 미리보기가 가장 나쁘다."""
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("function subPath(")
    body = js[i:js.index("\n  }\n", i)]
    j = body.index("forced_penalty_mm")
    assert "return null" in body[j:j + 220], \
        "벌점을 못 받았는데 그리려 든다"


def test_받은_벌점이_화면_상태까지_살아_남는다():
    """★위 네 검사가 모두 통과하는 동안 미리보기는 **한 번도 안 그려졌다**.

    서버는 값을 보냈고(검사 2), 화면 코드는 그 이름을 읽고(검사 3), 못 받으면
    접는다(검사 4). 그런데 정작 응답을 **골라 담는** 자리에서 이 필드만
    빠져 있었다 — `S.subGraph = {nodes, edges, adj, forced, …}`. 그래서 화면은
    늘 «못 받은» 상태였고, 선은 어느 도면에서도 따라오지 않았다. 오류도
    콘솔도 없다.

    소스를 읽는 검사로는 이 틈을 못 본다. 응답을 넣고 **경로가 실제로 나오는지**
    를 본다 — 값이 도중에 떨어지면 여기서 null 이 된다.
    """
    import json
    import shutil
    import subprocess

    import pytest

    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")

    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()

    def fn(head):
        i = js.index(head)
        return js[i:js.index("\n  }\n", i) + 4]

    # 서버 payload 그대로 — 절점 3개가 한 줄로 이어진 최소 그래프.
    payload = {
        "ok": True,
        "nodes": [[0, 0], [1000, 0], [2000, 0]],
        "edges": [[0, 1, 1000.0, 0], [1, 2, 1000.0, 0]],
        "forced": 0, "components": 1, "layers": [], "chosen": None,
        "forced_penalty_mm": 1e9,
    }
    prog = "\n".join([
        # 화면의 실제 모양 그대로 — `S.sub` 는 늘 있다(미리보기 경로가
        # 여기 산다). 껍데기만 주면 새 코드가 없는 자리를 짚는다.
        "const S = {sid: 'x', subGraph: null,"
        " sub: {picks: [null, null], preview: null}};",
        f"const PAYLOAD = {json.dumps(payload)};",
        "const post = async () => PAYLOAD;",
        "const renderSubLayers = () => {};",
        "const draw = () => {};",
        "const say = () => {};",
        fn("async function loadSubGraph("),
        fn("function subNearest("),
        fn("function subPath("),
        "loadSubGraph().then(() => {",
        "  const a = subNearest(0, 0), b = subNearest(2000, 0);",
        "  console.log(JSON.stringify({",
        # undefined 는 JSON 에서 사라진다 — null 로 바꿔 «없다» 를 남긴다.
        "    pen: S.subGraph.forced_penalty_mm === undefined",
        "         ? null : S.subGraph.forced_penalty_mm,",
        "    path: subPath(a, b),",
        "  }));",
        "});",
    ])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-600:]
    got = json.loads(out.stdout)
    assert got["pen"] == 1e9, (
        "서버가 보낸 벌점이 화면 상태까지 못 왔다 — 미리보기가 통째로 죽는다")
    assert got["path"] == [0, 1, 2], got["path"]
