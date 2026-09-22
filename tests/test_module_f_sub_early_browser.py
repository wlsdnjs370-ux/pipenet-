# -*- coding: utf-8 -*-
"""[오너 2026-09-22 · 그림 38 ②] 계통도 칸 — 도면을 먼저 띄우고 ★추적은 이어서.

  서버는 레이어 표까지 만든 도면(world)을 먼저 내려보내고, ★추적 레이어는 그 뒤에
  고른다. 그동안 world 의 ★칸은 «고르는 중»(mode = pending) 이다.

  ⑴ «고르는 중» 인 표에는 ★추적 표가 하나도 없고, 안내 줄이 «고르는 중» 이다
     (자동 레이어에 ★를 잘못 달지 않는다).
  ⑵ 잡이 끝나고 `/sub/graph` 가 ★추적을 주면 표가 다시 그려져 ★가 붙는다.
  ⑶ 그 사이 사람이 켜고 끈 것은 그대로다.
  ⑷ ★를 못 고른 도면(서버가 null)은 «고르는 중» 에 머물지 않는다.
"""
from __future__ import annotations

import copy
import json
from pathlib import Path

from test_module_f_system_layers_browser import WORLD
from test_module_f_viewport_browser import install_page, pw_api

PENDING = {"mode": "pending", "layers": None, "junk": [], "candidates": [], "forced": None,
           "split": None, "reason": "경로 추적 레이어(★)를 고르는 중입니다 — 잠시 뒤 채워집니다."}
PICKED = {"mode": "pipe", "layers": ["6-소화-가지관"], "junk": ["6-소화-TXT"],
          "candidates": ["6-소화-가지관"], "forced": 0, "split": {"split_entities": 0, "cuts": 0},
          "reason": "켜진 배관 1개로만 추적합니다"}
GRAPH = {"ok": True, "kind": "system", "nodes": [[0, 0], [0, 1000]], "edges": [[0, 1, 1000.0, False]],
         "forced": 0, "components": 1, "forced_penalty_mm": 100000.0, "auto_layers": [],
         "chosen_auto": False, "narrowed": None, "layers": [], "chosen": None,
         "snap_default_mm": 2500.0, "trace": PICKED, "trace_on": True}


def test_pending_trace_then_filled_by_sub_graph():
    source = (Path(__file__).resolve().parents[1] / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();',
                            '  window.__sysUi = {build: buildLayers, loadSubGraph};\n'
                            '  setStage("open");\n  loadSaved();')
    world = copy.deepcopy(WORLD)
    world["sub_layers"]["trace"] = PENDING
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        errors = install_page(page, source)
        page.route("http://module-f.test/api/module-f/sub/graph",
                   lambda req: req.fulfill(body=json.dumps(GRAPH), content_type="application/json"))
        page.wait_for_function("!!window.__sysUi")
        page.evaluate("""w => { Object.assign(__mf, {slot: 'system', key: 'B1F', sid: 'early'});
                               __mf.world = w; __sysUi.build(); }""", world)
        # ⑴ 고르는 중 — ★ 없음 · 안내 줄
        assert page.locator("#layers .tag.trace").count() == 0
        assert "고르는 중" in page.locator("#layers .lynote").inner_text()
        # ⑶ 그 사이 사람이 건축 레이어를 끈다
        page.evaluate("""() => { const cb = document.querySelector('#layers input[data-layer="A-B1"]');
                                 cb.checked = false; cb.dispatchEvent(new Event('change')); }""")
        hidden = set(page.evaluate("[...__mf.hidden]"))
        assert "A-B18" in hidden
        # ⑵ 잡이 끝나 /sub/graph 가 ★를 준다 → 표에 ★
        page.evaluate("async () => { await __sysUi.loadSubGraph(); }")
        assert page.locator("#layers .tag.trace").count() == 1
        assert "6-소화-가지관" in page.locator("#layers label.sub",
                                             has=page.locator(".tag.trace")).inner_text()
        assert "켜진 배관 1개로만" in page.locator("#layers .lynote").inner_text()
        assert page.evaluate("__mf.world.sub_layers.trace.mode") == "pipe"
        assert set(page.evaluate("[...__mf.hidden]")) == hidden
        # 같은 ★를 다시 받으면 표를 새로 그리지 않는다(바뀐 것이 없다)
        page.evaluate("document.querySelector('#layers .grp').dataset.mark = 'kept'")
        page.evaluate("async () => { await __sysUi.loadSubGraph(); }")
        assert page.evaluate("document.querySelector('#layers .grp').dataset.mark || ''") == "kept"
        # ★를 못 고른 도면(서버 null) — «고르는 중» 에 머물지 않고 자동 레이어 표시로
        page.unroute("http://module-f.test/api/module-f/sub/graph")
        page.route("http://module-f.test/api/module-f/sub/graph",
                   lambda req: req.fulfill(body=json.dumps(dict(GRAPH, trace=None, trace_on=False)),
                                           content_type="application/json"))
        page.evaluate("w => { __mf.world = w; __sysUi.build(); }", world)
        assert "고르는 중" in page.locator("#layers .lynote").inner_text()
        page.evaluate("async () => { await __sysUi.loadSubGraph(); }")
        assert "고르는 중" not in page.locator("#layers .lynote").inner_text()
        assert page.evaluate("__mf.world.sub_layers.trace") is None
        # 자동 레이어(auto) 에 ★ — 종전 규칙(sysTraceSet 의 자동 표시)
        assert page.locator("#layers .tag.trace").count() >= 1
        assert not errors, errors
        browser.close()
