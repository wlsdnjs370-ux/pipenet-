# -*- coding: utf-8 -*-
"""[오너 2026-09-22 · 그림 42] 화면 — 계통도 칸 옆 «＋ 계통도 추가» 와 칸마다의 두 점 이름.

  ⑴ 계통도가 한 장이면 칸 줄은 종전과 같고, 계통도 칸 바로 옆에 «＋ 계통도 추가» 가 붙는다.
  ⑵ 누르면 «계통도 2» 칸이 생기고 그 칸으로 넘어간다. 이름은 계통도 1 · 2, 추가 단추는
     맨 끝 계통도 옆으로 옮겨 가고, 더한 칸에만 × 가 있다.
  ⑶ 두 점의 이름이 칸의 자리를 따른다(① 기계실 쪽 끝 · ② 평면도 쪽 끝).
"""
from __future__ import annotations

from pathlib import Path

from test_module_f_viewport_browser import install_page, pw_api

ROOT = Path(__file__).resolve().parents[1]


def slot(kind, label, active=False, removable=False, role=None):
    return {"kind": kind, "label": label, "active": active, "opened": False, "stage": "",
            "method": None, "key": None, "designed": False,
            "role": role or ("system" if kind.startswith("system") else kind),
            "removable": removable}


ONE = {"ok": True, "active": "system", "can_add_system": True,
       "slots": [slot("plan", "평면도"), slot("system", "계통도", active=True),
                 slot("machineroom", "기계실")]}
TWO = {"ok": True, "active": "system", "can_add_system": True, "added": "system2",
       "added_label": "계통도 2",
       "slots": [slot("plan", "평면도"), slot("system", "계통도 1", active=True),
                 slot("system2", "계통도 2", removable=True), slot("machineroom", "기계실")]}
SWITCHED = {**TWO, "active": "system2", "switched": "system2",
            "slots": [slot("plan", "평면도"), slot("system", "계통도 1"),
                      slot("system2", "계통도 2", active=True, removable=True),
                      slot("machineroom", "기계실")]}


def texts(page):
    return [t.strip() for t in page.locator("#slots > button").all_inner_texts()]


def test_add_system_slot_then_labels_follow_position():
    source = (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();',
                            '  window.__slotUi = {renderSlots, subSpec};\n'
                            '  setStage("open");\n  loadSaved();')
    posted = []
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        errors = install_page(page, source)

        def add(req):
            posted.append(("add", req.request.post_data_json))
            req.fulfill(json=TWO)

        def switch(req):
            posted.append(("switch", req.request.post_data_json))
            req.fulfill(json=SWITCHED)

        page.route("http://module-f.test/api/module-f/slot/add-system", add)
        page.route("http://module-f.test/api/module-f/slot/switch", switch)
        page.wait_for_function("!!window.__slotUi")
        page.evaluate("s => { __mf.sid = 'chain'; __slotUi.renderSlots(s); }", ONE)
        # ⑴ 한 장 — 종전 이름 그대로, 계통도 칸 바로 옆에 추가 단추
        assert texts(page) == ["평면도", "계통도", "＋ 계통도 추가", "기계실"]
        assert "1/3" not in page.locator("#slots .note").inner_text()
        assert "0/3 열림" in page.locator("#slots .note").inner_text()
        # 계통도 1 로 보면 두 점은 종전 이름(펌프 · 알람밸브)
        spec = page.evaluate("__slotUi.subSpec()")
        assert (spec["a"], spec["b"], spec["title"]) == ("펌프", "알람밸브", "계통도")
        # ⑵ 누른다 → 칸이 생기고 그 칸으로 넘어간다
        page.click("#slot-add-system")
        page.wait_for_function("window.__mf.slot === 'system2'")
        assert posted[0] == ("add", {"sid": "chain"})
        assert posted[1] == ("switch", {"sid": "chain", "kind": "system2"})
        got = texts(page)
        assert got[:2] == ["평면도", "계통도 1"]
        assert got[2].startswith("계통도 2") and got[2].endswith("×")
        assert got[3:] == ["＋ 계통도 추가", "기계실"]
        assert page.locator("#slots .slot-x").count() == 1
        assert "0/4 열림" in page.locator("#slots .note").inner_text()
        # ⑶ 두 점 이름 — 계통도 2(맨 끝)는 ① 펌프 · ② 계통도 1과 만나는 점
        spec = page.evaluate("__slotUi.subSpec()")
        assert spec["title"] == "계통도 2"
        assert spec["a"] == "펌프 (기계실과 만나는 점)"
        assert spec["b"] == "계통도 1과 만나는 점"
        assert spec["keys"] == ["pump_x", "pump_y", "av_x", "av_y"]
        # 계통도 1 로 돌아가 보면 ① 계통도 2와 만나는 점 · ② 알람밸브
        spec = page.evaluate("() => { __mf.slot = 'system'; return __slotUi.subSpec(); }")
        assert spec["a"] == "계통도 2와 만나는 점"
        assert spec["b"] == "알람밸브 (평면도와 만나는 점)"
        assert not errors, errors
        browser.close()
