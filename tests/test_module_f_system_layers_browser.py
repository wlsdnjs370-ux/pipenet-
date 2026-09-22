# -*- coding: utf-8 -*-
"""[오너 2026-09-22] 계통도 칸 «레이어 표시» 표 — 실제 화면 코드로 확인한다.

  「계통도 쪽에는 기본값은 "건축 레이어"와 배관망(헤드, 배관, 밸브, 부속류 등) 관련
   레이어만 기본값으로 표시되게 하고 나머지는 아예 숨김처리한 뒤 별도의 건축 정리
   레이어처럼 좌측하단에 테이블 형태로 활성화/비활성화로 표시만 할 수 있게」

  ⑴ 처음 열면 숨긴 레이어 묶음은 안 그린다(배관망 · 건축만).
  ⑵ 묶음 표 셋 — 배관망 · 건축 · 숨긴 레이어. 추적 레이어에는 ★추적 표가 붙는다.
  ⑶ 체크는 화면 표시만 바꾼다. «기본값» 단추가 처음 상태로 되돌린다.
  ⑷ 평면도 칸은 종전 목록 그대로다.
"""
from __future__ import annotations

from pathlib import Path

from test_module_f_viewport_browser import install_page, pw_api


def _bundle(layer, color, css, x):
    return {"id": f"{layer}{color}", "layer": layer, "color": color, "name": str(color),
            "css": css, "cat": "OTHER", "segs": [x, 0, x, 1000], "circles": [], "arcs": [],
            "n_seg": 1, "n_all": 1, "n_circle": 0, "n_arc": 0, "len_m": 1.0, "len_mid": 1000}


WORLD = {
    "bundles": [_bundle("6-소화-가지관", 3, "#00e000", 0), _bundle("6-소화-가지관", 1, "#e00000", 50),
                _bundle("6-소화-밸브", 5, "#007fff", 100), _bundle("A-B1", 8, "#a0a0a0", 200),
                _bundle("6-소화-헤드-반경", 5, "#003fff", 300), _bundle("6-소화-TXT", 7, "#ffffff", 400)],
    "bounds": {"minx": -100, "miny": -100, "maxx": 600, "maxy": 1100},
    "counts": {"segs": 6, "circles": 0, "arcs": 0}, "dropped": {"segs": 0, "circles": 0, "arcs": 0},
    "sub_layers": {
        "rows": [
            {"layer": "6-소화-가지관", "group": "net", "why": "배관·헤드", "n": 2, "visible": True, "cat": "PIPE", "auto": True},
            {"layer": "6-소화-밸브", "group": "net", "why": "밸브·부속", "n": 1, "visible": True, "cat": "OTHER", "auto": True},
            {"layer": "A-B1", "group": "arch", "why": "건축", "n": 1, "visible": True, "cat": "ARCH", "auto": False},
            {"layer": "6-소화-헤드-반경", "group": "etc", "why": "CAD 꺼 둠", "n": 1, "visible": False, "cat": "HEAD", "auto": False},
            {"layer": "6-소화-TXT", "group": "etc", "why": "분류 안 됨", "n": 1, "visible": True, "cat": "OTHER", "auto": True}],
        "trace": {"mode": "pipe", "layers": ["6-소화-가지관"], "reason": "켜진 배관 1개로만 추적합니다"},
        "groups": {"net": "배관망", "arch": "건축", "etc": "숨긴 레이어"}},
}
HIDDEN_DEFAULT = {"6-소화-헤드-반경5", "6-소화-TXT7"}


def test_system_slot_layer_table_defaults_and_toggles():
    source = (Path(__file__).resolve().parents[1] / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();',
                            '  window.__sysUi = {build: buildLayers};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        errors = install_page(page, source)
        page.wait_for_function("!!window.__sysUi")
        page.evaluate("""w => { Object.assign(__mf, {slot: 'system', key: 'B1F'});
                               __mf.world = w; __sysUi.build(); }""", WORLD)
        # ⑴ 처음 열면 숨긴 레이어 묶음만 안 그린다
        assert set(page.evaluate("[...__mf.hidden]")) == HIDDEN_DEFAULT
        # ⑵ 묶음 표 셋 · ★추적
        assert page.locator("#layers .grp").count() == 3
        assert page.locator("#layers .grp.net .gname").inner_text() == "배관망"
        assert page.locator("#layers .tag.trace").count() == 1
        assert "6-소화-가지관" in page.locator("#layers label.sub", has=page.locator(".tag.trace")).inner_text()
        # 숨긴 묶음은 접혀 있다 — 줄이 안 보이고 묶음 칸만 있다
        assert page.locator('#layers label.sub input[data-layer="6-소화-TXT"]').count() == 0
        assert not page.locator("#ly-default").evaluate("e => e.classList.contains('hidden')")
        # ⑶ 레이어 하나 끄기 → 그 레이어의 묶음(색 둘) 전부 숨김
        page.evaluate("""() => { const cb = document.querySelector('#layers input[data-layer="6-소화-가지관"]');
                                 cb.checked = false; cb.dispatchEvent(new Event('change')); }""")
        assert {"6-소화-가지관3", "6-소화-가지관1"} <= set(page.evaluate("[...__mf.hidden]"))
        # 숨긴 묶음 전체 켜기(묶음 칸 체크)
        page.evaluate("""() => { const cb = document.querySelector('#layers .grp.etc input');
                                 cb.checked = true; cb.dispatchEvent(new Event('change')); }""")
        assert not (HIDDEN_DEFAULT & set(page.evaluate("[...__mf.hidden]")))
        # 기본값으로
        page.evaluate("document.getElementById('ly-default').click()")
        assert set(page.evaluate("[...__mf.hidden]")) == HIDDEN_DEFAULT
        # 슬롯을 오가도 사람이 고친 것은 도면 이름 단위로 남는다
        page.evaluate("""() => { const cb = document.querySelector('#layers input[data-layer="A-B1"]');
                                 cb.checked = false; cb.dispatchEvent(new Event('change')); }""")
        page.evaluate("w => { __mf.world = JSON.parse(JSON.stringify(w)); __sysUi.build(); }", WORLD)
        assert "A-B18" in set(page.evaluate("[...__mf.hidden]"))
        # ⑷ 평면도 칸 — 종전 목록(묶음 칸 없음 · 기본값 단추 숨김)
        page.evaluate("""w => { __mf.slot = 'plan'; __mf.hidden = new Set();
                               const x = JSON.parse(JSON.stringify(w)); delete x.sub_layers;
                               __mf.world = x; __sysUi.build(); }""", WORLD)
        assert page.locator("#layers .grp").count() == 0
        assert page.locator("#layers label").count() == len(WORLD["bundles"])
        assert page.locator("#ly-default").evaluate("e => e.classList.contains('hidden')")
        assert not errors, errors
        browser.close()
