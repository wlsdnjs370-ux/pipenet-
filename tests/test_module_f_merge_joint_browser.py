# -*- coding: utf-8 -*-
"""[오너 2026-09-22 · 그림 45·46] 통합 화면 — 계통도마다 밑그림 칸 · «공통노드 : 도면-도면».

  ⑴ 밑그림 칸은 서버가 보낸 줄 수만큼 난다 — 평면도 · 계통도 1 · 계통도 2 · 기계실.
     이름도 그 줄의 이름을 쓰고, 더한 칸도 켜고 끄는 배선이 같다.
  ⑵ 계통도 칸이 줄면 밑그림 칸도 같이 사라진다(한 장이면 종전과 똑같은 세 칸).
  ⑶ 범례와 노드 카드가 번호가 아니라 «공통노드 : 평면도-계통도1» 로 말한다.
"""
from __future__ import annotations

from pathlib import Path

from test_module_f_viewport_browser import install_page, pw_api, settle

ROOT = Path(__file__).resolve().parents[1]

HOOK = ("  window.__mergeUi = {renderMergeUnderOptions, renderMergeLegend, mgNode};\n"
        '  setStage("open");\n  loadSaved();')


def row(kind, label, note):
    return {"kind": kind, "label": label, "available": True, "source": f"{kind}-source",
            "token": f"{kind}:1", "note": note, "reason": "", "matrix": [1, 0, 0, 1, 0, 0]}


UNDER_TWO = [
    row("plan", "평면도", "원도면의 평면 위치를 공통노드(평면도-계통도1) 높이에 겹칩니다."),
    row("system", "계통도 1", "앞 공통노드(평면도-계통도1)에 수직으로 세우고 "
                           "뒤 공통노드(계통도1-계통도2) 높이에 맞춥니다."),
    row("system2", "계통도 2", "앞 공통노드(계통도1-계통도2)에 수직으로 세우고 "
                            "뒤 공통노드(계통도2-기계실) 높이에 맞춥니다."),
    row("machineroom", "기계실", "기계실 접속점을 기준으로 결합망과 같은 배율로 겹칩니다."),
]
UNDER_ONE = [
    row("plan", "평면도", "원도면의 평면 위치를 공통노드(평면도-계통도) 높이에 겹칩니다."),
    row("system", "계통도", "알람밸브 기준으로 수직 정렬하고 펌프 높이에 맞춥니다."),
    row("machineroom", "기계실", "기계실 접속점을 기준으로 결합망과 같은 배율로 겹칩니다."),
]
NODES = [
    {"label": "10", "x": 0, "y": 0, "e": 0, "part": "system", "anchor": True,
     "joint": "평면도-계통도1", "key": None},
    {"label": "1", "x": 1, "y": 1, "e": -1, "part": "system", "anchor": True,
     "joint": "계통도1-계통도2", "key": None},
    {"label": "s21", "x": 2, "y": 2, "e": -2, "part": "system", "anchor": True,
     "joint": "계통도2-기계실", "key": None},
    {"label": "11", "x": 3, "y": 3, "e": 0, "part": "plan", "key": None},
]
COUNTS = {"plan": 57, "system": 5, "machineroom": 11, "seam": 2,
          "anchor": ["1", "10", "s21"],
          "joints": ["평면도-계통도1", "계통도1-계통도2", "계통도2-기계실"]}


def boxes(page):
    return page.evaluate("""() => [...document.querySelectorAll('.mg-under-options input')]
        .map((el) => [el.id, el.closest('label').querySelector('span').textContent,
                      !!el.onchange])""")


def test_밑그림_칸은_계통도_수를_따르고_공통_노드는_두_도면_이름으로_말한다():
    source = (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();', HOOK)
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        errors = install_page(page, source)

        # ⑴ 계통도 두 장 — 칸도 넷, 이름도 그 줄의 이름, 배선도 칸마다.
        page.evaluate("""(rows) => {
            __mf.mergeView = {nodes: [], pipes: [], underlays: rows};
            document.getElementById('mg-under').checked = true;
            __mergeUi.renderMergeUnderOptions();
        }""", UNDER_TWO)
        settle(page)
        assert boxes(page) == [["mg-under-plan", "평면도", True],
                              ["mg-under-system", "계통도 1", True],
                              ["mg-under-system2", "계통도 2", True],
                              ["mg-under-machineroom", "기계실", True]]
        note = page.inner_text("#mg-under-note")
        assert "계통도 2: 앞 공통노드(계통도1-계통도2)" in note and "기준점" not in note

        # ⑵ 한 장으로 줄면 더한 칸은 사라진다 — 종전 세 칸 그대로.
        page.evaluate("""(rows) => {
            __mf.mergeView = {nodes: [], pipes: [], underlays: rows};
            __mergeUi.renderMergeUnderOptions();
        }""", UNDER_ONE)
        settle(page)
        assert [b[0] for b in boxes(page)] == ["mg-under-plan", "mg-under-system",
                                              "mg-under-machineroom"]
        assert boxes(page)[1][1] == "계통도"

        # ⑶ 범례 · 노드 카드 — 번호가 아니라 두 도면 이름.
        legend = page.evaluate("""(counts) => {
            __mergeUi.renderMergeLegend({counts, iso: true});
            return document.getElementById('mg-legend').textContent;
        }""", COUNTS)
        assert "공통노드 : 평면도-계통도1, 계통도1-계통도2, 계통도2-기계실" in legend
        assert "기준점" not in legend
        title = page.evaluate("""(nodes) => {
            __mf.mergeView = {nodes, pipes: [], underlays: []};
            __mergeUi.mgNode('10');
            return document.getElementById('dg-ins-title').textContent;
        }""", NODES)
        assert "공통노드 : 평면도-계통도1" in title and "기준점" not in title
        assert not errors, errors
        browser.close()
