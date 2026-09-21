# -*- coding: utf-8 -*-
"""모듈 D 출력 워크벤치 — 실제 브라우저 런타임 검증.

콘솔 오류를 한 건이라도 잡으면 실패로 본다 (함수-지역 헬퍼 타 스코프 참조 전례).
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
SUB = ROOT / "routes" / "제출용[최종]"
SDF, XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"
SHOT = ROOT / "data" / "_dout_shots"
SHOT.mkdir(exist_ok=True)
BASE = "http://127.0.0.1:5057"
PASSWORD = os.environ["LOGIN_PASSWORD"]

problems: list[str] = []

with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1600, "height": 1000})
    page.on("console", lambda m: problems.append(f"console.{m.type}: {m.text}")
            if m.type in ("error", "warning") else None)
    page.on("pageerror", lambda e: problems.append(f"pageerror: {e}"))

    page.goto(f"{BASE}/module-d-output", wait_until="load")
    print("로그인 페이지로 넘어갔나:", page.url)
    page.fill("input[type=password]", PASSWORD)
    page.click("button[type=submit]")
    page.wait_for_load_state("load")
    print("로그인 후:", page.url)
    if "/module-d-output" not in page.url:
        page.goto(f"{BASE}/module-d-output", wait_until="load")

    page.set_input_files("#sdf-input", str(SDF))
    page.set_input_files("#xml-input", str(XML))
    print("파일 라벨:", page.inner_text("#sdf-label"), "/", page.inner_text("#xml-label"))

    page.click("#btn-load")
    page.wait_for_function("document.querySelector('#load-status').textContent.startsWith('✓')",
                           timeout=90_000)
    print("읽기:", page.inner_text("#load-status"))
    print("요약: 관로", page.inner_text("#sum-pipes"), "노드", page.inner_text("#sum-nodes"),
          "노즐", page.inner_text("#sum-nozzles"), "결과", page.inner_text("#sum-results"))
    print("결합 상자:", page.inner_text("#join-box").replace("\n", " | ")[:200])
    print("프리셋 버튼:", page.eval_on_selector_all(
        "#preset-row button", "els => els.map(e => e.textContent)"))
    print("관로 선택지", page.eval_on_selector_all("#link-item option", "e => e.length"),
          "· 기본값", page.input_value("#link-item"),
          "/ 노드", page.eval_on_selector_all("#node-item option", "e => e.length"),
          "기본값", page.input_value("#node-item"))
    page.screenshot(path=str(SHOT / "1_loaded.png"))

    # 표시 항목 변경 → 재계산 없이 다시 그린다.
    page.select_option("#link-item", "Pipe volumetric flow")
    page.select_option("#node-item", "Node pressure")
    page.click("#btn-draw")
    page.wait_for_function("document.querySelector('#out-status').textContent.startsWith('✓')",
                           timeout=120_000)
    print("도면 1:", page.inner_text("#out-status"), "|", page.inner_text("#overlay-info"))
    shown = page.eval_on_selector("#preview-img",
                                  "e => [e.style.display, e.naturalWidth, e.naturalHeight]")
    print("미리보기 표시/폭/높이:", shown)
    if shown[0] != "block" or not shown[1]:
        problems.append(f"미리보기가 그려지지 않았다: {shown}")
    page.screenshot(path=str(SHOT / "2_drawn_flow.png"))

    # 프리셋 버튼 — 드롭다운을 바꾸고 곧바로 다시 그린다.
    page.click("#preset-row button:has-text('압력본_옥내소화전')")
    page.wait_for_function(
        "document.querySelector('#out-status').textContent.startsWith('✓')", timeout=120_000)
    print("프리셋 후 드롭다운:", page.input_value("#link-item"), "/", page.input_value("#node-item"))
    print("도면 2:", page.inner_text("#out-status"), "|", page.inner_text("#overlay-info"))
    page.screenshot(path=str(SHOT / "3_preset_bore.png"))

    # 이름표를 끄고 표시 항목도 None 이면 글자가 하나도 남지 않아야 한다.
    page.uncheck("#sw-labels")
    page.select_option("#link-item", "None")
    page.select_option("#node-item", "None")
    page.click("#btn-draw")
    page.wait_for_function(
        "document.querySelector('#out-status').textContent.startsWith('✓')", timeout=120_000)
    print("이름표·표시항목 모두 끈 뒤:", page.inner_text("#overlay-info"))
    if "라벨 0" not in page.inner_text("#overlay-info"):
        problems.append("이름표를 끄고 표시 항목도 None 인데 글자가 남았다")
    page.check("#sw-labels")

    page.click("#btn-report")
    page.wait_for_function(
        "document.querySelector('#out-status').textContent.startsWith('✓')", timeout=180_000)
    print("계산서:", page.inner_text("#out-status"))
    print("리포트 상자:", page.inner_text("#out-box").replace("\n", " | ")[:200])
    page.screenshot(path=str(SHOT / "4_report.png"))

    page.click("#btn-book")
    page.wait_for_function(
        "document.querySelector('#out-status').textContent.startsWith('✓')", timeout=300_000)
    print("합본:", page.inner_text("#out-status"))
    print("합본 상자:", page.inner_text("#out-box").replace("\n", " | ")[:300])
    page.screenshot(path=str(SHOT / "5_book.png"))

    links = page.eval_on_selector_all(
        ".overlay-bar .actions a",
        "els => els.filter(e => !e.classList.contains('hidden')).map(e => [e.textContent, e.href])")
    print("내려받기 링크:", links)
    for label, href in links:
        got = page.request.get(href)
        head = got.body()[:5]
        print(f"  {label} → {got.status} {len(got.body())} bytes {head}")
        if got.status != 200 or head != b"%PDF-":
            problems.append(f"{label} 내려받기 실패: {got.status}")

    browser.close()

print()
if problems:
    print("문제", len(problems), "건")
    for p in problems:
        print("  ×", p)
    sys.exit(1)
print("콘솔 오류 0 · 링크 전부 PDF — 통과")
