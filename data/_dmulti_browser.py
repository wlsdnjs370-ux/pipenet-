# -*- coding: utf-8 -*-
"""모듈 D 표시 항목 다중 선택 — 실제 브라우저 런타임 검증.

콘솔 오류/경고를 한 건이라도 잡으면 실패로 본다.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
SUB = ROOT / "routes" / "제출용[최종]"
SDF, XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"
SHOT = ROOT / "data" / "_dmulti_shots"
SHOT.mkdir(exist_ok=True)
BASE = "http://127.0.0.1:5057"
PASSWORD = os.environ["LOGIN_PASSWORD"]

problems: list[str] = []


def chips(page, host):
    return page.eval_on_selector_all(
        f"#{host} .chip",
        "els => els.map(e => [e.dataset.name, e.getAttribute('aria-pressed')])")


def pressed(page, host):
    return [n for n, p in chips(page, host) if p == "true"]


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1600, "height": 1000})
    page.on("console", lambda m: problems.append(f"console.{m.type}: {m.text}")
            if m.type in ("error", "warning") else None)
    page.on("pageerror", lambda e: problems.append(f"pageerror: {e}"))

    page.goto(f"{BASE}/module-d-output", wait_until="load")
    page.fill("input[type=password]", PASSWORD)
    page.click("button[type=submit]")
    page.wait_for_load_state("load")
    if "/module-d-output" not in page.url:
        page.goto(f"{BASE}/module-d-output", wait_until="load")

    page.set_input_files("#sdf-input", str(SDF))
    page.set_input_files("#xml-input", str(XML))
    page.click("#btn-load")
    page.wait_for_function("document.querySelector('#load-status').textContent.startsWith('✓')",
                           timeout=90_000)
    print("읽기:", page.inner_text("#load-status"))
    print("관로 칩", len(chips(page, "link-chips")), "· 눌린 것", pressed(page, "link-chips"))
    print("노드 칩", len(chips(page, "node-chips")), "· 눌린 것", pressed(page, "node-chips"))
    print("장 수 안내:", page.inner_text("#pick-count"))
    page.screenshot(path=str(SHOT / "1_chips.png"))

    # ── 클릭/비클릭 토글 ── 처음 눌려 있는 것을 빼고, 두 가지를 새로 고른다.
    page.click('#link-chips .chip[data-name="Pipe velocity"]')
    if "Pipe velocity" in pressed(page, "link-chips"):
        problems.append("눌린 칩을 다시 눌렀는데 빠지지 않았다")
    page.click('#link-chips .chip[data-name="Pipe volumetric flow"]')
    page.click('#link-chips .chip[data-name="Pipe bore"]')
    if set(pressed(page, "link-chips")) != {"Pipe volumetric flow", "Pipe bore"}:
        problems.append(f"고른 칩이 다르다: {pressed(page, 'link-chips')}")
    page.click('#node-chips .chip[data-name="None"]')
    page.click('#node-chips .chip[data-name="Node pressure"]')
    print("고른 뒤 관로:", pressed(page, "link-chips"), "노드:", pressed(page, "node-chips"))
    print("장 수 안내:", page.inner_text("#pick-count"))

    # ── 여러 장 생성 ──
    page.click("#btn-draw")
    page.wait_for_function("document.querySelector('#out-status').textContent.startsWith('✓')",
                           timeout=180_000)
    print("생성:", page.inner_text("#out-status"))
    tabs = page.eval_on_selector_all(
        "#sheet-tabs .sheet-tab", "els => els.map(e => [e.textContent, e.getAttribute('aria-pressed')])")
    print("탭:", tabs)
    if len(tabs) != 2:
        problems.append(f"관로 2 × 노드 1 인데 탭이 {len(tabs)}개다")
    if tabs and tabs[0][1] != "true":
        problems.append("첫 장이 눌린 상태로 보이지 않는다")
    page.wait_for_function("document.querySelector('#preview-img').naturalWidth > 0",
                           timeout=30_000)
    first = page.eval_on_selector("#preview-img", "e => [e.src, e.naturalWidth]")
    print("미리보기 1:", first[0].split("/")[-1], first[1])
    print("오버레이:", page.inner_text("#overlay-info"))
    page.screenshot(path=str(SHOT / "2_sheet1.png"))

    # ── 탭 전환 ──
    page.click("#sheet-tabs .sheet-tab:nth-child(2)")
    page.wait_for_function("document.querySelector('#preview-img').naturalWidth > 0",
                           timeout=30_000)
    second = page.eval_on_selector("#preview-img", "e => [e.src, e.naturalWidth]")
    if not second[1]:
        problems.append("두 번째 장 미리보기가 그려지지 않았다")
    print("미리보기 2:", second[0].split("/")[-1], second[1])
    print("오버레이:", page.inner_text("#overlay-info"))
    if first[0] == second[0]:
        problems.append("탭을 바꿨는데 미리보기가 그대로다")
    page.screenshot(path=str(SHOT / "3_sheet2.png"))

    # ── 낱장 + 합본 링크 ──
    links = page.eval_on_selector_all(
        ".overlay-bar .actions a",
        "els => els.filter(e => !e.classList.contains('hidden')).map(e => [e.textContent, e.href])")
    print("내려받기 링크:", [l[0] for l in links])
    if not any(t == "여러 장 PDF" for t, _ in links):
        problems.append("두 장을 그렸는데 합본 링크가 없다")
    for label, href in links:
        got = page.request.get(href)
        head = got.body()[:5]
        print(f"  {label} → {got.status} {len(got.body())} bytes {head}")
        if got.status != 200 or head != b"%PDF-":
            problems.append(f"{label} 내려받기 실패: {got.status}")

    # ── 프리셋 버튼 (칩을 갈아끼우고 곧바로 다시 그린다) ──
    page.click("#preset-row button:has-text('압력본_옥내소화전')")
    page.wait_for_function("document.querySelector('#out-status').textContent.startsWith('✓')",
                           timeout=180_000)
    print("프리셋 후 관로:", pressed(page, "link-chips"), "노드:", pressed(page, "node-chips"))
    print("프리셋 후 탭:", page.eval_on_selector_all("#sheet-tabs .sheet-tab", "e => e.length"))
    print("프리셋 생성:", page.inner_text("#out-status"))
    if len(pressed(page, "link-chips")) != 1:
        problems.append("프리셋이 칩 한 가지만 남기지 못했다")
    page.screenshot(path=str(SHOT / "4_preset.png"))

    # ── 상한 초과는 미리 막는다 ──
    for name, _p in chips(page, "link-chips"):
        if name not in pressed(page, "link-chips"):
            page.click(f'#link-chips .chip[data-name="{name}"]')
    for name, _p in chips(page, "node-chips"):
        if name not in pressed(page, "node-chips"):
            page.click(f'#node-chips .chip[data-name="{name}"]')
    over = page.inner_text("#pick-count")
    disabled = page.eval_on_selector("#btn-draw", "e => e.disabled")
    print("전부 고른 뒤:", over, "· 생성 버튼 잠김:", disabled)
    if not disabled:
        problems.append("24장을 넘겼는데 생성 버튼이 열려 있다")
    page.screenshot(path=str(SHOT / "5_toomany.png"))

    browser.close()

print()
if problems:
    print("문제", len(problems), "건")
    for p in problems:
        print("  ×", p)
    sys.exit(1)
print("콘솔 오류 0 · 다중 선택/탭/합본/상한 전부 통과")
