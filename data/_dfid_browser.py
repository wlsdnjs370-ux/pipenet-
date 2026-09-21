# -*- coding: utf-8 -*-
"""모듈 D — 원본 SDF 표시 설정 따르기(DF1) 실브라우저 검증.

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
SHOT = ROOT / "data" / "_dfid_shots"
SHOT.mkdir(exist_ok=True)
BASE = os.environ.get("PROBE_BASE", "http://127.0.0.1:5058")
PASSWORD = os.environ["LOGIN_PASSWORD"]

problems: list[str] = []


def switches(page):
    return {sid: page.eval_on_selector(f"#{sid}", "e => e.checked")
            for sid in ("sw-link-labels", "sw-node-labels", "sw-arrows")}


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

    before = switches(page)
    print("읽기 전 스위치:", before)
    if not all(before.values()):
        problems.append(f"파일을 읽기 전인데 스위치가 켜져 있지 않다: {before}")

    page.set_input_files("#sdf-input", str(SDF))
    page.set_input_files("#xml-input", str(XML))
    page.click("#btn-load")
    page.wait_for_function("document.querySelector('#load-status').textContent.startsWith('✓')",
                           timeout=90_000)
    print("읽기:", page.inner_text("#load-status"))

    after = switches(page)
    print("읽은 뒤 스위치:", after)
    print("안내:", page.inner_text("#source-display"))
    # 자동본 SDF: link-labels='0' node-labels='0' flow-arrows='1'
    want = {"sw-link-labels": False, "sw-node-labels": False, "sw-arrows": True}
    if after != want:
        problems.append(f"원본 설정을 따르지 않았다 — 기대 {want} 실측 {after}")
    if "원본 표시 설정" not in page.inner_text("#source-display"):
        problems.append("무엇을 따랐는지 화면에 적히지 않았다")
    page.screenshot(path=str(SHOT / "1_seeded.png"))

    # ── 원본 그대로 한 장 ──
    page.click('#link-chips .chip[data-name="Pipe velocity"]')
    page.click('#node-chips .chip[data-name="None"]')
    page.click("#btn-draw")
    page.wait_for_function("document.querySelector('#out-status').textContent.startsWith('✓')",
                           timeout=180_000)
    print("원본대로 생성:", page.inner_text("#out-status"))
    page.wait_for_function("document.querySelector('#preview-img').naturalWidth > 0",
                           timeout=30_000)
    asis = page.eval_on_selector("#preview-img", "e => e.src")
    page.screenshot(path=str(SHOT / "2_asis.png"))

    # ── 사람이 이름표를 켜면 그 말을 듣는다 ──
    page.check("#sw-link-labels")
    page.check("#sw-node-labels")
    page.click("#btn-draw")
    page.wait_for_function("document.querySelector('#out-status').textContent.startsWith('✓')",
                           timeout=180_000)
    page.wait_for_function("document.querySelector('#preview-img').naturalWidth > 0",
                           timeout=30_000)
    tagged = page.eval_on_selector("#preview-img", "e => e.src")
    print("이름표 켜고 생성:", page.inner_text("#out-status"))
    if asis == tagged:
        problems.append("이름표를 켰는데 같은 그림이 돌아왔다")
    page.screenshot(path=str(SHOT / "3_tagged.png"))

    # 두 장을 실제로 내려받아 크기를 견준다 — 글자가 늘면 파일도 커진다.
    sizes = {}
    for tag, src in (("원본대로", asis), ("이름표", tagged)):
        got = page.request.get(src.replace(".png", ".pdf"))
        sizes[tag] = len(got.body())
        if got.status != 200 or got.body()[:5] != b"%PDF-":
            problems.append(f"{tag} PDF 내려받기 실패: {got.status}")
    print("PDF 크기:", sizes)
    if sizes.get("이름표", 0) <= sizes.get("원본대로", 0):
        problems.append(f"이름표를 켠 쪽이 더 크지 않다: {sizes}")

    browser.close()

print()
if problems:
    print("문제", len(problems), "건")
    for p in problems:
        print("  ×", p)
    sys.exit(1)
print("콘솔 오류 0 · 원본 설정 따르기/사람 지시 우선 전부 통과")
