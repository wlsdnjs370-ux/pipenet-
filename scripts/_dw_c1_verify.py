# -*- coding: utf-8 -*-
"""PR-4e C1 인식 배선 런타임 검증 — 일회용.

문법 검사로는 스코프 밖 참조를 못 잡는다. 실제 브라우저로 도면을 올리고
[C1 인식 시작] 까지 눌러 콘솔 오류와 오버레이 개수를 본다. 돌고 있는 서버는
건드리지 않고 별도 포트로 띄운다.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5099
DXF = (ROOT / "data" / "reference_library" / "2. 고가수조_양주옥정 중상1블럭"
       / "CAD" / "XR" / "XR-단위세대 평면도 (공동주택).dxf")

os.environ["DESIGN_WORKBENCH_ENABLED"] = "1"
os.environ["LOGIN_PASSWORD"] = "dwverify"
sys.path.insert(0, str(ROOT))

spec = importlib.util.spec_from_file_location("daejo_server", ROOT / "대조 서버.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

threading.Thread(
    target=lambda: mod.app.run(host="127.0.0.1", port=PORT, threaded=True),
    daemon=True,
).start()

from playwright.sync_api import sync_playwright  # noqa: E402

errors: list[str] = []
with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1600, "height": 950})
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))

    base = f"http://127.0.0.1:{PORT}"
    page.goto(f"{base}/login", wait_until="domcontentloaded")
    page.fill("input[type=password]", "dwverify")
    page.press("input[type=password]", "Enter")
    page.wait_for_load_state("networkidle")

    page.goto(f"{base}/design-workbench", wait_until="networkidle")
    page.set_input_files("#dw-dxf", str(DXF))
    page.wait_for_selector("#dw-layer-list .layer-row", timeout=120_000)
    page.wait_for_timeout(1000)
    print("load status:", page.inner_text("#dw-load-status"))
    print("WALL 후보칸 보임:", page.is_visible("#dw-wall-field"),
          "/ option", page.locator("#dw-wall-layers option").count())
    print("버튼 활성:", page.is_enabled("#dw-recognize-btn"))

    page.select_option("#dw-wall-layers", "ARCHI")
    page.click("#dw-recognize-btn")
    # 인식은 실도면 기준 10초 내외. 상태줄이 결과 문구로 바뀔 때까지 기다린다.
    page.wait_for_function(
        "() => !/인식 중|도면 읽는 중/.test("
        "document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식 상태:", page.inner_text("#dw-recognize-status"))
    print("단계 행:", page.locator("#dw-recognize-stages .stage-row").count())
    print("stages:", " | ".join(
        t.replace("\n", " ") for t in
        page.locator("#dw-recognize-stages .stage-row").all_inner_texts()))
    print("stepper:", page.inner_text("#dw-stepper").replace("\n", " | "))
    print("설계 산출물:", page.inner_text("#dw-design-list").replace("\n", " | "))
    page.screenshot(path=str(ROOT / "data" / "_dw_c1_verify.png"))

    # 두 번째 실행 — 캐시 hit 경로에서도 화면이 비지 않는지.
    page.click("#dw-recognize-btn")
    page.wait_for_function(
        "() => !/인식 중|도면 읽는 중|이전 인식/.test("
        "document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("2회차 상태:", page.inner_text("#dw-recognize-status"))
    page.screenshot(path=str(ROOT / "data" / "_dw_c1_verify_cached.png"))
    browser.close()

print("\nJS 오류:", errors or "없음")
sys.exit(1 if errors else 0)
