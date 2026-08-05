# -*- coding: utf-8 -*-
"""PR-3 워크벤치 런타임 검증 — 일회용.

템플릿/JS 는 문법 검사만으로는 회귀를 못 잡는다(함수-지역 헬퍼를 다른 스코프에서
부르면 ReferenceError 가 실행 시점에야 난다). 그래서 실제 브라우저로 띄워
콘솔 오류를 수집한다. 돌고 있는 서버는 건드리지 않고 별도 포트로 띄운다.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5099
DXF = ROOT / "data" / "uploads" / "LH_306.dxf"

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
    print("session:", page.inner_text("#dw-session"))
    print("stepper:", page.inner_text("#dw-stepper").replace("\n", " | "))
    print("design list rows:", page.locator("#dw-design-list input").count())

    page.set_input_files("#dw-dxf", str(DXF))
    page.wait_for_selector("#dw-layer-list .layer-row", timeout=60_000)
    page.wait_for_timeout(1500)
    print("load status:", page.inner_text("#dw-load-status"))
    print("layer rows:", page.locator("#dw-layer-list .layer-row").count())
    print("unclassified note:", page.inner_text("#dw-unclassified-note"))

    # 12종 토글: 한 레이어를 WALL 로 옮기고 그 카테고리 헤더가 생겼는지 본다.
    page.locator("#dw-layer-list select").first.select_option("WALL")
    page.wait_for_timeout(400)
    heads = page.locator("#dw-layer-list .cat-head").all_inner_texts()
    print("WALL 그룹:", [h.replace("\n", " ") for h in heads if h.startswith("WALL")])

    # 카테고리 일괄 끄기
    page.locator('#dw-layer-list input[data-cat="WALL"]').uncheck()
    page.wait_for_timeout(400)
    print("WALL 끈 뒤 체크된 레이어:",
          page.locator("#dw-layer-list .layer-row input:checked").count())

    page.screenshot(path=str(ROOT / "data" / "_dw_verify.png"))
    browser.close()

print("\nJS 오류:", errors or "없음")
sys.exit(1 if errors else 0)
