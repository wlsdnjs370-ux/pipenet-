# -*- coding: utf-8 -*-
"""PR-5a GATE 결손 폼 런타임 검증 — 일회용.

문법 검사로는 스코프 밖 참조를 못 잡는다. 실도면을 올려 C1 을 돌리고, 결손 폼을
열어 일괄 적용으로 전부 채운 뒤 실제로 게이트가 열리는지까지 본다. 화면의 "남음
0건" 과 서버의 판정이 어긋나면 여기서 422 로 드러난다.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5098
DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    ROOT / "data" / "reference_library" / "2. 고가수조_양주옥정 중상1블럭"
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
    page.select_option("#dw-wall-layers", "ARCHI")
    page.click("#dw-recognize-btn")
    # 인식 성공은 상태줄이 아니라 결손 폼이 열리는 것으로 판정한다 — 캐시 hit 이면
    # 상태줄이 "이전 인식 결과 재사용" 으로 먼저 바뀌어 문구만 보면 일찍 통과한다.
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-open').disabled"
        " || /오류/.test(document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식:", page.inner_text("#dw-recognize-status"))
    print("게이트 진행:", page.inner_text("#dw-gate-progress"))
    print("[결손 항목 채우기] 활성:", page.is_enabled("#dw-gate-open"))

    page.click("#dw-gate-open")
    page.wait_for_selector("#dw-gate-tbody tr")
    print("표 행:", page.locator("#dw-gate-tbody tr").count(),
          "/ 열:", page.locator("#dw-gate-thead th").count())
    print("남음(처음):", page.inner_text("#dw-gate-remain"))
    print("건물사실 칸:", page.locator("#dw-gate-facts .cell").count())
    print("제안 꼬리표:", page.locator("#dw-gate-tbody .sugg").count())
    page.screenshot(path=str(ROOT / "data" / "_dw_gate_open.png"))

    def bulk(field_label, value_label, scope="all"):
        page.select_option("#dw-bulk-field", label=field_label)
        page.select_option("#dw-bulk-value", label=value_label)
        page.select_option("#dw-bulk-scope", scope)
        page.click("#dw-bulk-apply")
        print(f"  일괄 {field_label} = {value_label}: {page.inner_text('#dw-bulk-note')}"
              f" → {page.inner_text('#dw-gate-remain')}")

    # 제안이 하나라도 있으면 "제안값 그대로" 가 나와야 한다(실명 귀속이 붙은 실).
    page.select_option("#dw-bulk-field", label="용도")
    has_suggest = page.locator("#dw-bulk-value option[value='sugg']").count() > 0
    print("제안값 그대로 선택지:", has_suggest)
    if has_suggest:
        bulk("용도", "제안값 그대로")
    bulk("용도", "공동주택")
    bulk("반자 유무", "있음")
    print("반자고 열이 열렸는지(남음 증가):", page.inner_text("#dw-gate-remain"))

    page.select_option("#dw-bulk-field", label="반자고(mm)")
    page.fill("#dw-bulk-value", "2300")
    page.select_option("#dw-bulk-scope", "all")
    page.click("#dw-bulk-apply")
    print("  일괄 반자고 2300:", page.inner_text("#dw-bulk-note"),
          "→", page.inner_text("#dw-gate-remain"))

    page.select_option("#dw-bulk-field", label="천장고(슬래브, mm)")
    page.fill("#dw-bulk-value", "2800")
    page.select_option("#dw-bulk-scope", "all")
    page.click("#dw-bulk-apply")
    print("  일괄 천장고 2800:", page.inner_text("#dw-bulk-note"),
          "→", page.inner_text("#dw-gate-remain"))

    # 건물 사실 + 코어
    # 숫자칸은 사람처럼 치고 초점을 옮긴다 — change 는 blur 때 난다.
    floors = "#dw-gate-facts [data-field='building.floors_total']"
    page.click(floors)
    page.fill(floors, "15")
    page.press(floors, "Tab")
    page.select_option("#dw-gate-facts [data-field='building.structure']", label="내화구조")
    page.select_option("#dw-gate-facts [data-field='building.use']", label="공동주택")
    page.select_option("#dw-gate-facts [data-field='obstacles.status']", label="partial")
    print("건물 사실 4건 →", page.inner_text("#dw-gate-remain"),
          "/ 붉은 사실칸:", page.locator("#dw-gate-facts .cell.miss").count())
    cores = page.locator("#dw-gate-facts [data-field='confirmed']")
    for i in range(cores.count()):
        cores.nth(i).select_option(label="입상관으로 쓴다")
    print("코어 확정:", cores.count(), "개 →", page.inner_text("#dw-gate-remain"))
    print("아직 붉은 칸 — 표:", page.locator("#dw-gate-tbody td.miss").count(),
          "/ 사실:", page.locator("#dw-gate-facts .cell.miss").count(),
          page.locator("#dw-gate-facts .cell.miss").all_inner_texts())

    page.click("#dw-gate-confirm")
    page.wait_for_function(
        "() => !/확정 중/.test(document.getElementById('dw-gate-msg').textContent)",
        timeout=60_000)
    print("확정자 미입력 시:", page.inner_text("#dw-gate-msg"))

    page.fill("#dw-gate-operator", "jinwon")
    page.screenshot(path=str(ROOT / "data" / "_dw_gate_filled.png"))
    page.click("#dw-gate-confirm")
    page.wait_for_function(
        "() => !/확정 중/.test(document.getElementById('dw-gate-msg').textContent)"
        " || !document.getElementById('dw-gate-panel').classList.contains('open')",
        timeout=120_000)
    page.wait_for_timeout(500)
    print("모달 닫힘:", not page.is_visible("#dw-gate-panel"))
    print("게이트 상태:", page.inner_text("#dw-gate-status"))
    print("게이트 메시지:", page.inner_text("#dw-gate-msg"))
    print("stepper:", page.inner_text("#dw-stepper").replace("\n", " | "))
    print("[C2 실행] 활성:", page.is_enabled("#dw-run-c2"))
    page.screenshot(path=str(ROOT / "data" / "_dw_gate_passed.png"))

    page.click("#dw-run-c2")
    page.wait_for_function(
        "() => !/실행 중/.test(document.getElementById('dw-gate-status').textContent)",
        timeout=60_000)
    print("C2 응답:", page.inner_text("#dw-gate-status"))
    browser.close()

# C2 의 501 은 설계대로다(PR-6 미구현). 브라우저는 그것도 console.error 로 적으니
# 네트워크 상태 로그와 실제 JS 예외를 갈라야 종료코드가 의미를 갖는다.
network = [e for e in errors if "Failed to load resource" in e]
js = [e for e in errors if e not in network]
print("\n네트워크 로그(예상됨):", network or "없음")
print("JS 오류:", js or "없음")
sys.exit(1 if js else 0)
