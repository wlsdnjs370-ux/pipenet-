# -*- coding: utf-8 -*-
"""PR-6b C2 기준표·103B 승인 UI 런타임 검증 — 일회용.

여기서 보고 싶은 것은 세 가지다. **기준표가 조문과 함께 뜨는가**, **같은 기준을
다시 굽지 않는가**, **103B 는 세 조건을 다 답하기 전에는 확정 버튼이 열리지
않는가**. 마지막이 특히 중요한데, 버튼이 먼저 열리면 안 물어본 조건이 조용히
'아니오' 로 확정된다.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5096
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
fails: list[str] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'OK ' if ok else 'FAIL'}] {label}{(' — ' + detail) if detail else ''}")
    if not ok:
        fails.append(label)


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
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-open').disabled"
        " || /오류/.test(document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식:", page.inner_text("#dw-recognize-status"))

    # ── 게이트 통과 (PR-5a 검증에서 이미 본 경로라 여기선 통과만 시킨다) ──────
    page.click("#dw-gate-open")
    page.wait_for_selector("#dw-gate-tbody tr")

    def bulk(field_label, value, *, by_label=True):
        page.select_option("#dw-bulk-field", label=field_label)
        if by_label:
            page.select_option("#dw-bulk-value", label=value)
        else:
            page.fill("#dw-bulk-value", value)
        page.select_option("#dw-bulk-scope", "all")
        page.click("#dw-bulk-apply")

    bulk("용도", "공동주택")
    bulk("반자 유무", "있음")
    bulk("반자고(mm)", "2300", by_label=False)
    bulk("천장고(슬래브, mm)", "2800", by_label=False)

    floors = "#dw-gate-facts [data-field='building.floors_total']"
    page.click(floors)
    page.fill(floors, "15")
    page.press(floors, "Tab")
    page.select_option("#dw-gate-facts [data-field='building.structure']", label="내화구조")
    page.select_option("#dw-gate-facts [data-field='building.use']", label="공동주택")
    page.select_option("#dw-gate-facts [data-field='obstacles.status']", label="partial")
    cores = page.locator("#dw-gate-facts [data-field='confirmed']")
    for i in range(cores.count()):
        cores.nth(i).select_option(label="입상관으로 쓴다")
    print("게이트 남음:", page.inner_text("#dw-gate-remain"))

    page.fill("#dw-gate-operator", "jinwon")
    page.click("#dw-gate-confirm")
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-panel').classList.contains('open')"
        " || /오류|남았|확정자/.test(document.getElementById('dw-gate-msg').textContent)",
        timeout=120_000)
    page.wait_for_timeout(300)
    check("게이트 통과", not page.is_visible("#dw-gate-panel"),
          page.inner_text("#dw-gate-msg"))
    check("[C2 실행] 활성", page.is_enabled("#dw-run-c2"))
    check("[기준표 보기] 아직 비활성", not page.is_enabled("#dw-c2-open"))
    print("C2 상태(굽기 전):", page.inner_text("#dw-c2-status"))

    # ── C2 기준표 ──────────────────────────────────────────────────────────
    print("\n[1] C2 기준 굽기")
    page.click("#dw-run-c2")
    page.wait_for_function(
        "() => !/실행 중/.test(document.getElementById('dw-c2-status').textContent)",
        timeout=60_000)
    status1 = page.inner_text("#dw-c2-status")
    print("  상태:", status1)
    check("기준 생성됨", "기준 생성" in status1, status1)
    check("기준표 자동 열림", page.is_visible("#dw-c2-panel"))
    cards = page.locator("#dw-c2-grid .cn-card").count()
    traces = page.locator("#dw-c2-trace tr").count()
    check("기준 카드 렌더", cards >= 10, f"{cards}장")
    check("조문 출처 행 렌더", traces >= 3, f"{traces}행")
    print("  산출물:", page.inner_text("#dw-c2-artifact"))
    print("  배너:", page.inner_text("#dw-c2-banner"))
    page.screenshot(path=str(ROOT / "data" / "_dw_c2_panel.png"))

    page.click("#dw-c2-close")
    check("기준표 닫힘", not page.is_visible("#dw-c2-panel"))
    check("[기준표 보기] 활성화됨", page.is_enabled("#dw-c2-open"))

    # ── 같은 기준을 다시 굽지 않는다 ────────────────────────────────────────
    print("\n[2] 같은 입력으로 재실행")
    page.click("#dw-run-c2")
    page.wait_for_function(
        "() => !/실행 중/.test(document.getElementById('dw-c2-status').textContent)",
        timeout=60_000)
    status2 = page.inner_text("#dw-c2-status")
    print("  상태:", status2)
    check("세대를 늘리지 않음", "다시 굽지 않음" in status2, status2)
    check("constraints.json 그대로",
          "constraints.json" in page.inner_text("#dw-c2-artifact"),
          page.inner_text("#dw-c2-artifact"))
    page.click("#dw-c2-close")

    # ── 103B 승인 ──────────────────────────────────────────────────────────
    print("\n[3] 103B 판단 — 안 물어본 조건은 '아니오' 가 아니다")
    page.click("#dw-esfr-open")
    page.wait_for_selector("#dw-esfr-questions .esfr-q")
    unanswered = page.locator("#dw-esfr-questions .esfr-q.unanswered").count()
    checked = page.locator("#dw-esfr-questions input:checked").count()
    check("세 조건 모두 미답 표시", unanswered == 3, f"{unanswered}건")
    check("기본 선택 없음", checked == 0, f"선택 {checked}건")
    check("확정 버튼 잠김", not page.is_enabled("#dw-esfr-submit"))
    print("  판정:", page.inner_text("#dw-esfr-verdict"))
    page.screenshot(path=str(ROOT / "data" / "_dw_esfr_blank.png"))

    def answer(key, val):
        page.check(f"#dw-esfr-questions input[data-key='{key}'][data-val='{val}']")

    answer("owner_adopts_esfr", "true")
    answer("site_conditions_ok", "true")
    check("두 개만 답해도 버튼은 잠김", not page.is_enabled("#dw-esfr-submit"),
          page.inner_text("#dw-esfr-verdict"))
    answer("commodity_restricted", "false")
    print("  판정:", page.inner_text("#dw-esfr-verdict"))
    check("세 조건 답해도 확정자 없으면 잠김", not page.is_enabled("#dw-esfr-submit"))
    page.fill("#dw-esfr-operator", "jinwon")
    check("확정자까지 채우면 열림", page.is_enabled("#dw-esfr-submit"))
    page.fill("#dw-esfr-note", "발주자 확인 공문 2026-08-01")
    page.screenshot(path=str(ROOT / "data" / "_dw_esfr_filled.png"))

    page.click("#dw-esfr-submit")
    page.wait_for_function(
        "() => !document.getElementById('dw-esfr-panel').classList.contains('open')"
        " || /오류|HTTP|채워야/.test(document.getElementById('dw-esfr-msg').textContent)",
        timeout=60_000)
    page.wait_for_timeout(300)
    check("103B 판단 확정됨", not page.is_visible("#dw-esfr-panel"),
          page.inner_text("#dw-esfr-msg"))
    esfr_status = page.inner_text("#dw-esfr-status")
    print("  103B 상태:", esfr_status)
    check("103B 활성으로 표시", "활성" in esfr_status and "비활성" not in esfr_status,
          esfr_status)
    status3 = page.inner_text("#dw-c2-status")
    print("  C2 상태:", status3)
    check("기준을 함께 다시 구움", "constraints.v2.json" in status3, status3)

    print("\n[4] 무효가 된 NFTC 103 값")
    page.click("#dw-c2-open")
    page.wait_for_selector("#dw-c2-grid .cn-card")
    void_cards = page.locator("#dw-c2-grid .cn-card.void").count()
    check("무효 카드가 붉게 표시", void_cards == 4, f"{void_cards}장")
    banner = page.inner_text("#dw-c2-banner")
    print("  배너:", banner)
    check("배너가 103B 활성을 알림", "103B 활성" in banner, banner)
    trace_fields = page.locator("#dw-c2-trace tr td:first-child").all_inner_texts()
    check("무효 필드는 조문표에서 빠짐",
          "horizontal_distance_m" not in trace_fields,
          ", ".join(trace_fields[:6]))
    page.screenshot(path=str(ROOT / "data" / "_dw_c2_voided.png"))

    browser.close()

network = [e for e in errors if "Failed to load resource" in e]
js = [e for e in errors if e not in network]
print("\n네트워크 로그(예상됨):", network or "없음")
print("JS 오류:", js or "없음")
print("실패한 검사:", fails or "없음")
sys.exit(1 if (js or fails) else 0)
