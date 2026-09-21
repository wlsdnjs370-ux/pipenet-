# -*- coding: utf-8 -*-
"""«이어서 열기» 화면 — 원본이 정리된 저장본을 어떻게 보여 주나."""
from __future__ import annotations

import os
import sys

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5065")
PASSWORD = os.environ["LOGIN_PASSWORD"]

bad = []
with sync_playwright() as pw:
    b = pw.chromium.launch()
    page = b.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: bad.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad.append(f"pageerror: {e}"))
    page.goto(f"{BASE}/module-f", wait_until="load")
    if page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=30_000)

    opts = page.eval_on_selector_all(
        "#saved option",
        "es => es.map(o => ({v: o.value, t: o.textContent, gone: o.dataset.gone}))")
    print(f"저장본 {len(opts)}개")
    for o in opts:
        print(f"   {'X' if o['gone'] else 'O'}  {o['t']}")

    gone = [o for o in opts if o["gone"]]
    live = [o for o in opts if not o["gone"]]
    print(f"\n원본 없음 {len(gone)} · 열 수 있음 {len(live)}")

    if gone:
        page.select_option("#saved", gone[0]["v"])
        page.wait_for_timeout(200)
        print("[원본 없는 것 고름]")
        print("   단추 막힘 :", page.eval_on_selector("#btn-reopen", "e => e.disabled"))
        print("   사유 보임 :", not page.eval_on_selector(
            "#resume-why", "e => e.classList.contains('hidden')"))
        print("   사유 문구 :", page.inner_text("#resume-why")[:100])
    if live:
        page.select_option("#saved", live[0]["v"])
        page.wait_for_timeout(200)
        print("[성한 것 고름]")
        print("   단추 막힘 :", page.eval_on_selector("#btn-reopen", "e => e.disabled"))
        print("   사유 숨김 :", page.eval_on_selector(
            "#resume-why", "e => e.classList.contains('hidden')"))
    print("\n콘솔 오류:", bad or "없음")
    b.close()
