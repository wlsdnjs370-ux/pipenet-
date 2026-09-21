# -*- coding: utf-8 -*-
"""[F-12] 검증이 손질 진입에서 멈춘 이유를 본다 — 어느 패널이 켜지나."""
from __future__ import annotations

import os
import sys

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5065")
PASSWORD = os.environ["LOGIN_PASSWORD"]
KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")

with sync_playwright() as pw:
    b = pw.chromium.launch()
    page = b.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: print("   console." + m.type + ":", m.text)
            if m.type == "error" else None)
    page.on("pageerror", lambda e: print("   pageerror:", e))
    page.goto(f"{BASE}/module-f", wait_until="load")
    if page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=30_000)
    page.select_option("#saved", KEY)
    page.click("#btn-reopen")
    for i in range(40):
        page.wait_for_timeout(3000)
        info = page.evaluate("""() => {
          const on = [...document.querySelectorAll('section.card')]
            .filter(s => !s.classList.contains('hidden')).map(s => s.id);
          const steps = [...document.querySelectorAll('.steps div')]
            .map(e => e.id + (e.classList.contains('on') ? '*' : ''));
          return {on, steps,
                  status: (document.querySelector('#status')||{}).textContent,
                  job: (document.querySelector('#job-line')||{}).textContent,
                  busy: !document.querySelector('#busy').classList.contains('hidden')};
        }""")
        print(f"[{i * 3:3d}s] 패널 {info['on']}")
        print(f"       단계 {info['steps']}")
        print(f"       상태 {(info['status'] or '')[:80]}")
        print(f"       잡   {(info['job'] or '')[:80]} · busy={info['busy']}")
        if "panel-edit" in info["on"]:
            print("   → panel-edit 켜짐")
            break
    b.close()
