# -*- coding: utf-8 -*-
"""운영 5051 의 모듈 F 초기화면만 확인 — 콘솔오류 0 + 헤드 종류 그림 3장."""
import re
import sys
import urllib.request

from playwright.sync_api import sync_playwright

BASE = "http://127.0.0.1:5051"
html = urllib.request.urlopen(f"{BASE}/login", timeout=10).read().decode("utf-8", "replace")
m = re.search(r'<span class="pw-value">([^<]+)<', html)
if not m:
    sys.exit("로그인 힌트를 못 읽었습니다(루프백이 아닐 수 있음)")
PW = m.group(1)

problems = []
with sync_playwright() as pw:
    b = pw.chromium.launch()
    page = b.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda mm: problems.append(f"console.{mm.type}: {mm.text}")
            if mm.type == "error" else None)
    page.on("pageerror", lambda e: problems.append(f"pageerror: {e}"))
    page.goto(f"{BASE}/module-f", wait_until="load")
    if page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PW)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    print("url:", page.url)
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=20_000)
    figs = page.evaluate("""() => Array.from(
      document.querySelectorAll('.kinds .ekind img.diagram'),
      im => [im.alt, im.naturalWidth])""")
    print("헤드 종류 그림:", figs)
    if len(figs) != 3 or any(w < 10 for _a, w in figs):
        problems.append(f"헤드 종류 그림이 안 떴다: {figs}")
    btns = page.eval_on_selector_all(
        "#ed-aj-scan, #ed-aj-apply, #ed-aj-clear", "e => e.length")
    print("자동 이음 단추:", btns)
    if btns != 3:
        problems.append(f"자동 이음 단추가 3개가 아니다 ({btns})")
    print("저장된 찍기:", page.eval_on_selector_all("#saved option", "e => e.length"))
    page.screenshot(path="data/_mf_shots/prod_module_f.png")
    b.close()

if problems:
    print("\n!! 문제:")
    for p in problems:
        print("  -", p)
    sys.exit(1)
print("\n운영 모듈 F 초기화면 정상 — 콘솔오류 0")
