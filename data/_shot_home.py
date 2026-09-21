# -*- coding: utf-8 -*-
"""초기화면 카드 7장을 눈으로 확인 — 콘솔 오류도 함께 본다."""
import re, sys, urllib.request
from playwright.sync_api import sync_playwright

BASE = "http://127.0.0.1:5051"
html = urllib.request.urlopen(f"{BASE}/login", timeout=10).read().decode("utf-8", "replace")
PW = re.search(r'<span class="pw-value">([^<]+)<', html).group(1)

bad = []
with sync_playwright() as pw:
    b = pw.chromium.launch()
    page = b.new_page(viewport={"width": 1600, "height": 1100})
    page.on("console", lambda m: bad.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad.append(f"pageerror: {e}"))
    page.goto(f"{BASE}/", wait_until="load")
    if page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PW)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    cards = page.eval_on_selector_all(
        ".home-module-btn",
        "els => els.map(e => [e.querySelector('.module-chip').textContent,"
        " e.querySelector('strong').textContent, e.getAttribute('href')])")
    for c in cards:
        print(f"   {c[0]:<10} {c[1][:44]:<46} {c[2]}")
    if len(cards) != 7:
        bad.append(f"카드가 7장이 아니다: {len(cards)}")
    g = [c for c in cards if c[0].strip() == "Module G"]
    if not g or "/module-g-cad-editor" not in g[0][2]:
        bad.append(f"Module G 카드/링크 이상: {g}")
    page.locator(".home-grid").screenshot(path="data/_mf_shots/home_7cards.png")
    b.close()

if bad:
    print("\n!! 문제:")
    for x in bad:
        print("  -", x)
    sys.exit(1)
print("\nOK: home cards fine, console errors 0")
