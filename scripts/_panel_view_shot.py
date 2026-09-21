# -*- coding: utf-8 -*-
"""계통도·기계실 패널 캔버스 캡처 — 배색 육안 확인용."""
from __future__ import annotations

import sys
import threading
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import importlib  # noqa: E402

srv = importlib.import_module("대조 서버")
app = srv.app
app.before_request_funcs[None] = [
    f for f in app.before_request_funcs.get(None, [])
    if f.__name__ != "_require_login_gate"
]

from werkzeug.serving import make_server  # noqa: E402

httpd = make_server("127.0.0.1", 0, app, threaded=True)
port = httpd.server_port
threading.Thread(target=httpd.serve_forever, daemon=True).start()

SP = BASE / "data" / "sample_problem"
OUT = BASE / "data" / "_shots"
OUT.mkdir(exist_ok=True)
PANELS = [
    ("system", "#mode-btn-system", "#sys-dxf", SP / "대명동201동 계통도.dxf"),
    ("machineroom", "#mode-btn-machineroom", "#mr-dxf", SP / "201동 기계실.dxf"),
]

from playwright.sync_api import sync_playwright  # noqa: E402

errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.set_default_timeout(300000)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.on("console", lambda m: errors.append(f"console.error: {m.text}")
            if m.type == "error" else None)
    page.goto(f"http://127.0.0.1:{port}/remote30-prototype")
    page.wait_for_timeout(1500)

    for name, btn, inp, dxf in PANELS:
        page.click(btn)
        page.wait_for_timeout(300)
        page.set_input_files(inp, str(dxf))
        page.wait_for_timeout(25000)
        page.evaluate("() => animQueue.flush && animQueue.flush()")
        page.wait_for_timeout(1500)
        page.screenshot(path=str(OUT / f"panel_{name}.png"))
        print(f"  → panel_{name}.png")
    br.close()

httpd.shutdown()
print(f"JS 오류 {len(errors)}건")
for e in errors[:10]:
    print("  !", e)
