# -*- coding: utf-8 -*-
"""배관망 완성 진행바(mini 카드) 렌더 확인용 스크린샷."""
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

OUT = BASE / "data" / "_finalize_bar.png"

from playwright.sync_api import sync_playwright  # noqa: E402

with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.goto(f"http://127.0.0.1:{port}/remote30-prototype")
    page.wait_for_timeout(1200)
    page.evaluate("() => { FinalizeProgress.begin(); FinalizeProgress.select(0.42); }")
    page.wait_for_timeout(400)
    box = page.evaluate(
        "() => { const el = document.getElementById('boot-overlay');"
        " const r = el.getBoundingClientRect();"
        " return {mini: el.classList.contains('is-mini'),"
        "         vis: el.classList.contains('is-visible'),"
        "         w: Math.round(r.width), h: Math.round(r.height),"
        "         x: Math.round(r.x), y: Math.round(r.y),"
        "         pct: document.getElementById('boot-pct').textContent,"
        "         bar: document.getElementById('boot-bar').style.width}; }")
    page.screenshot(path=str(OUT))
    br.close()

httpd.shutdown()
print(box)
print(f"screenshot -> {OUT}")
