# -*- coding: utf-8 -*-
"""FX 검토 모달이 배관망 완성 시 자동으로 뜨지 않는지 실제 브라우저로 확인.

로그인은 건드리지 않는다 — Flask 템플릿만 렌더해 저장소 루트에 떨군 뒤,
같은 루트를 http.server 로 띄워 /static 자산까지 정상 로드되게 한다.
"""
from __future__ import annotations

import functools
import http.server
import socketserver
import sys
import threading
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

from flask import Flask, render_template  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

app = Flask(__name__, template_folder=str(BASE / "templates"),
            static_folder=str(BASE / "static"))
with app.test_request_context():
    html = render_template("remote30_prototype.html")

out = BASE / "data" / "_fx_modal_check.html"
out.write_text(html, encoding="utf-8")

handler = functools.partial(http.server.SimpleHTTPRequestHandler,
                            directory=str(BASE))
httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
httpd.RequestHandlerClass.log_message = lambda *a, **k: None
port = httpd.server_address[1]
threading.Thread(target=httpd.serve_forever, daemon=True).start()

FX = {
    "profiles": {"평균": {"eq_len_m": 3.0}, "FX_20A_216": {"eq_len_m": 2.16}},
    "default_profile": "평균",
    "equipment": [
        {"label": "FX-1", "desc": "테스트", "in": "N1", "out": "N2",
         "pipe": "P1", "spec_ref": "평균", "eq_len": 3.0, "source": "extracted",
         "drawing_len_mm": 1200.0},
    ],
}
TABLES = {"nodes": [{"label": "N1", "x": 0, "y": 0}, {"label": "N2", "x": 10, "y": 0}],
          "nozzles": [{"in": "N1"}]}

errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page()
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.goto(f"http://127.0.0.1:{port}/data/_fx_modal_check.html")
    page.wait_for_timeout(1500)

    before = page.evaluate(
        "() => document.getElementById('fx-modal-overlay').className")
    page.evaluate("([fx, t]) => FxReview.open(fx, t)", [FX, TABLES])
    page.wait_for_timeout(300)
    after = page.evaluate("""() => ({
        overlay: document.getElementById('fx-modal-overlay').className,
        btn: getComputedStyle(document.getElementById('fx-open-btn')).display,
        rows: document.querySelectorAll('#fx-table tr.fx-row').length,
    })""")
    page.click("#fx-open-btn")
    page.wait_for_timeout(300)
    opened = page.evaluate(
        "() => document.getElementById('fx-modal-overlay').className")
    page.click("#fx-modal-close")
    page.wait_for_timeout(300)
    closed = page.evaluate(
        "() => document.getElementById('fx-modal-overlay').className")
    br.close()

httpd.shutdown()
out.unlink(missing_ok=True)

print(f"로드 직후 overlay      : {before!r}")
print(f"open() 후 overlay      : {after['overlay']!r}  (자동 오픈 없어야 함)")
print(f"open() 후 버튼 display : {after['btn']}  · 표 행 {after['rows']}")
print(f"버튼 클릭 후 overlay   : {opened!r}")
print(f"닫기 후 overlay        : {closed!r}")

ok = ("is-open" not in after["overlay"] and after["btn"] == "block"
      and after["rows"] == 1 and "is-open" in opened
      and "is-open" not in closed)
if errors:
    print("\nJS 오류:")
    for e in errors[:20]:
        print("  " + e)
print("\n" + ("PASS — 자동 오픈 안 됨, 버튼으로만 열림" if ok and not errors
             else "FAIL"))
sys.exit(0 if (ok and not errors) else 1)
