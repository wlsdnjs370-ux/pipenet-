# -*- coding: utf-8 -*-
"""disabled 버튼이 화면상 활성 버튼과 구분되는지 실측.

.run-btn:disabled 는 회색으로 죽이지만, 인라인 style 로 흰/보라 아웃라인을 준
버튼들은 인라인이 이겨서 disabled 여도 멀쩡해 보인다 — 그 가설을 잰다.
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
out = BASE / "data" / "_dis_probe.html"
with app.test_request_context():
    out.write_text(render_template("remote30_prototype.html"), encoding="utf-8")

handler = functools.partial(http.server.SimpleHTTPRequestHandler,
                            directory=str(BASE))
httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
httpd.RequestHandlerClass.log_message = lambda *a, **k: None
port = httpd.server_address[1]
threading.Thread(target=httpd.serve_forever, daemon=True).start()

IDS = ["wb-run", "sys-extract", "mr-extract", "mr-emit", "cmb-merge",
       "cmb-exp-sdf", "cmb-exp-kfp"]

with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.goto(f"http://127.0.0.1:{port}/data/_dis_probe.html")
    page.wait_for_timeout(1000)
    rows = page.evaluate("""(ids) => ids.map(id => {
      const el = document.getElementById(id);
      if (!el) return {id, missing: true};
      const cs = getComputedStyle(el);
      return {id, disabled: el.disabled, bg: cs.backgroundColor,
              fg: cs.color, cursor: cs.cursor};
    })""", IDS)
    br.close()

httpd.shutdown()
out.unlink(missing_ok=True)

GREY = "rgb(203, 213, 225)"   # .run-btn:disabled 배경 #cbd5e1

print(f"{'id':14s} {'disabled':9s} {'배경':22s} {'글자':22s} {'커서':12s} 회색표시")
bad = []
for r in rows:
    if r.get("missing"):
        print(f"{r['id']:14s} (없음)")
        continue
    grey = r["bg"] == GREY
    print(f"{r['id']:14s} {str(r['disabled']):9s} {r['bg']:22s} {r['fg']:22s} "
          f"{r['cursor']:12s} {'예' if grey else '아니오 ← 구분불가'}")
    if r["disabled"] and not grey:
        bad.append(r["id"])

print(f"\ndisabled 인데 회색으로 안 죽는 버튼 {len(bad)}개: {', '.join(bad) or '없음'}")
sys.exit(1 if bad else 0)
