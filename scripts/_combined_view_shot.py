# -*- coding: utf-8 -*-
"""통합 배관망 뷰(평면/등각/3D)를 실브라우저로 캡처 — 헤드 위치·배색 육안 확인용."""
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

PLANE = BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"
OUT = BASE / "data" / "_shots"
OUT.mkdir(exist_ok=True)

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

    page.set_input_files("#wb-dxf", str(PLANE))
    page.wait_for_function("() => !document.getElementById('wb-run').disabled")
    page.click("#wb-run")
    page.wait_for_function("() => !document.getElementById('wb-finalize').disabled")
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.click("#wb-finalize")
    page.wait_for_function(
        "() => document.getElementById('fx-open-btn').style.display === 'block'")
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.evaluate("""async () => {
      const fd = new FormData();
      fd.append("use_legacy_template", "true");
      fd.append("pump_x", "0"); fd.append("pump_y", "0");
      fd.append("av_x", "0");   fd.append("av_y", "-3000");
      const r = await fetch("/api/remote30/system/extract", {method: "POST", body: fd});
      _revealSystemRiser(await r.json());
    }""")
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.click("#mode-btn-combined")
    page.wait_for_timeout(400)
    page.click("#cmb-merge")
    page.wait_for_function(
        "() => document.getElementById('cmb-merge').textContent.indexOf('통합 중') < 0")
    page.wait_for_timeout(1200)

    def shot(name):
        page.screenshot(path=str(OUT / f"{name}.png"))
        print(f"  → {name}.png")

    for view in ("iso2d", "plan"):
        page.evaluate(f"""() => {{
          state.combined_view = "{view}";
          state.view.fitMode = null; fitCombinedToView(); _renderCombinedMode();
        }}""")
        page.wait_for_timeout(400)
        shot(f"cmb_{view}")

        # 헤드 하나를 골라 6배 확대 — 끝단/밑단 부착 상태 확대 확인
        info = page.evaluate("""() => {
          const g0 = state.combined_geometry;
          const {g, proj} = _combinedDisplayAndProjector(g0, state.combined_view);
          const hl = (g.head_labels || [])[0];
          const n = g.nodes.find(n => n.label === hl);
          if (!n) return null;
          const [px, py] = proj(n.x, n.y, n.z);
          const r = canvas.getBoundingClientRect();
          state.view.zoom *= 8;
          state.view.panX = r.width / 2 - px * state.view.zoom;
          state.view.panY = r.height / 2 + py * state.view.zoom;
          _renderCombinedMode();
          return {head: hl, view: state.combined_view};
        }""")
        page.wait_for_timeout(300)
        print(f"  확대 대상 {info}")
        shot(f"cmb_{view}_zoom")

    page.evaluate("() => setCombinedView && setCombinedView('iso3d')")
    page.wait_for_timeout(1500)
    shot("cmb_3d")
    br.close()

httpd.shutdown()
print(f"JS 오류 {len(errors)}건")
for e in errors[:10]:
    print("  !", e)
