# -*- coding: utf-8 -*-
"""PR-B 검증 — remote30 워크벤치가 ARC/HATCH/SOLID 를 실제로 그리는지 브라우저로 확인.

서버는 {"t":"A"|"H"|"S"} 를 방출하는데 클라이언트 drawEntity 는 "ARC" 를 찾고
H/S 분기가 없었다. 구문검사로는 안 잡히므로(모두 유효한 JS) 실렌더로만 검출된다.
캔버스 2D 컨텍스트의 arc/fill 호출 수를 렌더 1회 기준으로 세어 판정한다.
"""
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
print(f"서버 :{port}")

DXF = BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"
if not DXF.is_file():
    print(f"도면 없음: {DXF}")
    sys.exit(2)

# 이 도면의 방출 실측치 (scripts/_prb_entity_hist.py)
EXPECT = {"A": 6090, "C": 671, "H": 145, "S": 180, "I": 388}

from playwright.sync_api import sync_playwright  # noqa: E402

fails: list[str] = []


def check(label, got, want):
    ok = got == want
    print(f"  {'OK  ' if ok else 'FAIL'} {label}: {got!r}" + ("" if ok else f"  (기대 {want!r})"))
    if not ok:
        fails.append(label)


errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.set_default_timeout(300000)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.on("console", lambda m: errors.append(f"console.error: {m.text}") if m.type == "error" else None)

    page.goto(f"http://127.0.0.1:{port}/remote30-workbench")
    page.add_init_script("""
      window.__prb = { arc: 0, fill: 0, stroke: 0 };
      const P = CanvasRenderingContext2D.prototype;
      for (const m of ["arc", "fill", "stroke"]) {
        const orig = P[m];
        P[m] = function (...a) { window.__prb[m]++; return orig.apply(this, a); };
      }
    """)
    page.reload()

    page.set_input_files("#wb-dxf", str(DXF))
    page.wait_for_selector("#wb-export-btn:not([disabled])")

    # 렌더 1회분만 세도록 초기화 후 [화면 맞춤] 으로 강제 렌더
    page.evaluate("window.__prb = { arc: 0, fill: 0, stroke: 0 }")
    page.click("#wb-fit-btn")
    page.wait_for_timeout(1500)
    n = page.evaluate("window.__prb")

    # 화면에 실제로 보이는(레이어 on + 픽셀 임계 통과) 엔티티만 기대값으로 삼는다.
    exp = page.evaluate("""() => {
      const z = state.view.zoom, out = { A: 0, C: 0, H: 0, S: 0, I: 0, vis: 0 };
      for (const e of state.entities) {
        const ls = state.layerState[e.l];
        if (!ls || !ls.visible) continue;
        out.vis++;
        if (e.t === "A" && e.r * z >= 0.3) out.A++;
        else if (e.t === "C" && e.r * z >= 0.5) out.C++;
        else if (e.t === "H" && e.p.length >= 3) out.H++;
        else if (e.t === "S" && e.p.length >= 3) out.S++;
        else if (e.t === "I") out.I++;
      }
      return out;
    }""")
    print(f"\n캔버스 호출 수(렌더 1회): {n}")
    print(f"화면 표시 대상 실측:       {exp}  (zoom={page.evaluate('state.view.zoom'):.5g})")

    check("ARC 렌더 (arc 호출)", n["arc"], exp["A"] + exp["C"])
    check("HATCH/SOLID/INSERT 렌더 (fill 호출)", n["fill"], exp["H"] + exp["S"] + exp["I"])
    check("HATCH 표시 대상 > 0", exp["H"] > 0, True)
    check("SOLID 표시 대상 > 0", exp["S"] > 0, True)

    # 화면맞춤 줌에서는 문 호(반지름 ~0.9m)가 서브픽셀로 컬링된다. 호가 실제로
    # 보이는 배율까지 확대해 ARC 분기를 따로 검증한다.
    page.evaluate("""() => {
      const arcs = state.entities.filter(e => e.t === "A" && state.layerState[e.l]?.visible)
                                 .sort((x, y) => x.r - y.r);
      const a = arcs[Math.floor(arcs.length / 2)];
      state.view.zoom = 3 / a.r;
      const c = document.getElementById("wb-canvas");
      state.view.panX = c.clientWidth / 2 - a.c[0] * state.view.zoom;
      state.view.panY = c.clientHeight / 2 + a.c[1] * state.view.zoom;
      window.__prb = { arc: 0, fill: 0, stroke: 0 };
      render();
    }""")
    page.wait_for_timeout(800)
    n2 = page.evaluate("window.__prb")
    exp2 = page.evaluate("""() => {
      const z = state.view.zoom; let A = 0, C = 0;
      for (const e of state.entities) {
        const ls = state.layerState[e.l];
        if (!ls || !ls.visible) continue;
        if (e.t === "A" && e.r * z >= 0.3) A++;
        else if (e.t === "C" && e.r * z >= 0.5) C++;
      }
      return { A, C };
    }""")
    print(f"\n확대 후 호출 수: {n2} / 표시 대상 {exp2} (zoom={page.evaluate('state.view.zoom'):.5g})")
    check("확대 시 ARC 표시 대상 > 0", exp2["A"] > 0, True)
    check("확대 시 ARC 렌더 (arc 호출)", n2["arc"], exp2["A"] + exp2["C"])
    check("JS 오류 0건", errors, [])

    # 검수용 캡처 — 호가 가장 밀집한 구역을 약 12m 폭으로 잡는다(문 호 판독 가능 배율).
    page.evaluate("""() => {
      const arcs = state.entities.filter(e => e.t === "A" && state.layerState[e.l]?.visible);
      const cell = 6000, buckets = new Map();
      for (const a of arcs) {
        const k = Math.round(a.c[0] / cell) + "," + Math.round(a.c[1] / cell);
        buckets.set(k, (buckets.get(k) || 0) + 1);
      }
      const best = [...buckets.entries()].sort((x, y) => y[1] - x[1])[0][0].split(",").map(Number);
      const c = document.getElementById("wb-canvas");
      state.view.zoom = c.clientWidth / 12000;
      state.view.panX = c.clientWidth / 2 - best[0] * cell * state.view.zoom;
      state.view.panY = c.clientHeight / 2 + best[1] * cell * state.view.zoom;
      render();
    }""")
    page.wait_for_timeout(800)
    out = BASE / "data" / "_prb_workbench_zoom.png"
    page.locator("#wb-canvas").screenshot(path=str(out))
    page.click("#wb-fit-btn"); page.wait_for_timeout(800)
    out2 = BASE / "data" / "_prb_workbench_fit.png"
    page.locator("#wb-canvas").screenshot(path=str(out2))
    print(f"캡처: {out} / {out2}")
    br.close()

httpd.shutdown()
print("\n" + ("PR-B 검증 통과" if not fails else f"실패 {len(fails)}건: {fails}"))
sys.exit(1 if fails else 0)
