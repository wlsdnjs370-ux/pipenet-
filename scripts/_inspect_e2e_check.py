# -*- coding: utf-8 -*-
"""통합 요소 인스펙터를 실서버 + 실브라우저로 왕복 검증한다.

대명동 도면으로 평면도→완성→계통도→통합까지 실제로 눌러 통합망을 만든 뒤,
캔버스의 배관/헤드를 좌클릭(조회)·우클릭(편집)하고 값을 고쳐 재출력한다.
마지막에 서버가 낸 .sdf 를 파싱해 편집치가 그대로 실렸는지 확인한다.
"""
from __future__ import annotations

import importlib
import sys
import threading
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

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

PLANE = BASE / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"
if not PLANE.is_file():
    print(f"평면도 도면 없음: {PLANE}")
    sys.exit(2)

from playwright.sync_api import sync_playwright  # noqa: E402

fails: list[str] = []


def check(label, got, want=True):
    ok = got == want
    print(f"  {'OK ' if ok else 'FAIL'} {label}: {got!r}" +
          ("" if ok else f"  (기대 {want!r})"))
    if not ok:
        fails.append(label)


PICK_JS = """() => {
  const g0 = state.combined_geometry;
  const { g, proj } = _combinedDisplayAndProjector(g0, state.combined_view);
  const nm = new Map(); for (const n of g.nodes) nm.set(n.label, n);
  const r = canvas.getBoundingClientRect();
  const S2 = (n) => { const p = proj(n.x, n.y, n.z);
    return [p[0] * state.view.zoom + state.view.panX,
            -p[1] * state.view.zoom + state.view.panY]; };
  const inView = (x, y) => x > 24 && y > 24 && x < r.width - 24 && y < r.height - 24;
  const pts = g.nodes.map(S2);
  const heads = new Set((g0.nozzles || []).map(z => z.in));
  let head = null;
  for (const n of g.nodes) {
    if (!heads.has(n.label)) continue;
    const [x, y] = S2(n);
    if (inView(x, y)) { head = { label: n.label, cx: x + r.left, cy: y + r.top }; break; }
  }
  let pipe = null;
  for (const p of g0.pipes) {
    const a = nm.get(p.in), b = nm.get(p.out); if (!a || !b) continue;
    const [x1, y1] = S2(a), [x2, y2] = S2(b);
    if (Math.hypot(x2 - x1, y2 - y1) < 70) continue;
    const mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
    if (!inView(mx, my)) continue;
    if (pts.some(([px, py]) => Math.hypot(px - mx, py - my) < 18)) continue;
    pipe = { label: p.label, cx: mx + r.left, cy: my + r.top }; break;
  }
  return { head, pipe, nozzles: (g0.nozzles || []).length,
           fittings: (g0.fittings || []).length,
           equipment: (g0.equipment || []).length };
}"""

errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950},
                       accept_downloads=True)
    page.set_default_timeout(300000)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.on("console",
            lambda m: errors.append(f"console.error: {m.text}")
            if m.type == "error" else None)
    page.on("download", lambda d: None)
    page.goto(f"http://127.0.0.1:{port}/remote30-prototype")
    page.wait_for_timeout(1500)
    check("로드 JS 오류 0", errors, [])

    print("\n① 평면도 추출 → 배관망 완성 → 계통도 → 통합")
    page.set_input_files("#wb-dxf", str(PLANE))
    page.wait_for_function("() => !document.getElementById('wb-run').disabled")
    page.click("#wb-run")
    page.wait_for_function(
        "() => !document.getElementById('wb-finalize').disabled", timeout=300000)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.click("#wb-finalize")
    page.wait_for_function(
        "() => document.getElementById('fx-open-btn').style.display === 'block'",
        timeout=300000)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    riser = page.evaluate("""async () => {
      const fd = new FormData();
      fd.append("use_legacy_template", "true");
      fd.append("pump_x", "0"); fd.append("pump_y", "0");
      fd.append("av_x", "0");   fd.append("av_y", "-3000");
      const r = await fetch("/api/remote30/system/extract", {method: "POST", body: fd});
      const d = await r.json();
      if (!d.ok) return {ok: false, msg: d.message};
      _revealSystemRiser(d); return {ok: true};
    }""")
    check("riser 수신", riser.get("ok"), True)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.click("#mode-btn-combined")
    page.wait_for_timeout(400)
    page.evaluate("""() => { window.__alerts = [];
                             window.alert = (m) => window.__alerts.push(String(m)); }""")
    page.click("#cmb-merge")
    page.wait_for_function(
        "() => document.getElementById('cmb-merge').textContent.indexOf('통합 중') < 0",
        timeout=300000)
    page.wait_for_timeout(1200)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    check("통합 geometry", bool(page.evaluate("() => state.combined_geometry")))
    check("alert 없음", page.evaluate("() => window.__alerts"), [])

    pick = page.evaluate(PICK_JS)
    print(f"     부속테이블 — 노즐 {pick['nozzles']} · 부속 {pick['fittings']} "
          f"· 등가길이 {pick['equipment']}")
    check("geometry.nozzles 존재", pick["nozzles"] > 0)
    check("클릭 대상 배관 선정", bool(pick["pipe"]))
    check("클릭 대상 헤드 선정", bool(pick["head"]))
    if not (pick["pipe"] and pick["head"]):
        print("대상 선정 실패 — 중단")
        br.close(); httpd.shutdown(); sys.exit(1)

    print("\n② 좌클릭 = 조회")
    page.mouse.click(pick["pipe"]["cx"], pick["pipe"]["cy"])
    page.wait_for_timeout(300)
    check("인스펙터 열림",
          page.eval_on_selector("#cmb-inspect", "e => e.classList.contains('is-open')"))
    check("배지 조회", page.eval_on_selector("#ci-badge", "e => e.textContent"), "조회")
    check("편집 세션 미개시", page.evaluate("() => CombinedEditor.isActive()"), False)
    body = page.eval_on_selector("#ci-body", "e => e.textContent")
    for kw in ("내경 (mm)", "재질", "C 계수", "부속류", "등가길이"):
        check(f"배관 조회 항목 «{kw}»", kw in body)

    page.keyboard.press("Escape")
    page.mouse.click(pick["head"]["cx"], pick["head"]["cy"])
    page.wait_for_timeout(300)
    hbody = page.eval_on_selector("#ci-body", "e => e.textContent")
    check("헤드 인스펙터", "설치 방향" in hbody)
    check("상/하향식 표기", ("상향식" in hbody) or ("하향식" in hbody))

    print("\n③ 우클릭 = 편집 (배관)")
    page.keyboard.press("Escape")
    page.mouse.click(pick["pipe"]["cx"], pick["pipe"]["cy"], button="right")
    page.wait_for_timeout(400)
    check("배지 편집", page.eval_on_selector("#ci-badge", "e => e.textContent"), "편집")
    check("편집 세션 자동 개시", page.evaluate("() => CombinedEditor.isActive()"), True)
    plabel = page.eval_on_selector("#ci-title", "e => e.textContent").replace("배관 ", "")
    print(f"     대상 배관: {plabel!r}")

    page.select_option("#ci-body select[data-k='material']", "CPVC2")
    page.wait_for_timeout(250)
    mat = page.evaluate("""(lbl) => {
      const p = state.combined_geometry.pipes.find(x => x.label === lbl);
      return p ? {type: p.type, c: p.c} : null;
    }""", plabel)
    check("재질 → CPVC2", mat and mat["type"], "CPVC2")
    check("C 계수 동기화", mat and mat["c"], 150)

    page.click("#ci-body button[data-k='f_add']")
    page.wait_for_timeout(250)
    page.select_option("#ci-body [data-fit] select[data-k='f_type']", "elbow")
    page.wait_for_timeout(200)
    page.fill("#ci-body [data-fit] input[data-k='f_count']", "2")
    page.dispatch_event("#ci-body [data-fit] input[data-k='f_count']", "change")
    page.wait_for_timeout(250)

    page.click("#ci-body button[data-k='e_add']")
    page.wait_for_timeout(250)
    page.fill("#ci-body [data-eq] input[data-k='e_eq_len']", "7.5")
    page.dispatch_event("#ci-body [data-eq] input[data-k='e_eq_len']", "change")
    page.wait_for_timeout(250)
    page.fill("#ci-body [data-eq] input[data-k='e_note']", "인스펙터 실측")
    page.dispatch_event("#ci-body [data-eq] input[data-k='e_note']", "change")
    page.wait_for_timeout(250)

    sub = page.evaluate("""(lbl) => {
      const g = state.combined_geometry;
      const f = (g.fittings || []).filter(x => x.pipe === lbl);
      const e = (g.equipment || []).filter(x => x.pipe === lbl);
      return {f, e};
    }""", plabel)
    check("부속 추가·종류", any(x["type"] == "elbow" for x in sub["f"]))
    check("부속 개수 2", any(int(x["count"]) == 2 for x in sub["f"]))
    check("등가길이 7.5", any(abs(float(x["eq_len"]) - 7.5) < 1e-6 for x in sub["e"]))
    check("근거=직접입력", any(x["override_flag"] and x["override_note"] == "인스펙터 실측"
                          for x in sub["e"]))

    print("\n④ 우클릭 = 편집 (헤드 방향)")
    page.keyboard.press("Escape")
    page.mouse.click(pick["head"]["cx"], pick["head"]["cy"], button="right")
    page.wait_for_timeout(400)
    nlabel = page.eval_on_selector("#ci-title", "e => e.textContent").replace("노드 ", "")
    before = page.evaluate("""(lbl) => {
      const nz = (state.combined_geometry.nozzles || []).find(z => z.in === lbl);
      return nz ? nz.orientation : null;
    }""", nlabel)
    flip = "upright" if before != "upright" else "pendent"
    page.select_option("#ci-body select[data-k='nz_orientation']", flip)
    page.wait_for_timeout(250)
    page.fill("#ci-body input[data-k='nz_flow']", "88")
    page.dispatch_event("#ci-body input[data-k='nz_flow']", "change")
    page.wait_for_timeout(250)
    after = page.evaluate("""(lbl) => {
      const nz = (state.combined_geometry.nozzles || []).find(z => z.in === lbl);
      return nz ? {o: nz.orientation, f: nz.flow_lmin} : null;
    }""", nlabel)
    check(f"헤드 방향 {before} → {flip}", after and after["o"], flip)
    check("방수량 88 L/min", after and after["f"], 88)

    print("\n⑤ 편집본 재출력 → 서버 SDF")
    page.keyboard.press("Escape")
    check("재출력 버튼 활성",
          page.eval_on_selector("#ce-rebuild", "e => e.disabled"), False)
    page.click("#ce-rebuild")
    page.wait_for_function(
        "() => document.getElementById('ce-rebuild').textContent.indexOf('재출력 중') < 0",
        timeout=300000)
    page.wait_for_timeout(800)
    check("재출력 alert 없음", page.evaluate("() => window.__alerts"), [])
    sdf_url = page.evaluate("() => state.combinedBuild.download_url_sdf")
    print(f"     SDF: {sdf_url}")
    check("SDF URL", bool(sdf_url))

    print("\n⑥ 기존 도구 회귀 — 삭제 도구 (인스펙터가 클릭을 가로채지 않는지)")
    page.evaluate("() => { window.confirm = () => true; }")
    n0 = page.evaluate("() => state.combined_geometry.nodes.length")
    page.click("#cmb-edit .ce-tools button[data-tool='delete']")
    page.wait_for_timeout(200)
    pick2 = page.evaluate(PICK_JS)
    if pick2["head"]:
        page.mouse.click(pick2["head"]["cx"], pick2["head"]["cy"])
        page.wait_for_timeout(300)
    n1 = page.evaluate("() => state.combined_geometry.nodes.length")
    check(f"삭제 도구 동작 ({n0} → {n1})", n1 == n0 - 1)
    check("삭제 시 인스펙터 미개방",
          page.eval_on_selector("#cmb-inspect", "e => e.classList.contains('is-open')"),
          False)
    page.click("#ce-revert")
    page.wait_for_timeout(400)
    check("원본 복원",
          page.evaluate("() => state.combined_geometry.nodes.length"), n0)

    print("\n⑦ 전체 JS 오류")
    check("오류 0건", errors, [])
    page.screenshot(path=str(BASE / "data" / "_ins_e2e.png"))
    br.close()

if sdf_url:
    with urllib.request.urlopen(f"http://127.0.0.1:{port}{sdf_url}") as r:
        sdf_txt = r.read().decode("utf-8", "replace")
    root = ET.fromstring(sdf_txt)
    pipe_el = next((p for p in root.iter("Pipe")
                    if p.attrib.get("label") == plabel), None)
    print(f"\n⑧ SDF 검증 — <Pipe label={plabel!r}>")
    check("SDF 에 해당 배관 존재", pipe_el is not None)
    if pipe_el is not None:
        check("roughness-or-c=150", pipe_el.attrib.get("roughness-or-c"), "150")
        fit = [f.attrib for f in pipe_el.iter("Fitting")]
        eq = [e.attrib for e in pipe_el.iter("Equipment")]
        print(f"     Fitting {fit}")
        print(f"     Equipment {eq}")
        check("Fitting type=elbow", any(f.get("type") == "elbow" for f in fit))
        check("Fitting count=2", any(str(f.get("count")) == "2" for f in fit))
        check("equivalent-length=7.5",
              any(abs(float(e.get("equivalent-length", 0)) - 7.5) < 1e-6 for e in eq))

httpd.shutdown()

if errors:
    print("\nJS 오류:")
    for e in errors[:20]:
        print("  " + e)

print("\n" + ("PASS — 인스펙터 조회/편집/출력 왕복 정상" if not fails
             else "FAIL — " + ", ".join(fails)))
sys.exit(1 if fails else 0)
