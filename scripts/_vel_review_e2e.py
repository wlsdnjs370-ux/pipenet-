# -*- coding: utf-8 -*-
"""내경 역산표 버튼·모달을 실서버 + 실브라우저로 검증한다.

_combined_e2e_check.py 와 같은 경로(평면도 → 계통도 → 통합 빌드)를 탄 뒤,
[내경 역산표 보기] 를 실제로 눌러 표/요약/CSV 가 나오는지 확인한다.
함수-지역 헬퍼 스코프 사고는 구문검사로 안 잡히므로 이 경로로만 검출된다.
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

PLANE = BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"
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


errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.set_default_timeout(300000)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.on("console",
            lambda m: errors.append(f"console.error: {m.text}")
            if m.type == "error" else None)
    page.goto(f"http://127.0.0.1:{port}/remote30-prototype")
    page.wait_for_timeout(1500)
    check("로드 JS 오류 0", errors, [])
    check("역산표 버튼 초기 비활성",
          page.eval_on_selector("#cmb-vel-btn", "e => e.disabled"), True)

    print("\n① 평면도 추출")
    page.set_input_files("#wb-dxf", str(PLANE))
    page.wait_for_function("() => !document.getElementById('wb-run').disabled")
    page.click("#wb-run")
    page.wait_for_function(
        "() => !document.getElementById('wb-finalize').disabled", timeout=300000)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.wait_for_timeout(500)

    print("② 배관망 완성")
    page.click("#wb-finalize")
    page.wait_for_function(
        "() => document.getElementById('fx-open-btn').style.display === 'block'",
        timeout=300000)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.wait_for_timeout(500)

    print("③ 계통도 riser")
    riser = page.evaluate("""async () => {
      const fd = new FormData();
      fd.append("use_legacy_template", "true");
      fd.append("pump_x", "0"); fd.append("pump_y", "0");
      fd.append("av_x", "0");   fd.append("av_y", "-3000");
      const r = await fetch("/api/remote30/system/extract", {method: "POST", body: fd});
      const d = await r.json();
      if (!d.ok) return {ok: false, msg: d.message};
      _revealSystemRiser(d);
      return {ok: true};
    }""")
    check("riser 수신", riser.get("ok"), True)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.wait_for_timeout(500)

    print("④ 통합 빌드")
    page.click("#mode-btn-combined")
    page.wait_for_timeout(400)
    page.evaluate("""() => { window.__alerts = [];
                             window.alert = (m) => window.__alerts.push(String(m)); }""")
    page.click("#cmb-merge")
    page.wait_for_function(
        "() => document.getElementById('cmb-merge').textContent.indexOf('통합 중') < 0",
        timeout=300000)
    page.wait_for_timeout(1000)
    check("alert 없음", page.evaluate("() => window.__alerts"), [])
    check("combinedBuild 생성", bool(page.evaluate("() => state.combinedBuild")))

    print("\n⑤ 내경 역산표")
    rep = page.evaluate("() => (state.combinedBuild || {}).velocity_report || null")
    check("velocity_report 수신", bool(rep))
    if rep:
        print(f"     수렴 {rep['iterations']}회차 · 변경 {rep['changed']} · "
              f"최대유속 {rep['max_velocity_before']}→{rep['max_velocity_after']} · "
              f"초과 {rep['violations_before']}→{rep['violations_after']} · "
              f"changes {len(rep.get('changes') or [])}행")
    check("버튼 활성화", page.eval_on_selector("#cmb-vel-btn", "e => e.disabled"), False)

    page.click("#cmb-vel-btn")
    page.wait_for_timeout(400)
    check("모달 열림",
          page.eval_on_selector("#vel-modal-overlay",
                                "e => e.classList.contains('is-open')"), True)
    rows = page.eval_on_selector_all("#vel-table tbody tr", "els => els.length")
    cols = page.eval_on_selector_all("#vel-table thead th", "els => els.length")
    n_changes = len(rep.get("changes") or []) if rep else 0
    print(f"     표 {rows}행 × {cols}열")
    check("헤더 10열", cols, 10 if n_changes else 0)
    check("행 수 = changes 길이", rows, n_changes)
    summary = page.eval_on_selector("#vel-summary", "e => e.textContent.trim()")
    print(f"     요약: {summary!r}")
    check("요약 채워짐", bool(summary))
    if rows:
        first = page.eval_on_selector_all(
            "#vel-table tbody tr:first-child td", "els => els.map(e => e.textContent)")
        print(f"     첫 행: {first}")
        check("첫 행 사유 있음", bool(first[-1].strip()))

    warn_on = page.eval_on_selector("#vel-warning", "e => e.style.display !== 'none'")
    ex = (rep or {}).get("export") or {}
    want_warn = bool(rep and (rep["violations_after"] > 0
                              or rep.get("head_stub_violations", 0) > 0
                              or (ex.get("surplus_bar") or 0) > 0.05))
    check("경고 표시 = 한계도달/과토출/규격고정 중 하나", warn_on, want_warn)
    if rep and rep.get("export"):
        print(f"     산출물 경계: 수원 {ex['source_pressure_bar']} bar "
              f"(필요 {ex['required_source_bar']}, 잉여 {ex['surplus_bar']}) · "
              f"헤드 {ex['head_flow_min']}~{ex['head_flow_max']} L/min · "
              f"최대유속 {ex['max_velocity']} · 초과 {ex['violations']}구간")
        print(f"     FX 스텁 {rep['head_stub_bore_mm']}A · 초과 "
              f"{rep['head_stub_violations']}개 · 최대 {rep['head_stub_max_velocity']} m/s")

    # 대명동 실망은 도면 관경만으로 상한을 만족해 changes 가 비는 경우가 있다.
    # 표/CSV 의 채워진 경로도 반드시 태워야 하므로, 솔버가 실제로 만든 변경 이력을
    # 하나 더 주입해 같은 렌더러로 그린다(형태 위조가 아니라 진짜 솔버 산출물).
    if not (rep.get("changes") if rep else None):
        print("\n⑤-b 변경이 0건 — 솔버 실산출 변경표를 주입해 채워진 경로 검증")
        import json as _json
        sys.path.insert(0, str(BASE / "core"))
        from hydraulic_solver import converge_bores_by_velocity  # noqa: E402
        sys.path.insert(0, str(BASE / "tests"))
        from test_hydraulic_solver import make_net  # noqa: E402
        _n, _p, _z = make_net()
        forced = converge_bores_by_velocity(_n, _p, _z, keep_existing=False)
        print(f"     주입 리포트: 변경 {forced['changed']} · {len(forced['changes'])}행")
        page.evaluate("(r) => { state.combinedBuild.velocity_report = r; "
                      "updateCombinedStatus(); VelReview.open(); }",
                      _json.loads(_json.dumps(forced)))
        page.wait_for_timeout(300)
        rows = page.eval_on_selector_all("#vel-table tbody tr", "els => els.length")
        cols = page.eval_on_selector_all("#vel-table thead th", "els => els.length")
        print(f"     표 {rows}행 × {cols}열")
        check("헤더 10열(주입)", cols, 10)
        check("행 수 = changes 길이(주입)", rows, len(forced["changes"]))
        first = page.eval_on_selector_all(
            "#vel-table tbody tr:first-child td", "els => els.map(e => e.textContent)")
        print(f"     첫 행: {first}")
        check("첫 행 사유 있음(주입)", bool(first[-1].strip()))
        check("승급 강조 클래스",
              page.eval_on_selector_all("#vel-table td.vel-up", "e => e.length") > 0)

    # CSV — 클릭 대신 내부 생성 경로를 그대로 태워 본문을 확인한다.
    page.evaluate("""() => { window.__csv = null;
      const origCreate = URL.createObjectURL;
      URL.createObjectURL = (b) => { window.__csvBlob = b; return origCreate.call(URL, b); };
    }""")
    page.click("#vel-csv-btn")
    page.wait_for_timeout(300)
    csv_head = page.evaluate("""async () => {
      if (!window.__csvBlob) return null;
      const t = await window.__csvBlob.text();
      return t.split("\\r\\n").slice(0, 3);
    }""")
    print(f"     CSV: {csv_head}")
    check("CSV 생성", bool(csv_head))
    if csv_head:
        check("CSV BOM+한글 헤더", csv_head[0].lstrip("﻿").startswith("배관,구분"))

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    check("Esc 닫힘",
          page.eval_on_selector("#vel-modal-overlay",
                                "e => e.classList.contains('is-open')"), False)

    print("\n⑥ 전체 JS 오류")
    check("오류 0건", errors, [])
    page.screenshot(path=str(BASE / "data" / "_vel_e2e.png"))
    br.close()

httpd.shutdown()

if errors:
    print("\nJS 오류:")
    for e in errors[:20]:
        print("  " + e)

print("\n" + ("PASS — 내경 역산표 e2e 정상" if not fails
             else "FAIL — " + ", ".join(fails)))
sys.exit(1 if fails else 0)
