# -*- coding: utf-8 -*-
"""실서버 + 실브라우저로 통합 [배관망 통합] 버튼을 처음부터 끝까지 재현한다.

로그인 게이트만 벗긴 진짜 Flask 앱을 임시 포트에 띄우고, 대명동 도면으로
평면도 추출 → 배관망 완성 → 계통도 riser → 통합 버튼까지 실제로 눌러 본다.
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

# 로컬 하니스 — 로그인 게이트 제거 후 세션을 authed 로 고정.
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

    print("\n① 평면도 추출 (Stage 0~2)")
    page.set_input_files("#wb-dxf", str(PLANE))
    page.wait_for_function("() => !document.getElementById('wb-run').disabled")
    page.click("#wb-run")
    page.wait_for_function(
        "() => !document.getElementById('wb-finalize').disabled",
        timeout=300000)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.wait_for_timeout(500)
    check("jobId 확보", bool(page.evaluate("() => state.jobId")))
    print(f"     헤드 {page.evaluate('() => state.detected_heads.length')}개")

    print("\n② 배관망 완성 (Stage 3~5)")
    page.click("#wb-finalize")
    page.wait_for_function(
        "() => document.getElementById('fx-open-btn').style.display === 'block'",
        timeout=300000)
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.wait_for_timeout(500)
    check("Stage 5 완료(FX 버튼 노출)", True)

    print("\n③ 계통도 riser — 실제 엔드포인트 응답을 실제 핸들러에 통과")
    riser_ok = page.evaluate("""async () => {
      const fd = new FormData();
      fd.append("use_legacy_template", "true");
      fd.append("pump_x", "0"); fd.append("pump_y", "0");
      fd.append("av_x", "0");   fd.append("av_y", "-3000");
      const r = await fetch("/api/remote30/system/extract", {method: "POST", body: fd});
      const d = await r.json();
      if (!d.ok) return {ok: false, msg: d.message};
      _revealSystemRiser(d);
      return {ok: true, keys: Object.keys(d.riser).slice(0, 12)};
    }""")
    check("riser 수신", riser_ok.get("ok"), True)
    if not riser_ok.get("ok"):
        print(f"     message: {riser_ok.get('msg')}")
    page.evaluate("() => animQueue.flush && animQueue.flush()")
    page.wait_for_timeout(500)
    check("state.system_riser", bool(page.evaluate("() => state.system_riser")))

    print("\n④ 통합 탭 진입")
    page.click("#mode-btn-combined")
    page.wait_for_timeout(400)
    print(f"     평면 status: "
          f"{page.eval_on_selector('#cmb-plane-status', 'e => e.textContent')!r}")
    print(f"     계통 status: "
          f"{page.eval_on_selector('#cmb-system-status', 'e => e.textContent')!r}")
    check("버튼 활성화", page.eval_on_selector("#cmb-merge", "e => e.disabled"), False)

    print("\n⑤ [배관망 통합] 클릭 — 진짜 서버 빌드")
    page.evaluate("""() => { window.__alerts = [];
                             window.alert = (m) => window.__alerts.push(String(m)); }""")
    if not page.eval_on_selector("#cmb-merge", "e => e.disabled"):
        page.click("#cmb-merge")
        page.wait_for_function(
            "() => document.getElementById('cmb-merge').textContent.indexOf('통합 중') < 0",
            timeout=300000)
        page.wait_for_timeout(1000)
    check("alert 없음", page.evaluate("() => window.__alerts"), [])
    box = page.eval_on_selector("#cmb-result-box", "e => e.textContent.trim()")
    print(f"     결과박스: {box[:250]!r}")
    check("combinedBuild 생성", bool(page.evaluate("() => state.combinedBuild")))
    check("combined_geometry", bool(page.evaluate("() => state.combined_geometry")))

    print("\n⑥ 전체 JS 오류")
    check("오류 0건", errors, [])
    page.screenshot(path=str(BASE / "data" / "_cmb_e2e.png"), full_page=False)
    br.close()

httpd.shutdown()

if errors:
    print("\nJS 오류:")
    for e in errors[:20]:
        print("  " + e)

print("\n" + ("PASS — 통합 버튼 e2e 정상" if not fails
             else "FAIL — " + ", ".join(fails)))
sys.exit(1 if fails else 0)
