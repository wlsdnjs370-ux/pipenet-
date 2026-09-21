# -*- coding: utf-8 -*-
"""배관망 완성(finalize) 시 stage4 애니메이션이 건너뛰어지는 경로 실측.

EventSource 생성/재연결과 animQueue.enqueue 순서를 기록해, stage3_switch 가
stage4_* 뒤에 끼어드는지(=뷰가 Stage 3 으로 되돌아가는지) 확인한다.
"""
from __future__ import annotations

import json
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

INSTRUMENT = """() => {
  window.__log = [];
  const t0 = performance.now();
  const mark = (kind, detail) => window.__log.push(
      {t: Math.round(performance.now() - t0), kind, detail});

  const OrigES = window.EventSource;
  let esN = 0;
  window.EventSource = function (url, opts) {
    const id = ++esN;
    mark("es_open", `#${id} ${url}`);
    const es = new OrigES(url, opts);
    es.addEventListener("error", () => mark("es_error", `#${id} rs=${es.readyState}`));
    const origClose = es.close.bind(es);
    es.close = () => { mark("es_close", `#${id}`); origClose(); };
    return es;
  };
  window.EventSource.prototype = OrigES.prototype;

  const origEnq = animQueue.enqueue.bind(animQueue);
  animQueue.enqueue = (evt) => { mark("enqueue", evt.type); return origEnq(evt); };

  const origSSV = window.setStageView;
  window.setStageView = (i) => { mark("setStageView", String(i)); return origSSV(i); };

  const origSP = BootOverlay.setProgress.bind(BootOverlay);
  BootOverlay.setProgress = (r, label) => {
    const vis = document.getElementById("boot-overlay").classList.contains("is-visible");
    mark("progress", `${Math.round((r || 0) * 100)}% ${label || ""} vis=${vis}`);
    return origSP(r, label);
  };
  const origHide = BootOverlay.hide.bind(BootOverlay);
  BootOverlay.hide = () => { mark("overlay_hide", ""); return origHide(); };
}"""

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
    page.wait_for_timeout(1200)
    page.evaluate(INSTRUMENT)

    page.set_input_files("#wb-dxf", str(PLANE))
    page.wait_for_function("() => !document.getElementById('wb-run').disabled")
    page.click("#wb-run")
    # 애니메이션 큐를 비우지 않고, 버튼이 열리는 즉시 누른다 — 사용자 실사용 패턴.
    page.wait_for_function("() => !document.getElementById('wb-finalize').disabled")
    page.evaluate("() => window.__log.push({t: -1, kind: 'CLICK', detail: 'finalize'})")
    page.click("#wb-finalize")
    page.wait_for_function(
        "() => document.getElementById('fx-open-btn').style.display === 'block'")
    page.wait_for_timeout(3000)

    log = page.evaluate("() => window.__log")
    view = page.evaluate("() => state.current_stage_view")
    trace_total = page.evaluate("() => TracePlayer.total")
    br.close()

httpd.shutdown()

_noise = ("head_reveal", "layer_seen", "layer_classified")
for r in log:
    if r["kind"] == "enqueue" and r["detail"] in _noise:
        continue
    print(f"  {r['t']:>7} ms  {r['kind']:<13} {r['detail']}")
print(f"\n최종 stage view = {view} (4 여야 정상) · TracePlayer.total = {trace_total}")
order = [r["detail"] for r in log if r["kind"] == "enqueue"]
if "stage4_trace" in order:
    i4 = order.index("stage4_trace")
    late3 = [d for d in order[i4:] if d == "stage3_switch"]
    print(f"stage4_trace 이후 stage3_switch 재enqueue: {len(late3)}건")
else:
    print("stage4_trace 가 큐에 들어오지 않음")

click_i = next((i for i, r in enumerate(log) if r["kind"] == "CLICK"), len(log))
after = log[click_i + 1:]
# 클릭 직후의 setStageView(3) 은 finalize 핸들러가 뷰를 확정하려고 동기 호출한 것.
# 회귀는 "Stage 4 로 넘어간 뒤에 다시 3 으로 되돌아가는" 경우뿐이다.
views = [r["detail"] for r in after if r["kind"] == "setStageView"]
back3 = views[views.index("4") + 1:].count("3") if "4" in views else -1
prog = [r["detail"] for r in after if r["kind"] == "progress"]
print(f"클릭 후 뷰 전환 순서: {views} · Stage 4 이후 3 복귀 {back3}건 (0 이어야 정상)")
print(f"클릭 후 진행바 샘플 {len(prog)}건: {prog[:3]} … {prog[-2:]}")
print(f"JS 오류 {len(errors)}건")
for e in errors[:10]:
    print("  !", e)
