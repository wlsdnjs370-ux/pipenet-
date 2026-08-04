# -*- coding: utf-8 -*-
"""과제 정보 패널(T4d) 런타임 검증 — 자연낙차 후보 표·경고 블록·payload 확인.

node --check 는 구문만 본다. renderContextReport 를 클릭 핸들러 밖에서 부르면
구문은 통과하고 실행할 때 ReferenceError 가 난다 — 그래서 진짜 브라우저에서 부른다.
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

PROBE = """() => {
  const errs = [];
  window.addEventListener("error", (e) => errs.push(String(e.message)));

  // /run 응답이 온 것과 동일한 상태를 만든다. 층 이름에 태그를 섞어 XSS 도 본다.
  renderContextReport({
    natural_fall_candidates: [
      {floor_label: "49F", head_drop_m: 5.0, required_m: 10.2, ok: false},
      {floor_label: "<img src=x onerror=alert(1)>26F", head_drop_m: 12.0, required_m: 10.2, ok: true},
    ],
    suggested_natural_fall_floor: "<img src=x onerror=alert(1)>26F",
    context_warnings: ["[미확정] zone_name — 존 이름 (미입력)"],
  });

  const box = document.getElementById("ov-context-report");
  const sel = document.getElementById("ov-nf-select");
  const tbody = document.getElementById("ov-nf-tbody");
  const snap = {
    visible: box.style.display !== "none",
    rows: tbody.querySelectorAll("tr").length,
    option_values: Array.from(sel.options).map((o) => o.value),
    selected: sel.value,
    warn_lines: Array.from(document.querySelectorAll("#ov-context-warnings div"))
                     .map((d) => d.textContent),
    injected_img: tbody.querySelectorAll("img").length,
    fx_options: Array.from(document.getElementById("ov-fx-profile").options).map((o) => o.value),
    new_inputs: ["ov-project-title", "ov-zone-name", "ov-machine-room-h",
                 "ov-roof-tank-level"].map((id) => !!document.getElementById(id)),
  };

  // SSE context_warning 이 도착한 경우도 같은 블록을 갱신해야 한다.
  handleEvent({type: "context_warning", lines: ["[미확정] fx_profile_key — 신축배관 규격 (기본값)"]});
  snap.warn_after_sse = Array.from(document.querySelectorAll("#ov-context-warnings div"))
                             .map((d) => d.textContent);

  // 후보 없음 + 경고 없음이면 블록은 숨어야 한다.
  renderContextReport({natural_fall_candidates: [], context_warnings: []});
  snap.hidden_when_empty = box.style.display === "none";
  snap.errors = errs;
  return snap;
}"""

from playwright.sync_api import sync_playwright  # noqa: E402

with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    console_errors: list[str] = []
    page.on("pageerror", lambda e: console_errors.append(str(e)))
    page.on("dialog", lambda d: (console_errors.append(f"XSS dialog: {d.message}"), d.dismiss()))
    page.goto(f"http://127.0.0.1:{port}/remote30-overall")
    page.wait_for_timeout(1200)
    result = page.evaluate(PROBE)
    page.screenshot(path=str(BASE / "data" / "_ctx_report.png"))
    br.close()

httpd.shutdown()

result["pageerrors"] = console_errors
for k, v in result.items():
    print(f"{k}: {v}")

ok = (
    not console_errors
    and result["visible"]
    and result["rows"] == 2
    # ok=false 인 층은 고를 수 없어야 한다 — 빈 값 + 26F 만.
    and len(result["option_values"]) == 2
    and result["selected"] == "<img src=x onerror=alert(1)>26F"
    and result["warn_lines"] == ["[미확정] zone_name — 존 이름 (미입력)"]
    and result["injected_img"] == 0
    and result["fx_options"] == ["", "평균", "한백표준"]
    and all(result["new_inputs"])
    and result["warn_after_sse"] == ["[미확정] fx_profile_key — 신축배관 규격 (기본값)"]
    and result["hidden_when_empty"]
)
print("VERDICT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
