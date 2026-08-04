# -*- coding: utf-8 -*-
"""구역 유형 드롭다운(T1) 런타임 검증 — 실제 브라우저에서 스코프·DOM·payload 확인.

node --check 는 구문만 본다. 함수-지역 헬퍼를 다른 스코프에서 부르면 구문은
통과하고 클릭할 때 ReferenceError 가 난다 — 그래서 진짜 브라우저에서 부른다.
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

from playwright.sync_api import sync_playwright  # noqa: E402

PROBE = """() => {
  const errs = [];
  window.addEventListener("error", (e) => errs.push(String(e.message)));

  // 구역 2개를 그린 것과 같은 상태를 만들고 UI 를 갱신한다.
  state.zones = [[0, 0, 1000, 1000], [5000, 5000, 6000, 6000]];
  state.zone_kinds = [];
  updateEditCounts();

  const box = document.getElementById("wb-zone-kinds");
  const sels = Array.from(document.querySelectorAll("#wb-zone-kind-rows select"));
  const before = materialZonesPayload();

  // 1번 구역을 단위세대로 고른다 (사람이 드롭다운을 바꾼 것과 동일).
  sels[0].value = "unit_dwelling";
  sels[0].dispatchEvent(new Event("change"));
  const after = materialZonesPayload();
  const snapshot = {
    visible_with_zones: box.style.display !== "none",
    select_count: sels.length,
    options: Array.from(sels[0].options).map((o) => o.value),
    payload_before: before,
    payload_after: after,
    kinds_after: state.zone_kinds.slice(),
  };

  // 영역을 전부 해제하면 드롭다운도 같이 사라져야 한다.
  document.getElementById("wb-zones-clear").click();
  snapshot.errors = errs;
  snapshot.cleared_zone_kinds = state.zone_kinds.length;
  snapshot.hidden_after_clear = box.style.display === "none";
  return snapshot;
}"""

with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    console_errors: list[str] = []
    page.on("pageerror", lambda e: console_errors.append(str(e)))
    page.goto(f"http://127.0.0.1:{port}/remote30-prototype")
    page.wait_for_timeout(1200)
    result = page.evaluate(PROBE)
    page.screenshot(path=str(BASE / "data" / "_zone_kind.png"))
    br.close()

httpd.shutdown()

result["pageerrors"] = console_errors
for k, v in result.items():
    print(f"{k}: {v}")

ok = (
    not console_errors
    and result["visible_with_zones"]
    and result["select_count"] == 2
    and result["options"] == ["", "unit_dwelling", "parking", "corridor"]
    and result["payload_before"] == []
    and result["payload_after"] == [{"rect": [0, 0, 1000, 1000], "kind": "unit_dwelling"}]
    and result["kinds_after"] == ["unit_dwelling", ""]
    and result["cleared_zone_kinds"] == 0
    and result["hidden_after_clear"]
)
print("VERDICT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
