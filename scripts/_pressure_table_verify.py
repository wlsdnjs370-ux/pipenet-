# -*- coding: utf-8 -*-
"""압력표 편집기 런타임 검증 — 실제 브라우저에서 표↔JSON 왕복과 제출값을 확인.

node --check 는 구문만 본다. 표 편집기는 함수 스코프가 여러 겹이라 진짜
브라우저에서 클릭·입력해 보지 않으면 ReferenceError 를 못 잡는다.
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
  const snap = {};
  const tbody = document.getElementById("ov-pt-tbody");
  const jsonEl = document.getElementById("ov-pressure-json");
  const setInput = (tr, key, val) => {
    const inp = tr.querySelector(`input[data-pt-field="${key}"]`);
    inp.value = val;
    inp.dispatchEvent(new Event("input"));
  };

  snap.initial_rows = tbody.rows.length;          // 빈 행 1개로 시작
  snap.initial_payload = ptReadRows();            // 층 이름이 없으니 비어야 한다

  // 사람이 직접 두 층을 채우는 것과 동일
  setInput(tbody.rows[0], "floor_label", "옥상층");
  setInput(tbody.rows[0], "height_m", "4");
  setInput(tbody.rows[0], "head_drop_m", "0");
  document.getElementById("ov-pt-add").click();
  setInput(tbody.rows[1], "floor_label", "18층");
  setInput(tbody.rows[1], "height_m", "2.9");
  setInput(tbody.rows[1], "head_drop_m", "13.8");
  setInput(tbody.rows[1], "after_prv_m", "2");
  setInput(tbody.rows[1], "note", "1차 감압");
  snap.typed_payload = ptReadRows();
  snap.mirrored = JSON.parse(jsonEl.value);

  // 행 삭제 (✕ 버튼)
  tbody.rows[1].querySelector("button").click();
  snap.after_delete = ptReadRows().length;

  // 원시 JSON 붙여넣기 → 표로 흡수
  jsonEl.value = JSON.stringify([
    {floor_label: "옥상층", height_m: 4, head_drop_m: 0},
    {floor_label: "16층", height_m: 2.9, head_drop_m: 19.6, after_prv_m: 7.8},
  ]);
  jsonEl.dispatchEvent(new Event("change"));
  snap.pasted_rows = tbody.rows.length;
  snap.pasted_payload = ptReadRows();

  // 깨진 JSON 은 표를 덮어쓰지 않아야 한다
  jsonEl.value = "{이건 JSON 이 아니다";
  jsonEl.dispatchEvent(new Event("change"));
  snap.after_bad_paste = ptReadRows();
  snap.bad_status = document.getElementById("ov-pt-status").textContent;

  // 층 이름에 스크립트 태그를 넣어도 DOM 이 만들어지지 않아야 한다
  setInput(tbody.rows[0], "floor_label", "<img src=x onerror=alert(1)>");
  snap.injected_img_count = document.querySelectorAll("#ov-pt-tbody img").length;

  // 표 비우기
  document.getElementById("ov-pt-clear").click();
  snap.cleared_rows = tbody.rows.length;
  snap.cleared_payload = ptReadRows();
  snap.cleared_json = jsonEl.value;
  return snap;
}"""

with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    console_errors: list[str] = []
    page.on("pageerror", lambda e: console_errors.append(str(e)))
    page.goto(f"http://127.0.0.1:{port}/remote30-overall")
    page.wait_for_timeout(1200)
    result = page.evaluate(PROBE)
    page.screenshot(path=str(BASE / "data" / "_pressure_table.png"))
    br.close()

httpd.shutdown()

for k, v in result.items():
    print(f"{k}: {v}")
print(f"pageerrors: {console_errors}")

expect_typed = [
    {"floor_label": "옥상층", "height_m": 4, "head_drop_m": 0},
    {"floor_label": "18층", "height_m": 2.9, "head_drop_m": 13.8,
     "after_prv_m": 2, "note": "1차 감압"},
]
expect_pasted = [
    {"floor_label": "옥상층", "height_m": 4, "head_drop_m": 0},
    {"floor_label": "16층", "height_m": 2.9, "head_drop_m": 19.6, "after_prv_m": 7.8},
]
ok = (
    not console_errors
    and result["initial_rows"] == 1
    and result["initial_payload"] == []
    and result["typed_payload"] == expect_typed
    and result["mirrored"] == expect_typed
    and result["after_delete"] == 1
    and result["pasted_rows"] == 2
    and result["pasted_payload"] == expect_pasted
    and result["after_bad_paste"] == expect_pasted
    and result["bad_status"].startswith("❌")
    and result["injected_img_count"] == 0
    and result["cleared_rows"] == 1
    and result["cleared_payload"] == []
    and result["cleared_json"] == ""
)
print("VERDICT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
