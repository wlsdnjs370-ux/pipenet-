# -*- coding: utf-8 -*-
"""라이저 형상 입력칸 런타임 검증 — 작업지시서 4-1/4-3.

node --check 는 구문만 본다. 실제로 폼을 채우고 "배관망 생성" 을 눌렀을 때
새 7칸이 FormData 에 실려 나가는지, 빈칸은 안 실리는지를 브라우저에서 본다.
/run 요청은 가로채서 본문만 읽고 서버까지 보내지 않는다.
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

NEW_IDS = ["ov-roof-run-1", "ov-roof-run-2", "ov-roof-drop",
           "ov-tee-to-av", "ov-tee-above-slab", "ov-top-floor-extra", "ov-arrester"]

# 셋만 채운다 — 나머지는 안 실려야 서버가 [미확정] 로 잡는다.
FILLED = {"ov-roof-run-1": "8.4", "ov-roof-drop": "3.75", "ov-tee-to-av": "0.7"}

SETUP = """(filled) => {
  const missing = %s.filter((id) => !document.getElementById(id));
  for (const [id, v] of Object.entries(filled)) document.getElementById(id).value = v;
  document.getElementById("ov-arrester").checked = true;
  document.getElementById("ov-target-floor").value = "16층";
  // 업로드 없이 핸들러만 태운다 — 버튼 활성화는 업로드 핸들러가 하던 일이다.
  state.uploadedFile = new File(["0\\n"], "probe.dxf", {type: "application/dxf"});
  document.getElementById("wb-run").disabled = false;
  return missing;
}""" % NEW_IDS

from playwright.sync_api import sync_playwright  # noqa: E402

posted: list[str] = []
dialogs: list[str] = []
pageerrors: list[str] = []

with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.on("pageerror", lambda e: pageerrors.append(str(e)))
    page.on("dialog", lambda d: (dialogs.append(d.message), d.dismiss()))

    def _intercept(route):
        posted.append(route.request.post_data or "")
        route.fulfill(status=200, content_type="application/json",
                      body='{"ok": false, "message": "probe"}')

    page.route("**/api/remote30/overall/run", _intercept)
    page.goto(f"http://127.0.0.1:{port}/remote30-overall")
    page.wait_for_timeout(1200)
    missing = page.evaluate(SETUP, FILLED)
    page.click("#wb-run")
    page.wait_for_timeout(1500)
    page.screenshot(path=str(BASE / "data" / "_riser_geometry_form.png"))
    br.close()

httpd.shutdown()

body = posted[0] if posted else ""
sent = {line.split('name="', 1)[1].split('"', 1)[0]
        for line in body.splitlines() if 'name="' in line}
geometry_keys = {"roof_run_to_riser_m", "roof_run_after_drop_m", "roof_to_top_floor_drop_m",
                 "tee_to_alarm_valve_m", "tee_branch_above_slab_m",
                 "top_floor_extra_height_m", "water_hammer_arrester"}

print("missing_controls:", missing)
print("requests_intercepted:", len(posted))
print("geometry_fields_sent:", sorted(sent & geometry_keys))
print("values:", [v for v in ("8.4", "3.75", "0.7") if v in body])
print("dialogs:", dialogs)
print("pageerrors:", pageerrors)

ok = (
    not missing
    and not pageerrors
    and len(posted) == 1
    and sent & geometry_keys == {"roof_run_to_riser_m", "roof_to_top_floor_drop_m",
                                 "tee_to_alarm_valve_m", "water_hammer_arrester"}
    and all(v in body for v in ("8.4", "3.75", "0.7"))
    and dialogs == ["실행 실패: probe"]   # 가로챈 응답이 만든 것 — 예상된 대화상자
)
print("VERDICT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
