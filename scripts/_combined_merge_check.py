# -*- coding: utf-8 -*-
"""통합 탭 [배관망 통합] 버튼이 왜 안 먹는지 실브라우저로 재현한다.

(a) 버튼이 disabled 에서 안 풀린다  vs  (b) 클릭은 되는데 요청/응답이 깨진다
둘 중 어느 쪽인지 가른다. fetch 는 stub 이라 서버 없이 돈다.
"""
from __future__ import annotations

import functools
import http.server
import socketserver
import sys
import threading
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

from flask import Flask, render_template  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

app = Flask(__name__, template_folder=str(BASE / "templates"),
            static_folder=str(BASE / "static"))
page_path = BASE / "data" / "_cmb_check.html"
with app.test_request_context():
    page_path.write_text(render_template("remote30_prototype.html"),
                         encoding="utf-8")

handler = functools.partial(http.server.SimpleHTTPRequestHandler,
                            directory=str(BASE))
httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
httpd.RequestHandlerClass.log_message = lambda *a, **k: None
port = httpd.server_address[1]
threading.Thread(target=httpd.serve_forever, daemon=True).start()

fails: list[str] = []


def check(label, got, want):
    ok = got == want
    print(f"  {'OK ' if ok else 'FAIL'} {label}: {got!r}" +
          ("" if ok else f"  (기대 {want!r})"))
    if not ok:
        fails.append(label)


errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1500, "height": 950})
    page.set_default_timeout(60000)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.goto(f"http://127.0.0.1:{port}/data/_cmb_check.html")
    page.wait_for_timeout(1200)

    print("① 로드 시점 JS 오류")
    check("오류 0건", errors, [])

    print("\n② 통합 탭 진입 — 미준비 상태")
    page.click("#mode-btn-combined")
    page.wait_for_timeout(200)
    check("버튼 disabled", page.eval_on_selector("#cmb-merge", "e => e.disabled"), True)
    check("현재 모드", page.evaluate("() => state.current_mode"), "combined")
    check("비활성이 회색으로 보임",
          page.eval_on_selector("#cmb-merge", "e => getComputedStyle(e).backgroundColor"),
          "rgb(203, 213, 225)")
    check("막힌 이유 툴팁",
          page.eval_on_selector("#cmb-merge", "e => e.title"),
          "평면도 추출 · 계통도 추출 을(를) 먼저 완료해야 누를 수 있습니다.")

    print("\n③ 평면도+계통도 준비 주입 후 재진입")
    page.evaluate("""() => {
      state.jobId = "job_abcdef123456";
      state.system_riser = {title: "SP-1 라이저", nodes: [
        {id: "R1", z: 0}, {id: "R2", z: 3000}, {id: "R3", z: 6000}]};
    }""")
    page.click("#mode-btn-plane")
    page.click("#mode-btn-combined")
    page.wait_for_timeout(200)
    check("평면 status", page.eval_on_selector("#cmb-plane-status", "e => e.textContent")
          .startswith("✓"), True)
    check("계통 status", page.eval_on_selector("#cmb-system-status", "e => e.textContent")
          .startswith("✓"), True)
    check("버튼 활성화", page.eval_on_selector("#cmb-merge", "e => e.disabled"), False)
    check("활성은 아웃라인(흰 배경) 유지",
          page.eval_on_selector("#cmb-merge", "e => getComputedStyle(e).backgroundColor"),
          "rgb(255, 255, 255)")
    check("활성이면 툴팁 없음", page.eval_on_selector("#cmb-merge", "e => e.title"), "")

    print("\n④ 클릭 → POST 발사 여부 (fetch stub)")
    page.evaluate("""() => {
      window.__posted = null;
      const _f = window.fetch;
      window.fetch = async function (url, opt) {
        if (String(url).includes("/combined/build")) {
          window.__posted = {url: String(url), body: JSON.parse(opt.body)};
          return {ok: true, json: async () => ({
            ok: true, n_nodes: 42, n_pipes: 41, machine_room_attached: false,
            download_url: "/x.zip", geometry: {nodes: [], pipes: []},
          })};
        }
        return _f.apply(this, arguments);
      };
      window.__alerts = [];
      window.alert = (m) => window.__alerts.push(String(m));
    }""")
    page.click("#cmb-merge")
    page.wait_for_timeout(1500)
    posted = page.evaluate("() => window.__posted")
    check("POST 발사", bool(posted), True)
    if posted:
        check("엔드포인트", posted["url"].endswith("/api/remote30/combined/build"), True)
        check("plane_job_id", posted["body"].get("plane_job_id"), "job_abcdef123456")
    check("alert 없음", page.evaluate("() => window.__alerts"), [])
    box = page.eval_on_selector("#cmb-result-box", "e => e.textContent.trim()")
    print(f"     결과박스: {box[:120]!r}")
    check("결과박스 갱신", box not in ("", "처리 중..."), True)
    check("state.combinedBuild", bool(page.evaluate("() => state.combinedBuild")), True)

    print("\n⑤ 클릭 후 JS 오류")
    check("오류 0건", errors, [])

    page.screenshot(path=str(BASE / "data" / "_cmb_check.png"))
    br.close()

httpd.shutdown()
page_path.unlink(missing_ok=True)

if errors:
    print("\nJS 오류:")
    for e in errors[:15]:
        print("  " + e)

print("\n" + ("PASS — 통합 버튼 경로 정상" if not fails
             else "FAIL — " + ", ".join(fails)))
sys.exit(1 if fails else 0)
