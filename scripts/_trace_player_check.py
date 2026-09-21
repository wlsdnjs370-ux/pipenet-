# -*- coding: utf-8 -*-
"""Stage 4 추적 재생(TracePlayer)을 실제 브라우저에서 눌러보며 확인.

가짜 stage-4 엔티티(AV → 체인 → 분기 → 트리밖 edge)를 주입해
재생·일시정지·스텝·스크럽·선단따라가기·뷰전환을 전부 구동한다.
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
with app.test_request_context():
    html = render_template("remote30_prototype.html")

out = BASE / "data" / "_trace_check.html"
out.write_text(html, encoding="utf-8")

handler = functools.partial(http.server.SimpleHTTPRequestHandler,
                            directory=str(BASE))
httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
httpd.RequestHandlerClass.log_message = lambda *a, **k: None
port = httpd.server_address[1]
threading.Thread(target=httpd.serve_forever, daemon=True).start()

E = lambda x1, y1, x2, y2, m, d, off=False: {  # noqa: E731
    "t": "L", "l": "_subgraph", "p": [x1, y1, x2, y2], "m": m, "d": d,
    **({"x": 1} if off else {})}
ENTS = [
    E(0, 0, 1000, 0, 1000, 1),
    E(1000, 0, 2000, 0, 1000, 2),
    E(2000, 0, 3000, 0, 1000, 3),
    E(2000, 0, 2000, 1000, 1000, 3),
    E(3000, 0, 4000, 0, 1000, 4),
    E(4000, 0, 2000, 1000, 2236, 0, off=True),
    {"t": "C", "l": "_subgraph_head", "c": [2000, 1000], "r": 80.0, "i": 3, "dm": 3000},
    {"t": "C", "l": "_subgraph_head", "c": [4000, 0], "r": 80.0, "i": 4, "dm": 4000},
    {"t": "C", "l": "_alarm_valve", "c": [0, 0], "r": 150.0},
]

errors: list[str] = []
fails: list[str] = []


def check(label, got, want):
    ok = got == want
    if not ok:
        fails.append(f"{label}: got {got!r} want {want!r}")
    print(f"  {'OK ' if ok else 'X  '} {label}: {got!r}")


with sync_playwright() as pw:
    br = pw.chromium.launch()
    page = br.new_page(viewport={"width": 1400, "height": 900})
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.goto(f"http://127.0.0.1:{port}/data/_trace_check.html")
    page.wait_for_timeout(1200)

    page.evaluate("""(ents) => {
        state.stage_entities[4] = ents;
        state.bbox = {x_min: -200, y_min: -200, x_max: 4200, y_max: 1200};
        fitToBBox();
        setStageView(4);
        TracePlayer.start(ents);
    }""", ENTS)
    page.wait_for_timeout(200)

    print("① 재생 시작 직후")
    check("바 노출", page.is_visible("#trace-bar"), True)
    check("재생 중 버튼", page.inner_text("#trace-play"), "⏸")

    page.click("#trace-play")           # 일시정지
    page.wait_for_timeout(150)
    print("② 일시정지 + 처음으로")
    check("정지 버튼", page.inner_text("#trace-play"), "▶")
    page.click("#trace-restart")
    check("revealed=0", page.evaluate("() => state.revealed_edges_4"), 0)
    check("대기 문구", "알람밸브에서 시작" in page.inner_text("#trace-readout"), True)

    print("③ 한 배관씩 전진")
    for _ in range(3):
        page.click("#trace-next")
    check("revealed=3", page.evaluate("() => state.revealed_edges_4"), 3)
    r3 = page.inner_text("#trace-readout")
    check("배관 3/6", "배관 3/6" in r3, True)
    check("깊이 3", "깊이 3" in r3, True)
    check("누적 3.0 m", "AV→선단 3.0 m" in r3, True)
    check("헤드 0/2", "헤드 0/2" in r3, True)

    page.click("#trace-next")           # e3 = 분기 → 헤드 도달
    r4 = page.inner_text("#trace-readout")
    print(f"     4번째: {r4}")
    check("헤드 1/2", "헤드 1/2" in r4, True)
    check("헤드 도달 3.0 m", "헤드 도달 3.0 m" in r4, True)

    print("④ 스크럽 — 트리 밖 edge 경고")
    page.fill("#trace-scrub", "6")
    page.dispatch_event("#trace-scrub", "input")
    check("끝 = 완성망(null)", page.evaluate("() => state.revealed_edges_4"), None)
    r6 = page.inner_text("#trace-readout")
    print(f"     6번째: {r6}")
    check("트리 밖 경고", "트리 밖" in r6, True)
    page.click("#trace-prev")
    check("revealed=5", page.evaluate("() => state.revealed_edges_4"), 5)
    check("한 칸 뒤는 경고 없음", "트리 밖" in page.inner_text("#trace-readout"), False)

    print("⑤ 선단 따라가기 — 줌인 + 팬")
    before = page.evaluate("() => [state.view.zoom, state.view.panX]")
    page.check("#trace-follow")
    page.wait_for_timeout(150)
    after = page.evaluate("() => [state.view.zoom, state.view.panX]")
    check("줌 확대됨", after[0] > before[0], True)
    check("팬 이동됨", abs(after[1] - before[1]) > 1, True)
    page.uncheck("#trace-follow")

    print("⑥ 재생 재개 — 16× 로 끝까지")
    page.select_option("#trace-speed", "16")
    page.click("#trace-restart")
    page.click("#trace-play")
    page.wait_for_timeout(1500)
    check("끝까지 진행", page.evaluate("() => state.revealed_edges_4"), None)
    check("재생 종료 상태", page.inner_text("#trace-play"), "▶")

    print("⑦ 다른 뷰로 나가면 바 숨김")
    page.evaluate("() => { state.stage_entities[3] = [{t:'L',l:'x',p:[0,0,1,1]}]; setStageView(3); }")
    page.wait_for_timeout(150)
    check("바 숨김", page.is_visible("#trace-bar"), False)
    page.evaluate("() => setStageView(4)")
    page.wait_for_timeout(150)
    check("바 재노출", page.is_visible("#trace-bar"), True)

    print("⑧ 물 팔레트 — 노란색이 남아있지 않아야 한다")
    check("CAT_COLORS._subgraph", page.evaluate("() => CAT_COLORS._subgraph"), "#38bdf8")
    check("공유 riser 색 불변",
          page.evaluate("() => GRAPH_STYLE.color_riser"), "#fde047")
    water = page.evaluate("() => WATER")
    print(f"     WATER: {water}")
    check("물 팔레트 5색", sorted(water), ["crest", "deep", "halo", "shallow", "spray"])

    # 실제 캔버스 픽셀을 세어 물색이 칠해졌는지 확인 — 렌더 경로가 정말 탔다는 증거.
    page.click("#trace-restart")
    for _ in range(4):
        page.click("#trace-next")
    page.evaluate("() => { fitToBBox(); render(); }")   # ⑤의 4배 줌 해제
    px = page.evaluate("""() => {
        const cv = document.querySelector('#wb-canvas');
        const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
        let water = 0, yellow = 0;
        for (let i = 0; i < d.length; i += 4) {
            const [r, g, b, a] = [d[i], d[i+1], d[i+2], d[i+3]];
            if (a < 40) continue;
            if (b > 120 && b > r + 40 && g > r) water++;          // 물색 계열
            if (r > 180 && g > 170 && b < 120) yellow++;          // #fde047 계열
        }
        return {water, yellow};
    }""")
    print(f"     픽셀: {px}")
    check("물색 픽셀 존재", px["water"] > 200, True)
    check("노란 픽셀 없음", px["yellow"] < 50, True)
    page.screenshot(path=str(BASE / "data" / "_trace_water.png"))

    br.close()

httpd.shutdown()
out.unlink(missing_ok=True)

if errors:
    print("\nJS 오류:")
    for e in errors[:20]:
        print("  " + e)
if fails:
    print("\n실패:")
    for f in fails:
        print("  " + f)
print("\n" + ("PASS — 추적 재생 전 기능 정상"
             if not fails and not errors else "FAIL"))
sys.exit(0 if (not fails and not errors) else 1)
