# -*- coding: utf-8 -*-
"""[계통도 꼬임] 화면에서 **배관 위를 찍었을 때** 계통이 좁혀지는가.

전 공정 검증(`_verify_module_f_pipeline.py`)은 캔버스의 정해진 자리를 찍는다 —
배관에서 멀리 떨어진 자리라 «두 점이 서로 다른 계통에 붙는다» 로 좁히기가
거절된다(그것도 맞는 동작이다). 여기서는 사람이 하듯 **배관 위**를 찍는다:
그래프 절점의 아래끝·위끝을 화면 좌표로 바꿔 클릭하고, 좁혀졌는지 · 뽑힌
연장이 좁힌 값과 같은지를 본다.
"""
from __future__ import annotations

import io
import sys
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_D = os.path.join(_ROOT, "routes", "제출용[최종]")
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(name)
    return ok


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def pick_points(dxf_path):
    """한 계통(첫 자동 배관 레이어)의 아래끝·위끝 — 살짝 빗나가게."""
    from remote30_prototype import _auto_pipe_layer_filter, build_system_graph
    from routes.module_f.subdrawing import parse_subdrawing
    ents, _ = parse_subdrawing(dxf_path)
    auto = sorted(_auto_pipe_layer_filter(ents))
    g, _el, _st = build_system_graph(ents, layer_filter={auto[0]},
                                     force_connect=True)
    ns = sorted(g, key=lambda n: n[1])
    return ([ns[0][0] + 120.0, ns[0][1] - 90.0],
            [ns[-1][0] - 80.0, ns[-1][1] + 140.0])


def main() -> int:
    from playwright.sync_api import sync_playwright

    files = {"계통도": os.path.join(_D, "1. 입력도면 대명동 단위세대 계통도.dxf"),
             "기계실": os.path.join(_D, "1. 입력도면 대명동 단위세대 기계실.dxf")}
    for p in files.values():
        if not os.path.isfile(p):
            print(f"표본 없음: {p}")
            return 0

    with sync_playwright() as pw:
        br = pw.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        errs: list[str] = []
        pg.on("console", lambda m: errs.append(m.text)
              if m.type == "error" else None)
        pg.on("pageerror", lambda e: errs.append(f"pageerror: {e}"))
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(1000)

        def idle(n=6000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        def click_world(q):
            scr = pg.evaluate("""(q) => {
                const v = window.__mf.view;
                const r = document.querySelector('#cv').getBoundingClientRect();
                return [r.x + (q[0] - v.ox) * v.scale,
                        r.y + r.height - (q[1] - v.oy) * v.scale];
            }""", list(q))
            pg.mouse.click(scr[0], scr[1])

        # 슬롯 탭은 «도면을 하나라도 연 뒤» 에 켜진다 — 평면도를 먼저 연다.
        plan = os.path.join(_D, "1. 입력도면 대명동 단위세대 평면도.dxf")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        check("슬롯 탭이 켜진다", pg.locator("#slots button").count() > 0)

        for label, path in files.items():
            print(f"\n=== {label}")
            pg.wait_for_selector("#slots button", timeout=120_000)
            pg.click(f'#slots button:has-text("{label}")')
            pg.wait_for_timeout(1000)
            pg.set_input_files("#dxf", path)
            pg.click("#btn-open")
            idle()
            pg.wait_for_timeout(1500)
            g = pg.evaluate("() => window.__mf.subGraph || null")
            if not check("경로 그래프를 받았다", bool(g and g.get("nodes"))):
                continue
            print(f"      섞인 배관 레이어: {g.get('auto_layers')}")
            # ★찍을 자리는 «한 계통의 두 끝» 이다 — 사람은 한 입상관의 아래·위를
            #   찍지, 고층 아래끝과 저층 위끝을 찍지 않는다. 섞인 그래프의 최저·
            #   최고 절점을 쓰면 서로 다른 계통이 뽑혀 좁히기가 (맞게) 거절된다.
            a, b = pick_points(path)
            print(f"      찍는 자리: {a} → {b}")
            pg.click("#sub-pick-a")
            click_world(a)
            pg.wait_for_timeout(400)
            pg.click("#sub-pick-b")
            click_world(b)
            idle()
            pg.wait_for_timeout(1200)
            st = pg.evaluate("""() => { const g = window.__mf.subGraph || {};
                return {chosen: g.chosen, auto: g.chosen_auto,
                        why: (g.narrowed || {}).reason,
                        nodes: (g.nodes || []).length}; }""")
            print(f"      좁히기: {st}")
            check(f"{label} — 한 계통으로 좁혔다", bool(st["auto"] and st["chosen"]),
                  str(st["why"]))
            panel = pg.inner_text("#sub-layers")
            check(f"{label} — 화면이 그 사실을 말한다",
                  ("계통 안에서만" in panel) or ("좁히지 못했습니다" in panel),
                  panel[:70].replace("\n", " "))
            pg.click("#sub-extract")
            idle()
            pg.wait_for_timeout(1500)
            s = pg.evaluate("() => (window.__mf.sub || {}).summary || null")
            check(f"{label} — 경로가 뽑힌다", bool(s and s.get("pipes")),
                  pg.inner_text("#status")[:70])
            if s:
                print(f"      절점 {s['nodes']} · 배관 {s['pipes']}"
                      f" · 연장 {s['total_m']} m")

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("계통 좁히기 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
