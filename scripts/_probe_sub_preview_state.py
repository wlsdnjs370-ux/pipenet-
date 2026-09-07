# -*- coding: utf-8 -*-
"""[§27 후속] 두 점 미리보기가 «왜» 안 그려지나 — 화면 상태를 직접 본다.

`_verify_module_f_sub_live.py` 가 「선이 안 따라온다」를 3건 잡았다. 화소만으로는
어디서 끊겼는지 모른다: 그래프를 못 받았나 · 찍기가 안 armed 인가 · 경로를 못
푸나 · 풀었는데 안 그리나. `window.__mf` 를 그대로 읽어 갈라 본다.
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def main() -> int:
    from playwright.sync_api import sync_playwright

    dxf = os.path.join(_ROOT, "routes", "제출용[최종]",
                       "1. 입력도면 대명동 단위세대 계통도.dxf")
    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        errs = []
        pg.on("console", lambda m: errs.append(m.text)
              if m.type == "error" else None)
        pg.on("pageerror", lambda e: errs.append(f"pageerror: {e}"))
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(1200)

        # ★슬롯 탭은 «도면을 연 뒤에야» 생긴다 — 첫 업로드는 늘 평면도로
        #   간다. 사람이 하는 그대로 평면도를 먼저 올리고 계통도로 바꾼다.
        plan = os.path.join(_ROOT, "routes", "제출용[최종]",
                            "1. 입력도면 대명동 단위세대 평면도.dxf")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        pg.wait_for_function(
            "() => document.querySelector('#busy').classList"
            ".contains('hidden')", timeout=900_000)
        pg.wait_for_selector('#slots button', timeout=120_000)
        pg.click('#slots button:has-text("계통도")')
        pg.wait_for_timeout(1200)
        pg.set_input_files("#dxf", dxf)
        pg.click("#btn-open")
        for _ in range(1800):
            pg.wait_for_timeout(200)
            if pg.is_visible("#panel-sub") and pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(1500)

        box = pg.eval_on_selector("#cv", """e => {
            const r = e.getBoundingClientRect();
            return {x: r.x, y: r.y, w: r.width, h: r.height};
        }""")

        def state(tag):
            s = pg.evaluate("""() => {
                const S = window.__mf;
                const g = S.subGraph;
                return {
                    graph: g ? {n: g.nodes.length, e: g.edges.length,
                                pen: g.forced_penalty_mm,
                                adj: !!g.adj} : null,
                    arm: S.sub ? S.sub.arm : "sub 없음",
                    picks: S.sub ? S.sub.picks : null,
                    preview: S.sub && S.sub.preview
                             ? S.sub.preview.length : null,
                    broken: S.sub ? S.sub.previewBroken : null,
                    stage: S.stage,
                };
            }""")
            print(f"  {tag:<22} {s}")
            return s

        print("[1] 도면을 연 직후")
        state("열린 뒤")

        print("[2] ① 단추를 누르고 캔버스 한가운데를 찍는다")
        pg.click("#sub-pick-a")
        pg.wait_for_timeout(200)
        state("① armed")
        pg.mouse.click(box["x"] + box["w"] * 0.5, box["y"] + box["h"] * 0.5)
        pg.wait_for_timeout(400)
        state("① 찍은 뒤")

        print("[3] 마우스를 옮긴다")
        pg.mouse.move(box["x"] + box["w"] * 0.30, box["y"] + box["h"] * 0.30)
        pg.wait_for_timeout(500)
        s = state("마우스 이동 뒤")

        if s["preview"]:
            print("  → 경로는 풀렸다. 그런데 화소가 0 이면 «그리는» 쪽 문제다.")
            drawn = pg.evaluate("""() => {
                const cv = document.querySelector('#cv');
                const ctx = cv.getContext('2d');
                const d = ctx.getImageData(0, 0, cv.width, cv.height).data;
                let n = 0;
                for (let i = 0; i < d.length; i += 4) {
                    if (d[i] > 200 && d[i+1] < 90 && d[i+2] < 90) n++;
                }
                return n;
            }""")
            print(f"  → 캔버스에서 붉은 화소 {drawn}")
        else:
            print("  → 경로 자체가 안 풀렸다(preview=null).")

        print("[4] 콘솔/페이지 오류")
        print("  ", [e[:160] for e in errs[:4]] or "없음")
        pg.screenshot(path=os.path.join(_ROOT, "data", "_sub_preview_state.png"))
        br.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
