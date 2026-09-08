# -*- coding: utf-8 -*-
"""[진단] 새 흐름 — 찍기에서 «수동으로 찍고» 넘어간 뒤 최불리가 도는가.

착지가 손질에서 찍기로 바뀐 뒤(D-F10-3 개정), 사람은 찍기에서 손을 대고
「배관망 구성 →」으로 올라간다. 사용자 지적은 그 뒤 「최불리 선정이 또
작동 안 한다」이다. 그 순서를 그대로 태우고, 각 걸음의 상태를 찍는다.
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

    plan = os.path.join(_ROOT, "routes", "제출용[최종]",
                        "1. 입력도면 대명동 단위세대 평면도.dxf")
    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        msgs: list = []
        pg.on("console", lambda m: msgs.append(f"[{m.type}] {m.text}"))
        pg.on("pageerror", lambda e: msgs.append(f"[pageerror] {e}"))
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(1000)

        def idle():
            for _ in range(3000):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        print(f"[1] 착지 = {pg.evaluate('() => window.__mf.stage')}")

        print("[2] 찍기에서 «수동으로» 몇 군데 찍는다")
        box = pg.eval_on_selector("#cv", """e => {
            const r = e.getBoundingClientRect();
            return {x: r.x, y: r.y, w: r.width, h: r.height};
        }""")
        def pk():
            return pg.evaluate("""() => {
                const S = window.__mf, p = S.pick || {};
                return {mats: (p.materials || p.mat || []).length,
                        adopted: S.adopted ? S.adopted.size : null,
                        suggest: (S.suggest || []).length,
                        keys: Object.keys(p)};
            }""")
        print(f"    찍기 전 : {pk()}")
        print(f"    ★헤드 수 표시: «{pg.inner_text('#pk-count')[:70]}»")
        for fx, fy in ((0.35, 0.4), (0.5, 0.5), (0.62, 0.55)):
            pg.mouse.click(box["x"] + box["w"] * fx, box["y"] + box["h"] * fy)
            pg.wait_for_timeout(500)
            print(f"    클릭 후 : {pg.inner_text('#status')[:74]}")
            print(f"              표시 «{pg.inner_text('#pk-count')[:70]}»")

        print("[3] 「배관망 구성 →」")
        pg.click("#pk-next")
        for _ in range(3000):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy") and pg.evaluate(
                    "() => window.__mf.stage") == "edit":
                break
        pg.wait_for_timeout(1500)
        st = pg.evaluate("() => window.__mf.stage")
        print(f"    단계 = {st} · 상태줄 {pg.inner_text('#status')[:80]}")
        if st != "edit":
            print("    ★손질로 못 올라갔다 — 여기가 문제다")
            print("    콘솔:", [m[:160] for m in msgs[-4:]])
            br.close()
            return 1

        print("[4] 알람밸브 클릭")
        pt = pg.evaluate("""() => {
            for (const g of (window.__mf.edit.body_groups || [])) {
                const s = g.segs || [];
                if (s.length >= 4) return [s[0], s[1]];
            }
            return null;
        }""")
        scr = pg.evaluate("""(q) => {
            const v = window.__mf.view;
            const r = document.querySelector('#cv').getBoundingClientRect();
            return [r.x + (q[0] - v.ox) * v.scale,
                    r.y + r.height - (q[1] - v.oy) * v.scale];
        }""", pt)
        pg.mouse.click(scr[0], scr[1])
        idle()
        pg.wait_for_timeout(900)
        e = pg.evaluate("() => window.__mf.edit")
        print(f"    급수원 {len(e.get('sources') or [])}곳"
              f" · 영역그리기 {pg.is_checked('#ed-zone-arm')}"
              f" · 안내 «{pg.inner_text('#ed-worst-why')[:44]}»")

        print("[5] ★최불리 선정")
        pg.click("#ed-worst")
        idle()
        pg.wait_for_timeout(1500)
        w = pg.evaluate("() => (window.__mf.edit || {}).worst || null")
        print(f"    worst = {(w or {}).get('k')}")
        print(f"    상태줄: {pg.inner_text('#status')[:120]}")
        print("[6] 콘솔")
        for m in msgs[-6:]:
            print("   ", m[:180])
        pg.screenshot(path=os.path.join(_ROOT, "data", "_worst_manual.png"))
        br.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
