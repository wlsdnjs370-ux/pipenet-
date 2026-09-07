# -*- coding: utf-8 -*-
"""[진단] 최불리 선정 → 해제 → 영역 지정 → 다시 선정.

사용자 지적: 「최불리 선정 후에 해제한 뒤 영역 지정 다시 하고, 최불리 선정하니까
또 작동이 안 된다」. 그 순서를 그대로 태워 어디서 끊기는지 본다.
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
        msgs: list[str] = []
        pg.on("console", lambda m: msgs.append(f"[{m.type}] {m.text}"))
        pg.on("pageerror", lambda e: msgs.append(f"[pageerror] {e}"))
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(1000)
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        for _ in range(3000):
            pg.wait_for_timeout(200)
            if pg.is_visible("#panel-edit") and pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(1500)

        def idle():
            for _ in range(900):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return
        def w():
            return pg.evaluate("() => (window.__mf.edit || {}).worst || null")
        def st():
            return pg.inner_text("#status")[:110]

        # ① 알람밸브
        pt = pg.evaluate("""() => {
            for (const g of (window.__mf.edit.body_groups || [])) {
                const s = g.segs || [];
                if (s.length >= 4) return [s[0], s[1]];
            }
            return null;
        }""")
        scr = pg.evaluate("""(p) => {
            const v = window.__mf.view;
            const r = document.querySelector('#cv').getBoundingClientRect();
            return [r.x + (p[0] - v.ox) * v.scale,
                    r.y + r.height - (p[1] - v.oy) * v.scale];
        }""", pt)
        pg.mouse.click(scr[0], scr[1])
        idle()
        pg.wait_for_timeout(600)
        print(f"[1] 알람밸브 {len((pg.evaluate('() => window.__mf.edit') or {}).get('sources') or [])}곳")

        print("[2] 최불리 선정")
        pg.click("#ed-worst")
        idle()
        pg.wait_for_timeout(800)
        print(f"    worst={(w() or {}).get('k')} · {st()}")

        print("[3] 해제")
        pg.click("#ed-worst-clear")
        pg.wait_for_timeout(900)
        e = pg.evaluate("() => window.__mf.edit")
        print(f"    worst={(w() or {}).get('k')}"
              f" · sources={len(e.get('sources') or [])}"
              f" · valves={len(e.get('valves') or [])}"
              f" · keep={e.get('keep')}")
        print(f"    안내 «{pg.inner_text('#ed-worst-why')[:50]}»")

        print("[4] 영역 지정 — 도면 절반을 사각형으로")
        pg.check("#ed-zone-arm")
        pg.wait_for_timeout(300)
        box = pg.eval_on_selector("#cv", """e => {
            const r = e.getBoundingClientRect();
            return {x: r.x, y: r.y, w: r.width, h: r.height};
        }""")
        pg.mouse.move(box["x"] + box["w"] * 0.08, box["y"] + box["h"] * 0.08)
        pg.mouse.down()
        pg.mouse.move(box["x"] + box["w"] * 0.92, box["y"] + box["h"] * 0.92,
                      steps=12)
        pg.mouse.up()
        pg.wait_for_timeout(700)
        print(f"    zones={pg.evaluate('() => window.__mf.zones.length')}"
              f" · 목록 «{pg.inner_text('#ed-zones')[:50]}»")
        print(f"    안내 «{pg.inner_text('#ed-worst-why')[:60]}»")
        diag = pg.evaluate("""() => {
            const S = window.__mf, e = S.edit, z = S.zones[0];
            const hs = e.heads || [];
            let inz = 0, minx = 1e18, miny = 1e18, maxx = -1e18, maxy = -1e18;
            for (const h of hs) {
                const x = h[0], y = h[1];
                if (x < minx) minx = x; if (y < miny) miny = y;
                if (x > maxx) maxx = x; if (y > maxy) maxy = y;
                if (z && x >= z[0] && x <= z[2] && y >= z[1] && y <= z[3]) inz++;
            }
            return {zone: z, heads: hs.length, in_zone: inz,
                    head_box: [minx, miny, maxx, maxy],
                    src: (e.sources || [])[0] || null,
                    bounds: e.bounds};
        }""")
        print(f"    진단 zone={diag['zone']}")
        print(f"          헤드 {diag['heads']}개 · 영역 안 {diag['in_zone']}개")
        print(f"          헤드범위 {diag['head_box']}")
        print(f"          급수원 {diag['src']} · bounds {diag['bounds']}")

        print("[5] ★다시 최불리 선정")
        pg.click("#ed-worst")
        idle()
        pg.wait_for_timeout(1200)
        print(f"    worst={(w() or {}).get('k')} · {st()}")

        print("[6] 콘솔")
        for m in msgs[-8:]:
            print("   ", m[:200])
        pg.screenshot(path=os.path.join(_ROOT, "data", "_worst_after_clear.png"))
        br.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
