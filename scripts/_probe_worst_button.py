# -*- coding: utf-8 -*-
"""[진단] 「최불리 선정」 **버튼** 이 실제로 도는가.

앞선 검증은 서버 경로(fetch)를 직접 태워서 통과했다. 그래서 «버튼» 자체가
죽은 것을 못 봤다 — 사람이 쓰는 것은 버튼이다. 여기서는 캔버스를 클릭해
알람밸브를 놓고, **버튼을 눌러** 무슨 일이 나는지 본다.
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
        print(f"[1] 손질 진입 · 버튼 disabled={pg.is_disabled('#ed-worst')}"
              f" · 안내 «{pg.inner_text('#ed-worst-why')[:40]}»")

        print("[1b] ★알람밸브 전에 눌러 본다 — 죽은 단추면 안 된다")
        pg.click("#ed-worst")
        pg.wait_for_timeout(700)
        print(f"    상태줄: {pg.inner_text('#status')[:80]}")

        # 캔버스에서 배관 위 한 점을 골라 «사람처럼» 클릭한다.
        pt = pg.evaluate("""() => {
            const e = window.__mf.edit;
            for (const g of (e.body_groups || [])) {
                const s = g.segs || [];
                if (s.length >= 4) return [s[0], s[1]];
            }
            return null;
        }""")
        print(f"[2] 배관 점 {pt}")
        scr = pg.evaluate("""(w) => {
            const S = window.__mf, v = S.view;
            const cv = document.querySelector('#cv');
            const r = cv.getBoundingClientRect();
            return [r.x + (w[0] - v.ox) * v.scale,
                    r.y + r.height - (w[1] - v.oy) * v.scale];
        }""", pt)
        pg.mouse.click(scr[0], scr[1])
        for _ in range(600):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(800)
        st = pg.evaluate("() => window.__mf.edit")
        print(f"[3] 알람밸브 {len(st.get('sources') or [])}곳"
              f" · 버튼 disabled={pg.is_disabled('#ed-worst')}"
              f" · 안내 «{pg.inner_text('#ed-worst-why')[:44]}»")

        print("[4] ★버튼을 «누른다»")
        before = pg.evaluate("() => !!(window.__mf.edit || {}).worst")
        try:
            pg.click("#ed-worst", timeout=8000)
            clicked = True
        except Exception as exc:            # noqa: BLE001
            clicked = False
            print(f"    누를 수가 없다: {str(exc)[:200]}")
        if clicked:
            for _ in range(900):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    break
            pg.wait_for_timeout(1000)
        after = pg.evaluate("() => (window.__mf.edit || {}).worst || null")
        print(f"[5] 눌렀나={clicked} · 전 worst={before} · 후 worst="
              f"{(after or {}).get('k')}")
        print(f"    상태줄: {pg.inner_text('#status')[:120]}")
        print("[6] 콘솔")
        for m in msgs[-8:]:
            print("   ", m[:200])
        pg.screenshot(path=os.path.join(_ROOT, "data", "_worst_btn.png"))
        br.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
