# -*- coding: utf-8 -*-
"""[진단] 슬롯을 오갔다 돌아오면 계통도·기계실 도면이 사라지는가.

사용자 지적: 「계통도·기계실 업로드 후 다시 평면도로 돌아갔다가 계통도·기계실로
돌아가면 그 도면이 아예 스크린에서 사라진다.」

화소로 잰다 — 처음 열었을 때와 돌아왔을 때의 캔버스를 견준다.
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


def _ink(png):
    from PIL import Image
    im = Image.open(io.BytesIO(png)).convert("RGB")
    px, (w, h) = im.load(), im.size
    bg = px[1, 1]
    n = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            c = px[x, y]
            if abs(c[0] - bg[0]) + abs(c[1] - bg[1]) + abs(c[2] - bg[2]) > 24:
                n += 1
    return n


def main() -> int:
    from playwright.sync_api import sync_playwright

    plan = os.path.join(_ROOT, "routes", "제출용[최종]",
                        "1. 입력도면 대명동 단위세대 평면도.dxf")
    sysd = os.path.join(_ROOT, "routes", "제출용[최종]",
                        "1. 입력도면 대명동 단위세대 계통도.dxf")
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

        def idle(n=3000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        def state(tag):
            s = pg.evaluate("""() => {
                const S = window.__mf, w = S.world;
                return {slot: S.slot, stage: S.stage,
                        world: w ? {segs: (w.segs || []).length,
                                    bundles: (w.bundles || []).length,
                                    bounds: !!w.bounds} : null,
                        view: {...S.view}};
            }""")
            ink = _ink(pg.query_selector("#cv").screenshot())
            print(f"  {tag:<26} 화소 {ink:>7,} · slot={s['slot']}"
                  f" · stage={s['stage']} · world={s['world']}")
            return ink, s

        print("[1] 평면도 업로드")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        state("평면도 처음")

        print("[2] 계통도 슬롯으로 바꿔 업로드")
        pg.wait_for_selector("#slots button", timeout=120_000)
        pg.click('#slots button:has-text("계통도")')
        pg.wait_for_timeout(1200)
        pg.set_input_files("#dxf", sysd)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        first, _ = state("계통도 처음")

        print("[3] 기계실 슬롯으로 바꿔 업로드")
        mr = os.path.join(_ROOT, "routes", "제출용[최종]",
                          "1. 입력도면 대명동 단위세대 기계실.dxf")
        pg.click('#slots button:has-text("기계실")')
        pg.wait_for_timeout(1200)
        pg.set_input_files("#dxf", mr)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        mr_first, _ = state("기계실 처음")

        print("[3b] 평면도로 돌아간다")
        pg.click('#slots button:has-text("평면도")')
        idle()
        pg.wait_for_timeout(1500)
        state("평면도 복귀")

        print("[4] ★다시 계통도 → 기계실 — 여기서 사라진다는 지적")
        pg.click('#slots button:has-text("계통도")')
        idle()
        pg.wait_for_timeout(1800)
        again, s2 = state("계통도 복귀")
        pg.click('#slots button:has-text("기계실")')
        idle()
        pg.wait_for_timeout(1800)
        mr_again, _ = state("기계실 복귀")
        if mr_first and mr_again < mr_first * 0.3:
            print(f"  ★기계실 재현됨 — {mr_first:,} → {mr_again:,}")
        else:
            print(f"  기계실 — {mr_first:,} → {mr_again:,}")

        print()
        want = {"plan": 30, "system": 24, "machineroom": 19}
        cur = pg.evaluate("() => ({slot: window.__mf.slot,"
                          " n: (window.__mf.world || {}).bundles ?"
                          " window.__mf.world.bundles.length : null})")
        print(f"  마지막 슬롯 {cur['slot']} · 묶음 {cur['n']}"
              f" (제 값 {want.get(cur['slot'])})")
        if first and again < first * 0.3:
            print(f"  ★재현됨 — 화소 {first:,} → {again:,}"
                  f" ({again / max(first, 1) * 100:.0f}%)")
        else:
            print(f"  재현 안 됨 — 화소 {first:,} → {again:,}")
        print("  콘솔:", [m[:150] for m in msgs[-5:]] or "없음")
        pg.screenshot(path=os.path.join(_ROOT, "data", "_slot_roundtrip.png"))
        br.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
