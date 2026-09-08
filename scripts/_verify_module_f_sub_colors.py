# -*- coding: utf-8 -*-
"""[계통도·기계실] 도면 색 그대로 보이는가 — 화면 화소로 센다.

사용자 지시: 「캐드 도면의 색깔별로 보이게 해서(기계실도 같은 논리로) 조금
구분이 쉽게 조치해 줬으면 좋겠다.」

종전에는 모든 도형에 색 7 을 박아 계통도·기계실이 통째로 한 색이었다. 배관·
기호·건축선이 눈으로 안 갈렸다. 여기서는 **화면에 실제로 몇 가지 색이
찍히는지** 를 센다 — 묶음 payload 만 보면 «색을 실었다» 까지밖에 모른다.
"""
from __future__ import annotations

import io
import os
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}"
          + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(f"{name} — {detail}")


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def _palette(png):
    """캔버스에 실제로 찍힌 색 — 배경과 다른 화소를 색깔별로 센다."""
    from PIL import Image
    im = Image.open(io.BytesIO(png)).convert("RGB")
    px, (w, h) = im.load(), im.size
    bg = px[1, 1]
    seen: Counter = Counter()
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            c = px[x, y]
            if abs(c[0] - bg[0]) + abs(c[1] - bg[1]) + abs(c[2] - bg[2]) > 40:
                # 24단계로 뭉쳐 센다 — 안티에일리어싱 화소를 딴 색으로 세면
                # «색이 많다» 가 거짓으로 참이 된다.
                seen[(c[0] // 24, c[1] // 24, c[2] // 24)] += 1
    return seen


def main() -> int:
    from playwright.sync_api import sync_playwright

    plan = os.path.join(_ROOT, "routes", "제출용[최종]",
                        "1. 입력도면 대명동 단위세대 평면도.dxf")
    subs = [("계통도", "1. 입력도면 대명동 단위세대 계통도.dxf"),
            ("기계실", "1. 입력도면 대명동 단위세대 기계실.dxf")]
    with sync_playwright() as p:
        br = p.chromium.launch()
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

        def idle():
            for _ in range(3000):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        print("[0] 평면도 한 장 — 슬롯 탭을 띄우려고")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)

        for label, fname in subs:
            path = os.path.join(_ROOT, "routes", "제출용[최종]", fname)
            if not os.path.isfile(path):
                print(f"  [건너뜀] {label} 표본 없음")
                continue
            print(f"[{label}]")
            pg.wait_for_selector("#slots button", timeout=120_000)
            pg.click(f'#slots button:has-text("{label}")')
            pg.wait_for_timeout(1200)
            pg.set_input_files("#dxf", path)
            pg.click("#btn-open")
            idle()
            pg.wait_for_timeout(1800)

            pal = _palette(pg.query_selector("#cv").screenshot())
            big = [c for c, n in pal.items() if n >= 30]
            check("★도면 색이 여러 가지로 보인다", len(big) >= 3,
                  f"뚜렷한 색 {len(big)}가지 (전체 {len(pal)})")
            # 묶음도 색이 갈려야 한다 — 한 색이면 레이어 토글이 뭉친다.
            cols = pg.evaluate("""() => {
                const b = (window.__mf.world || {}).bundles || [];
                return [...new Set(b.map((x) => x.color))];
            }""")
            check("묶음이 색으로도 갈린다", len(cols) >= 3,
                  f"색 {sorted(cols)}")

        print("[콘솔]")
        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:2]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("계통도·기계실 도면 색 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
