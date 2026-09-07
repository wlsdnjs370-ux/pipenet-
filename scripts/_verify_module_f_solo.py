# -*- coding: utf-8 -*-
"""[§27] «이 묶음만 크게» 가 화면에서 정말 도는가 — 화소로 잰다.

이름 사전이 도면에 따라 정반대로 읽는다(계통도 `SP` = 헤드 918개). 판정을
늘리는 대신 사람이 **보게** 하기로 했으므로, 그 «보기» 가 실제로 작동하는지가
전부다: 눌렀을 때 다른 묶음이 사라지고 그 묶음이 커지는가.
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(f"{name} — {detail}")


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in open(p, encoding="utf-8"):
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
    pw = _password()
    from playwright.sync_api import sync_playwright

    dxf = os.path.join(_ROOT, "routes", "제출용[최종]",
                       "1. 입력도면 대명동 단위세대 계통도.dxf")
    if not os.path.isfile(dxf):
        print("표본 도면 없음 — 건너뜀")
        return 0

    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        errs: list[str] = []
        pg.on("console", lambda m: errs.append(m.text)
              if m.type == "error" else None)
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", pw)
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(1200)

        print("[1] 도면을 올려 묶음 목록까지")
        pg.set_input_files("#dxf", dxf)
        pg.click("#btn-open")
        got = False
        for _ in range(900):
            pg.wait_for_timeout(200)
            if pg.eval_on_selector_all("#layers label", "e => e.length") \
                    and pg.is_hidden("#busy"):
                got = True
                break
        check("묶음 목록이 만들어진다", got,
              "" if got else pg.inner_text("#stage") if
              pg.query_selector("#stage") else "단계 표시 없음")
        if not got:
            pg.screenshot(path=os.path.join(_ROOT, "data",
                                            "_solo_verify_fail.png"))
            br.close()
            return 1

        # 레이어 묶음은 「찍기」 단계에 붙어 있다 — 도면만 읽은 상태에서는
        # 목록이 만들어져 있어도 패널이 안 열려 있다.
        pg.evaluate(
            "() => [...document.querySelectorAll('#steps div')]"
            ".find(d => d.textContent.includes('찍기'))?.click()")
        pg.wait_for_timeout(700)
        check("찍기 단계에서 레이어 패널이 뜬다", pg.is_visible("#panel-layers"))

        # 레이어 접힘을 편다. ★`pg.click` 은 쓰지 않는다 — 덮개(#busy)가
        #   점멸하며 포인터를 가로채 재시도만 반복한다(실제로 그랬다).
        #   ★이미 펼쳐져 있으면 누르지 않는다 — 누르면 닫힌다.
        if pg.is_hidden("#layers"):
            pg.evaluate("() => document.querySelector("
                        "'h2.fold[data-fold=\"layers-body\"]').click()")
            pg.wait_for_timeout(500)
        check("레이어 목록이 펼쳐진다", pg.is_visible("#layers"))

        n_solo = pg.eval_on_selector_all("#layers .solo", "e => e.length")
        n_row = pg.eval_on_selector_all("#layers label", "e => e.length")
        check("묶음마다 «이 묶음만» 단추가 있다", n_solo == n_row and n_solo > 1,
              f"단추 {n_solo} · 묶음 {n_row}")
        if not n_solo:
            br.close()
            return 1

        def ink():
            return _ink(pg.query_selector("#cv").screenshot())

        print("[2] 화소 — 누르면 그 묶음만 남는가")
        pg.wait_for_timeout(1200)
        before = ink()
        v1 = pg.evaluate("() => ({...window.__mf.view})")
        checked_before = pg.eval_on_selector_all(
            "#layers input:checked", "e => e.length")
        pg.evaluate("() => document.querySelectorAll("
                    "'#layers .solo')[0].click()")
        pg.wait_for_timeout(1500)
        after = ink()
        checked_after = pg.eval_on_selector_all(
            "#layers input:checked", "e => e.length")
        check("다른 묶음이 꺼진다", checked_after == 1 and checked_before > 1,
              f"켜진 묶음 {checked_before} → {checked_after}")
        check("그림이 실제로 달라진다", after != before,
              f"화소 {before:,} → {after:,}")
        check("빈 화면이 되지 않는다", after > 0, f"{after:,} 화소")
        # 「크게 보기」이므로 시점이 그 묶음으로 옮겨가야 한다.
        v2 = pg.evaluate("() => ({...window.__mf.view})")
        moved = any(abs(v2[k] - v1[k]) > 1e-9 for k in ("scale", "ox", "oy"))
        check("그 묶음 범위로 화면이 옮겨간다", moved,
              f"배율 {v1['scale']:.6g} → {v2['scale']:.6g}")

        print("[3] 다시 누르면 되돌아온다")
        pg.evaluate("() => document.querySelectorAll("
                    "'#layers .solo')[0].click()")
        pg.wait_for_timeout(1500)
        back = pg.eval_on_selector_all("#layers input:checked", "e => e.length")
        check("모든 묶음이 다시 켜진다", back == checked_before,
              f"{checked_after} → {back} (처음 {checked_before})")

        print("[4] 콘솔 오류")
        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:2]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("묶음 미리보기 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
