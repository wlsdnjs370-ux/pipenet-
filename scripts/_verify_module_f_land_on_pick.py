# -*- coding: utf-8 -*-
"""[D-F10-3 개정] 업로드하면 **찍기** 에서 멈추는가 — 살아 있는 서버로.

사용자 지시: 「업로드 후 자동까지는 좋은데, 3.손질로 넘어가지 말고 2.찍기로
넘어가서 수동 지정하기 더 쉽게 조치하자.」

세 가지를 본다.
  ① 도면을 올리면 **찍기** 화면에서 멈춘다(손질로 안 올라간다)
  ② 자동이 찍은 것이 그대로 있다 — 채택 수가 화면에 뜨고 도면이 보인다
  ③ 「배관망 구성 →」을 누르면 그때 손질로 간다
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
    if not os.path.isfile(plan):
        print("표본 도면 없음 — 건너뜀")
        return 0

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

        print("[1] 도면을 올린다 — 클릭은 «불러오기» 한 번뿐")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        for _ in range(3000):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy") and pg.evaluate(
                    "() => window.__mf.stage") in ("pick", "edit"):
                break
        pg.wait_for_timeout(2000)
        stage = pg.evaluate("() => window.__mf.stage")
        check("★찍기에서 멈춘다", stage == "pick", f"도착 = {stage}")
        check("찍기 판이 보인다", pg.is_visible("#panel-pick"))
        check("도면이 그려져 있다",
              _ink(pg.query_selector("#cv").screenshot()) > 500)

        print("[2] 자동이 찍은 것이 그대로 있다")
        note = pg.inner_text("#start-note")
        check("배너가 채택 결과를 말한다", "채택" in note,
              note.replace("\n", " ")[:70])
        check("다음에 무엇을 누를지 알려 준다", "배관망 구성" in note)
        check("채택 상자가 열려 있다", pg.is_visible("#pk-adopt-box"))

        print("[3] 「배관망 구성 →」을 누르면 그때 손질로")
        check("그 단추가 보인다", pg.is_visible("#pk-next"))
        pg.click("#pk-next")
        for _ in range(3000):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy") and pg.evaluate(
                    "() => window.__mf.stage") == "edit":
                break
        pg.wait_for_timeout(1200)
        check("★그때 손질로 간다",
              pg.evaluate("() => window.__mf.stage") == "edit",
              pg.evaluate("() => window.__mf.stage"))
        check("손질 판이 보인다", pg.is_visible("#panel-edit"))

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
    print("업로드 → 찍기에서 멈춘다 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
