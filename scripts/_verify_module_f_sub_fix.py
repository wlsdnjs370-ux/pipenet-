# -*- coding: utf-8 -*-
"""[§27 후속] 뽑힌 배관을 «화면에서» 고칠 수 있나 — 살아 있는 서버로.

시험은 합성 판으로 본다(`tests/test_module_f_sub_fix.py`). 여기서 보는 것은
사람이 실제로 하는 순서 그대로다: 계통도를 열고 → 두 점을 찍어 뽑고 →
추측 관경만 골라 → 고치고 → 다시 뽑아도 손질이 살아 있나.

★마지막이 핵심이다. 안 되붙이면 추출을 한 번 더 누른 순간 손질이 통째로
  사라지는데, 그 사실은 화면 어디에도 안 뜬다.
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


def main() -> int:
    from playwright.sync_api import sync_playwright

    plan = os.path.join(_ROOT, "routes", "제출용[최종]",
                        "1. 입력도면 대명동 단위세대 평면도.dxf")
    dxf = os.path.join(_ROOT, "routes", "제출용[최종]",
                       "1. 입력도면 대명동 단위세대 계통도.dxf")
    if not (os.path.isfile(plan) and os.path.isfile(dxf)):
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
        pg.wait_for_timeout(1200)

        print("[1] 계통도까지 — 첫 업로드는 늘 평면도 슬롯이다")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        pg.wait_for_function(
            "() => document.querySelector('#busy').classList"
            ".contains('hidden')", timeout=900_000)
        pg.wait_for_selector("#slots button", timeout=120_000)
        pg.click('#slots button:has-text("계통도")')
        pg.wait_for_timeout(1200)
        pg.set_input_files("#dxf", dxf)
        pg.click("#btn-open")
        for _ in range(2400):
            pg.wait_for_timeout(200)
            if pg.is_visible("#panel-sub") and pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(1500)
        check("계통도 화면이 열린다", pg.is_visible("#panel-sub"))

        print("[2] 두 점을 찍어 뽑는다")
        box = pg.eval_on_selector("#cv", """e => {
            const r = e.getBoundingClientRect();
            return {x: r.x, y: r.y, w: r.width, h: r.height};
        }""")
        pg.click("#sub-pick-a")
        pg.mouse.click(box["x"] + box["w"] * 0.5, box["y"] + box["h"] * 0.80)
        pg.wait_for_timeout(300)
        pg.click("#sub-pick-b")
        pg.mouse.click(box["x"] + box["w"] * 0.5, box["y"] + box["h"] * 0.20)
        pg.wait_for_timeout(300)
        pg.click("#sub-extract")
        for _ in range(2400):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(1200)

        pg.evaluate("() => document.querySelector("
                    "'h2.fold[data-fold=\"sub-fix-body\"]').click()")
        pg.wait_for_timeout(500)
        n_row = pg.eval_on_selector_all("#sub-fix-grid tr[data-k]",
                                        "e => e.length")
        chip = pg.inner_text("#sub-fix-chip")
        check("뽑힌 배관표가 보인다", n_row > 0, f"{n_row}행 · {chip}")
        if not n_row:
            pg.screenshot(path=os.path.join(_ROOT, "data", "_subfix_fail.png"))
            br.close()
            return 1
        before = pg.evaluate("() => window.__mf.sub.summary.total_m")

        print("[3] 추측 관경만 골라 고친다")
        pg.click("#sub-fix-guessed")
        pg.wait_for_timeout(300)
        n_sel = pg.inner_text("#sub-fix-n")
        check("추측 관경만 골라진다", "0" not in n_sel.split()[-1], n_sel)
        pg.fill("#sub-fix-dia", "100")
        pg.fill("#sub-fix-note", "도면 치수 확인 — 100A")
        pg.click("#sub-fix-save")
        pg.wait_for_timeout(1500)
        rows = pg.evaluate("() => window.__mf.sub && window.__mf.sub.summary")
        txt = pg.inner_text("#sub-fix-grid")
        check("표에 «원래 값» 이 같이 남는다", "원래 150" in txt,
              txt.split("\n")[1][:80] if "\n" in txt else txt[:80])
        check("근거가 «직접 입력» 으로 바뀐다", "직접 입력" in txt)
        n_fixed = pg.eval_on_selector_all(
            "#sub-fix-grid tr", "e => e.filter(t =>"
            " t.textContent.includes('직접 입력')).length")
        check("고친 행이 여러 개다", n_fixed > 1, f"{n_fixed}행")
        check("연장은 관경만 고쳐선 안 바뀐다",
              rows and abs(float(rows["total_m"]) - float(before)) < 1e-6,
              f"{before} → {rows and rows['total_m']}")

        print("[4] ★다시 뽑아도 손질이 살아 있나")
        pg.click("#sub-extract")
        for _ in range(2400):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(1500)
        txt2 = pg.inner_text("#sub-fix-grid")
        again = pg.eval_on_selector_all(
            "#sub-fix-grid tr", "e => e.filter(t =>"
            " t.textContent.includes('직접 입력')).length")
        check("다시 뽑아도 고친 값이 되붙는다", again == n_fixed,
              f"{n_fixed}행 → {again}행")
        check("원래 값도 같이 살아 있다", "원래 150" in txt2)

        print("[5] 되돌리기")
        pg.click("#sub-fix-reset")
        pg.wait_for_timeout(1200)
        txt3 = pg.inner_text("#sub-fix-grid")
        check("되돌리면 «직접 입력» 이 사라진다", "직접 입력" not in txt3)
        check("되돌리면 «원래» 표시도 사라진다", "원래 150" not in txt3)

        print("[6] 콘솔 오류")
        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:2]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("뽑힌 배관 손보기 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
