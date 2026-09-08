# -*- coding: utf-8 -*-
"""[C-1 · B] 찍기의 «재료 묶음 빼기» 와 수리계산의 «겹침 안내» — 화면에서 확인.

지시서 `ModuleF_헤드배관_꼬임_수정지시서.md` 의 C-1·B 가 실제 화면에서 도는가.

    · 찍기 화면에 재료 묶음이 레이어별로 뜨고, 신축배관에 추천 표가 붙는가
    · 체크를 끄면 그 레이어가 빠지고 «무엇을 잃는지» 를 말하는가
    · 되돌릴 수 있는가
    · 수리계산 화면이 «등각 겹침은 위상 문제가 아니다» 를 말하는가
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
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


def main() -> int:
    from playwright.sync_api import sync_playwright

    plan = os.path.join(_D, "1. 입력도면 대명동 단위세대 평면도.dxf")
    if not os.path.isfile(plan):
        print(f"표본 없음: {plan}")
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

        def idle(n=9000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        print("[C-1] 찍기 — 재료 묶음에서 신축배관을 뺀다")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)

        pg.evaluate("() => { const h = document.querySelector("
                    "'h2.fold[data-fold=\"pk-mats-body\"]'); if (h) h.click(); }")
        pg.wait_for_timeout(600)
        rows = pg.eval_on_selector_all(
            "#pk-mats label", "es => es.map(e => e.textContent.trim())")
        print(f"      재료 묶음: {rows}")
        check("재료 묶음이 레이어별로 뜬다", len(rows) >= 3, str(len(rows)))
        check("신축배관에 추천 표가 붙는다",
              any("신축배관?" in r for r in rows),
              str([r for r in rows if "신축배관" in r]))

        n0 = pg.evaluate("() => (window.__mf.pick.materials || []).length")
        pg.evaluate("""() => {
            for (const cb of document.querySelectorAll('#pk-mats input')) {
                if (cb.parentElement.textContent.indexOf('신축배관?') >= 0) {
                    cb.checked = false; cb.onchange(); return;
                }
            }
        }""")
        idle()
        pg.wait_for_timeout(1200)
        n1 = pg.evaluate("() => (window.__mf.pick.materials || []).length")
        msg = pg.inner_text("#status")
        check("★빼면 묶음이 준다", n1 < n0, f"{n0} → {n1}")
        check("무엇을 잃는지 말한다", "헤드가 떨어졌는지" in msg, msg[:80])

        # 되돌리기
        pg.evaluate("""() => {
            for (const cb of document.querySelectorAll('#pk-mats input')) {
                if (cb.parentElement.textContent.indexOf('신축배관?') >= 0) {
                    cb.checked = true; cb.onchange(); return;
                }
            }
        }""")
        idle()
        pg.wait_for_timeout(1200)
        n2 = pg.evaluate("() => (window.__mf.pick.materials || []).length")
        check("되돌릴 수 있다", n2 == n0, f"{n1} → {n2} (원래 {n0})")

        print("[B] 수리계산 — 등각 겹침은 위상 문제가 아니다")
        pg.click("#pk-next")
        idle()
        pg.wait_for_timeout(1500)
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
        pg.click("#ed-worst")
        idle()
        pg.wait_for_timeout(1500)
        pg.evaluate("() => { for (const d of "
                    "document.querySelectorAll('#steps div'))"
                    " { if (d.textContent.indexOf('수리계산') >= 0)"
                    " { d.click(); return; } } }")
        pg.wait_for_timeout(1200)
        pg.click("#dg-build")
        idle()
        pg.wait_for_timeout(2500)
        note = pg.inner_text("#dg-iso-note")
        cx = pg.evaluate("() => (((window.__mf.design || {}).stood || {})"
                         ".crossings || null)")
        print(f"      겹침 셈: {cx}")
        check("겹침 셈이 화면까지 온다", cx is not None, str(cx))
        if cx and cx.get("total"):
            check("위상 문제가 아니라고 말한다", "위상 문제가 아닙니다" in note,
                  " ".join(note.split())[:80])
        else:
            print("      (이 도면·이 K 에서는 겹친 접속관이 0 — 안내는 안 뜬다)")

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("신축배관 빼기 · 등각 안내 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
