# -*- coding: utf-8 -*-
"""[마무리 확인] 실서버에서 한 바퀴 — 열기 → 찍기 → 손질 → 수리계산 → 아이소.

이번에 손댄 것이 **찍기 → 손질 경로**에 들어갔으므로(같은 자리 점 합치기 ·
챔퍼 모서리 복원 · 헤드 상하향 · 신축배관 접기), 시험만으로는 「화면에서도
도는가」를 못 말한다. 띄워 놓은 서버로 실제 한 바퀴를 돈다.

    · 도면을 열고 찍기 → 손질까지 간다 (콘솔 오류 0)
    · 수리계산에서 표를 확정한다
    · **아이소를 켜고 끈다** — 사용자가 문제를 본 그 자리
    · 헤드가 전부 세로로 서는지 · 등각 겹침이 «위상 문제가 아니다» 라고
      말하는지 · 표 값이 화면에 오는지

    python scripts/_verify_module_f_corner_live.py
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_D = os.path.join(_ROOT, "routes", "제출용[최종]")
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}"
          + (f" · {detail}" if detail else ""))
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

        print("\n■ 실서버 한 바퀴")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        check("도면이 열린다", bool(pg.evaluate("() => !!window.__mf.view")))

        pg.click("#pk-next")
        idle()
        pg.wait_for_timeout(2000)
        # ★`edit` 상태에는 `pts` 가 없다 — 화면이 받는 것은 `body_groups` 다.
        #   («pts 0» 을 결함으로 읽어 한 번 헛짚었다.)
        n_seg = pg.evaluate("""() => (window.__mf.edit.body_groups || [])
                                .reduce((a, g) => a + ((g.segs || []).length / 4), 0)""")
        n_grp = pg.evaluate("() => (window.__mf.edit.body_groups || []).length")
        check("손질망이 선다", n_seg > 0, f"몸통 묶음 {n_grp} · 선분 {n_seg:.0f}")

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
        pg.wait_for_timeout(3000)

        got = pg.evaluate("""() => {
            const d = window.__mf.design || {};
            const v = d.view || {};
            return {nodes: (v.nodes || []).length,
                    pipes: (v.pipes || []).length,
                    stood: (d.stood || {}).vertical,
                    heads: (d.stood || {}).heads,
                    cross: ((d.stood || {}).crossings || {}).total};
        }""")
        check("표가 확정된다", (got.get("nodes") or 0) > 0, str(got))

        # ★사용자가 문제를 본 자리 — 아이소를 켜고 끈다.
        for on in (True, False, True):
            pg.evaluate("""(want) => {
                const cb = document.querySelector('#dg-iso');
                if (cb && cb.checked !== want) { cb.checked = want;
                    cb.dispatchEvent(new Event('change')); }
            }""", on)
            idle()
            pg.wait_for_timeout(1800)
        got2 = pg.evaluate("""() => {
            const d = window.__mf.design || {};
            return {stood: (d.stood || {}).vertical,
                    heads: (d.stood || {}).heads,
                    loose: (d.stood || {}).loose,
                    cross: ((d.stood || {}).crossings || {}).total};
        }""")
        check("★아이소를 켜면 헤드가 전부 세로로 선다",
              got2.get("heads") and got2.get("stood") == got2.get("heads"),
              str(got2))
        check("★관말이 아니라 못 세운 헤드 0", not got2.get("loose"),
              str(got2.get("loose")))
        note = pg.inner_text("#dg-iso-note")
        if got2.get("cross"):
            check("겹침이 있으면 «위상 문제가 아니다» 라고 말한다",
                  "위상 문제가 아닙니다" in note, " ".join(note.split())[:80])
        else:
            print("      (등각 겹침 0 — 안내는 안 뜬다)")

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("실서버 한 바퀴 — 열기·손질·표 확정·아이소 전환 모두 정상")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
