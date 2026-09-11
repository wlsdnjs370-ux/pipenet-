# -*- coding: utf-8 -*-
"""[두 화면 선정일치] 화면에서 확인한다 — 지시서 §3 «화면 확인».

  · 표를 확정하면 「평면에서 보기」가 **표에 들어간 선정**을 그린다
    (`edit.worst.from_design` · 헤드 수 == 표 노즐 수)
  · 화면이 어느 선정을 보고 있는지 **말한다** (배지 문구)
  · «고른 것 중 빠짐» 체크박스가 있고, 켜도 화면이 안 죽는다
  · 콘솔 오류 0

서버가 맞춘 것을 화면이 받아 그리는가는 파이썬 시험으로 못 말한다 — 그 사이에
`setEdit`·`renderPlanUnderlay`·`drawDesignMarks` 가 있고, 그것은 브라우저에서만
돈다(함수-지역 헬퍼를 다른 스코프에서 부르면 ReferenceError 로만 드러난다).

    python scripts/_verify_module_f_two_view_browser.py
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

        print("\n■ 두 화면 선정일치 — 브라우저")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        check("도면이 열린다", bool(pg.evaluate("() => !!window.__mf.view")))

        pg.click("#pk-next")
        idle()
        pg.wait_for_timeout(2000)
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
        picked = pg.evaluate(
            "() => ((window.__mf.edit || {}).worst || {}).heads?.length || 0")
        check("손질이 K개를 고른다", picked > 0, f"헤드 {picked}")

        pg.evaluate("() => { for (const d of "
                    "document.querySelectorAll('#steps div'))"
                    " { if (d.textContent.indexOf('수리계산') >= 0)"
                    " { d.click(); return; } } }")
        pg.wait_for_timeout(1200)
        pg.click("#dg-build")
        idle()
        pg.wait_for_timeout(3500)

        got = pg.evaluate("""() => {
            const m = window.__mf, e = m.edit || {}, d = m.design || {};
            const w = e.worst || {};
            const s = d.summary || {};
            return {shown: (w.heads || []).length,
                    from_design: !!w.from_design,
                    noz: ((s.counts || {}).nozzles) || 0,
                    k: s.k || 0,
                    filled: (s.handoff || {}).filled || 0,
                    swapped: ((s.handoff || {}).swapped_out || []).length,
                    marks: (((d.marks || {}).swapped_out) || {}).n || 0,
                    badge: (document.querySelector('#dg-edits') || {})
                             .textContent || ""};
        }""")
        check("★평면 보기가 «표에 들어간 선정» 을 든다",
              got["from_design"], str(got))
        check("★평면 헤드 수 == 표 노즐 수",
              got["shown"] and got["shown"] == got["noz"],
              f"평면 {got['shown']} · 노즐 {got['noz']}")
        # 문구는 «무엇이 같은가» 를 말한다 — 표가 있으면 배관망까지 같고
        # (net_from=design), 선정만 맞춘 상태면 「같은 선정」이다. 둘 다 인정한다.
        check("화면이 어느 망을 보는지 말한다",
              ("표와 같은 배관망" in got["badge"]
               or "표와 같은 선정" in got["badge"]), got["badge"].strip())
        if got["filled"]:
            check("교체가 있으면 목록도 온다",
                  got["swapped"] and got["marks"] == got["swapped"],
                  f"채움 {got['filled']} · 목록 {got['swapped']}"
                  f" · 도면표시 {got['marks']}")
        else:
            print(f"      (이 표본은 교체 0 — 그때는 새 문구가 안 뜨는 것이 맞다)")

        # «고른 것 중 빠짐» 을 켜도 화면이 죽지 않는다(빈 목록이어도).
        pg.evaluate("""() => { const cb = document.querySelector('#dg-mk-swap');
            cb.checked = true; cb.dispatchEvent(new Event('change')); }""")
        pg.wait_for_timeout(900)
        check("«고른 것 중 빠짐» 체크박스가 돈다",
              pg.evaluate("() => !!document.querySelector('#dg-mk-swap')"))

        # 평면 ↔ 아이소를 번갈아 — 사용자가 문제를 본 그 전환이다.
        for on in (True, False, True):
            pg.evaluate("""(want) => {
                const cb = document.querySelector('#dg-plan');
                if (cb && cb.checked !== want) { cb.checked = want;
                    cb.dispatchEvent(new Event('change')); }
            }""", on)
            pg.wait_for_timeout(1200)
        after = pg.evaluate(
            "() => ((window.__mf.edit || {}).worst || {}).heads?.length || 0")
        check("번갈아 봐도 선정이 그대로다", after == got["shown"],
              f"{got['shown']} → {after}")

        ink = pg.evaluate("""() => {
            const c = document.querySelector('#cv');
            const g = c.getContext('2d');
            const d = g.getImageData(0, 0, c.width, c.height).data;
            let n = 0;
            for (let i = 3; i < d.length; i += 4 * 37) if (d[i] > 8) n++;
            return n;
        }""")
        check("캔버스에 그림이 있다", ink > 50, f"표본 픽셀 {ink}")

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("두 화면 선정일치 — 평면 보기와 표가 같은 선정을 그린다")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
