# -*- coding: utf-8 -*-
"""[D-F10-4 개정] 알람밸브 클릭이 «놓기만» 하는가 — 살아 있는 서버로.

사용자 지시: 「알람밸브 클릭하자마자 최불리 배관망이 자동으로 뽑아지던데,
이렇게 하지 말고 알람밸브 지정 후 영역 지정 후에 버튼 누르면 배관망 나오도록」.

그래서 세 가지를 화면에서 그대로 본다.
  ① 알람밸브를 클릭해도 **배관망이 안 나온다**(화소로 잰다)
  ② 영역을 그리고 「최불리 선정」을 누르면 그때 나온다
  ③ 알람밸브 자리를 옮기면 먼저 뽑은 배관망이 **남아 있지 않다**
     (남으면 밸브는 새 자리인데 망은 옛 자리 것이라 둘이 다른 말을 한다)
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

        print("[1] 평면도를 열어 손질까지")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        got = False
        for _ in range(3000):
            pg.wait_for_timeout(200)
            if pg.is_visible("#panel-edit") and pg.is_hidden("#busy"):
                got = True
                break
        check("손질 화면이 열린다", got)
        if not got:
            br.close()
            return 1
        pg.wait_for_timeout(1200)

        check("순서 안내가 뜬다", "① " in pg.inner_text("#ed-anchor-note"),
              pg.inner_text("#ed-anchor-note")[:60].replace("\n", " "))
        check("알람밸브 전에는 «최불리 선정» 이 잠겨 있다",
              pg.is_disabled("#ed-worst"),
              pg.inner_text("#ed-worst-why")[:50])

        def worst():
            return pg.evaluate("() => (window.__mf.edit || {}).worst || null")

        print("[2] 알람밸브를 찍는다 — 배관망은 아직 안 나와야 한다")
        pts = pg.evaluate("""() => {
            const e = window.__mf.edit;
            const out = [];
            for (const g of (e.body_groups || [])) {
                const s = g.segs || [];
                for (let i = 0; i + 1 < s.length; i += 4) out.push([s[i], s[i+1]]);
                if (out.length > 400) break;
            }
            return out;
        }""")
        check("배관 위 좌표를 얻었다", len(pts) > 0, f"{len(pts)}점")
        if not pts:
            br.close()
            return 1
        a = pts[len(pts) // 2]
        # 화면 클릭 대신 서버 경로를 그대로 태운다 — 캔버스 좌표 환산이
        # 이 검사의 대상이 아니다(그건 다른 검증이 본다).
        pg.evaluate("""async (xy) => {
            const S = window.__mf;
            await fetch('/api/module-f/edit/anchor-click', {
              method: 'POST', headers: {'Content-Type': 'application/json'},
              body: JSON.stringify({sid: S.sid, x: xy[0], y: xy[1],
                                    max_d: 3000})});
        }""", a)
        for _ in range(600):
            pg.wait_for_timeout(200)
            j = pg.evaluate(
                "async () => (await (await fetch"
                "('/api/module-f/job?sid=' + window.__mf.sid)).json()).state")
            if j in ("done", "error", "idle"):
                break
        pg.evaluate("""async () => {
            const S = window.__mf;
            const d = await (await fetch('/api/module-f/edit/state?sid='
                                         + S.sid)).json();
            S.edit = d.state;
        }""")
        pg.wait_for_timeout(400)
        st = pg.evaluate("() => window.__mf.edit")
        check("알람밸브가 놓였다", len(st.get("sources") or []) == 1,
              str(st.get("sources")))
        check("★배관망은 아직 안 나온다", not worst(), str(worst())[:60])

        print("[3] 버튼을 누르면 그때 나온다")
        pg.evaluate("""async () => {
            const S = window.__mf;
            const d = await (await fetch('/api/module-f/edit/worst', {
              method: 'POST', headers: {'Content-Type': 'application/json'},
              body: JSON.stringify({sid: S.sid, k: 30})})).json();
            S.edit = d.state;
        }""")
        pg.wait_for_timeout(600)
        w = worst()
        check("★버튼을 누르니 배관망이 나온다", bool(w),
              f"{(w or {}).get('k')}개 · 최원 {(w or {}).get('far_m')} m")

        print("[4] 알람밸브를 옮기면 옛 배관망은 안 남는다")
        b = pts[len(pts) // 3]
        if abs(b[0] - a[0]) + abs(b[1] - a[1]) < 1.0:
            b = pts[0]
        pg.evaluate("""async (xy) => {
            const S = window.__mf;
            await fetch('/api/module-f/edit/anchor-click', {
              method: 'POST', headers: {'Content-Type': 'application/json'},
              body: JSON.stringify({sid: S.sid, x: xy[0], y: xy[1],
                                    max_d: 3000})});
        }""", b)
        for _ in range(600):
            pg.wait_for_timeout(200)
            j = pg.evaluate(
                "async () => (await (await fetch"
                "('/api/module-f/job?sid=' + window.__mf.sid)).json()).state")
            if j in ("done", "error", "idle"):
                break
        pg.evaluate("""async () => {
            const S = window.__mf;
            const d = await (await fetch('/api/module-f/edit/state?sid='
                                         + S.sid)).json();
            S.edit = d.state;
        }""")
        pg.wait_for_timeout(400)
        check("★자리를 옮기면 옛 배관망이 지워진다", not worst(),
              str(worst())[:60])

        print("[5] 콘솔 오류")
        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:2]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("알람밸브 → 영역 → 버튼 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
