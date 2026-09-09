# -*- coding: utf-8 -*-
"""[찍기] 배관을 하나 더 찍어도 헤드가 남는가 — **화면에서** 확인.

`node --check` 는 구문만 본다. 함수-지역 헬퍼를 다른 스코프에서 부르면
구문은 멀쩡한데 브라우저에서 ReferenceError 가 난다 — 이 저장소가 겪은
회귀다. 그래서 사람이 하는 그대로 브라우저로 밟는다::

    · 도면을 열고 「자동 채택」으로 배관·헤드를 찍는다
    · 화면이 말하는 «지금 헤드 N개» 를 적는다
    · 「배관 선택」으로 돌아가 **배관을 한 줄 더** 찍는다
    · 헤드가 0 으로 떨어지지 않는지 · 콘솔 오류가 없는지

    python scripts/_verify_pick_head_survives_live.py
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

        def idle(n=12000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        print("\n■ 찍기 — 배관을 하나 더 찍어도 헤드가 남는가")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)

        n0 = pg.evaluate("() => (window.__mf.pick || {}).n_head_circles || 0")
        mat0 = pg.evaluate("() => ((window.__mf.pick || {}).materials"
                           " || []).length")
        check("도면을 열어 배관·헤드가 찍혔다", n0 > 0 and mat0 > 0,
              f"재료 {mat0}묶음 · 헤드 원 {n0}개")

        # ── 「배관 선택」으로 돌아가 **아직 안 찍은** 묶음의 선분 중점을
        #    실제로 클릭한다. ★`pick/auto` 로는 안 된다 — 이미 찍힌 묶음을
        #    건너뛰어 클릭이 **0회**가 되고, 그러면 이 검사는 아무것도
        #    증명하지 못한다(한 번 그렇게 «통과» 했다).
        rep = pg.evaluate("""async () => {
            const sid = window.__mf.sid;
            await fetch('/api/module-f/pick/mode', {method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({sid, action: 'pipe'})});
            const w = window.__mf.world || {};
            const have = new Set(((window.__mf.pick || {}).materials || [])
                                 .map((m) => m.layer + '\\u0000' + m.color));
            const b = (w.bundles || []).find(
                (x) => !have.has(x.layer + '\\u0000' + x.color)
                       && (x.segs || []).length >= 20);
            if (!b) return null;
            const sg = b.segs;
            const i = 4 * Math.floor(sg.length / 8);
            const x = (sg[i] + sg[i + 2]) / 2, y = (sg[i + 1] + sg[i + 3]) / 2;
            const r = await fetch('/api/module-f/pick/click', {method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({sid, x, y})});
            const d = await r.json();
            window.__mf.pick = d.state;
            return {want: b.layer + '×' + b.color, report: d.report};
        }""")
        if rep is None or not rep.get("report"):
            print(f"      ★새 묶음을 못 찍었습니다 — {rep}. 여기서 멈춥니다.")
            br.close()
            return 1
        idle()
        pg.wait_for_timeout(1200)
        print(f"      배관 «{rep['want']}» 을 하나 더 찍음 — {rep['report']}")
        check("★클릭이 실제로 일어났다 (0회면 아무것도 못 잰다)",
              rep["report"].get("동작") == "추가", str(rep["report"])[:120])

        n1 = pg.evaluate("() => (window.__mf.pick || {}).n_head_circles || 0")
        mat1 = pg.evaluate("() => ((window.__mf.pick || {}).materials"
                           " || []).length")
        check("★배관을 다시 찍어도 헤드가 0 이 되지 않는다", n1 > 0,
              f"헤드 원 {n0} → {n1} · 재료 {mat0} → {mat1}")
        # ★같기를 요구하면 정상을 결함이라 부르게 된다. 재료가 바뀌면 헤드
        #   문양 지문이 바뀌므로 «찍히는 원» 이 달라지는 것은 옳은 결과다.
        #   막아야 할 것은 **통째로 날아가는 것**과 «잃음» 을 잘못 세는 것이다.
        lost = int((rep["report"] or {}).get("헤드잃음") or 0)
        back = int((rep["report"] or {}).get("헤드되살림") or 0)
        check("★헤드가 통째로 날아가지 않는다 (종전엔 111 → 0)",
              n1 > 0 and back > 0, f"헤드 원 {n0} → {n1} · 되살린 칸 {back}")
        check("★잃은 칸을 부풀려 말하지 않는다", lost == 0,
              f"헤드잃음 {lost} (되살림 {back})")

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("찍은 헤드가 배관 재찍기에서 살아남는다 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
