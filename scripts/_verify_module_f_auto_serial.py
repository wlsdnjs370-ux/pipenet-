# -*- coding: utf-8 -*-
"""[회로] 자동(A)도 **같은 회로**를 타는가 — 화면에서 확인.

사용자 지적: 「지금 로직 직렬 방식 아니야? 일단 자동으로 먼저 처리하고 그다음
수동으로 하는 걸로 아는데. 병렬 방식은 내 지향점이 아닌데.」

맞다. 기본 길은 직렬이다(올리면 정찰·자동채택이 먼저 돌고 사람이 찍기에서
고친다). 어긋나 있던 것은 «자동 추출(A)» 화면 하나였다 — 그 길만 단계바가
[열기 → 자동 추출 → 수리계산] 으로 갈라져 손질을 건너뛰었다.

여기서 보는 것: 자동 화면의 단계바가 수동과 **같은 꼬리**를 갖는가, 다음
걸음이 「손질로 이어받기」인가, 그리고 파일 저장 단추가 이 화면에 없는가.
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

        def steps():
            return pg.eval_on_selector_all(
                "#steps div", "es => es.map(e => e.textContent)")

        # ── 기본 길(직렬) — 올리면 자동 채택이 먼저 돌고 «찍기» 에서 멈춘다
        print("[①] 기본 길 — 자동이 먼저, 사람이 그 위에서")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        st = pg.evaluate("() => window.__mf.stage")
        check("올리면 «찍기» 에서 멈춘다", st == "pick", str(st))
        picked = pg.inner_text("#pk-count")
        check("자동이 이미 찍어 두었다", "0개" not in picked,
              " ".join(picked.split())[:60])
        man = steps()
        print(f"      수동 단계바: {man}")

        # ── 자동(A) 화면 — 같은 꼬리를 갖는가
        print("[②] 자동(A) 화면 — 갈라진 차선이 아니어야 한다")
        pg.evaluate("() => { const h = document.querySelector("
                    "'h2.fold[data-fold=\"adv-body\"]'); if (h) h.click(); }")
        pg.wait_for_timeout(300)
        pg.click("#adv-auto")
        idle()
        pg.wait_for_timeout(1500)
        aut = steps()
        print(f"      자동 단계바: {aut}")
        check("자동 단계바에 «손질» 이 있다", "손질" in aut, str(aut))
        check("꼬리가 수동과 같다", aut[-3:] == man[-3:], f"{aut[-3:]} vs {man[-3:]}")
        check("다음 걸음이 «손질로 이어받기»",
              "이어받기" in pg.inner_text("#au-handoff"),
              pg.inner_text("#au-handoff"))
        check("자동 화면에 파일 저장 단추가 없다",
              pg.locator("#dg-emit").count() == 0)

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("자동도 같은 회로 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
