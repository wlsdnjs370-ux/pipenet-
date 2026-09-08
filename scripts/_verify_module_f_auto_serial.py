# -*- coding: utf-8 -*-
"""[회로] 자동(A)은 지선이고, 합류는 «손질로 이어받기» 하나 — 화면 확인.

사용자 지적: 「지금 로직 직렬 방식 아니야? 일단 자동으로 먼저 처리하고 그다음
수동으로 하는 걸로 아는데. 병렬 방식은 내 지향점이 아닌데.」

맞다. 기본 길은 직렬이다(올리면 정찰·자동채택이 먼저 돌고 사람이 찍기에서
고친다). 자동(A)은 **지선**이다 — 접힌 «고급» 안에서 들어가고(D-F10-2), 끝내려면
「손질로 이어받기」로 본선에 합류한다. 한때 이 지선의 단계바를 본선과 같은
다섯 칸으로 폈다가 되돌렸다(사용자 반려: 「두 길의 화면 차례를 같게 만들면
지선이 본선과 대등해 보인다」).

여기서 보는 것: 지선 단계바가 **짧은가**(본선과 대등해 보이지 않는가), 나가는
문이 「손질로 이어받기」 하나인가, 파일 저장 단추가 없는가, 그리고 화면이
«파일을 내려면 이어받으라» 고 알려 주는가.
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

        # ── 자동(A) 화면 — 지선으로 서 있는가
        print("[②] 자동(A) 화면 — 짧은 지선 · 나가는 문 하나")
        pg.evaluate("() => { const h = document.querySelector("
                    "'h2.fold[data-fold=\"adv-body\"]'); if (h) h.click(); }")
        pg.wait_for_timeout(300)
        pg.click("#adv-auto")
        idle()
        pg.wait_for_timeout(1500)
        aut = steps()
        print(f"      자동 단계바: {aut}")
        check("지선 단계바가 짧다 (본선과 대등해 보이지 않는다)",
              len(aut) < len(man) and aut == ["도면 열기", "자동 추출"], str(aut))
        check("나가는 문이 «손질로 이어받기» 하나다",
              "이어받기" in pg.inner_text("#au-handoff")
              and pg.locator("#au-to-design").count() == 0,
              pg.inner_text("#au-handoff"))
        check("자동 화면에 파일 저장 단추가 없다",
              pg.locator("#dg-emit").count() == 0
              and pg.locator("#dg-download").count() == 0)
        panel = pg.inner_text("#panel-auto")
        check("파일을 내려면 무엇을 하라고 알려 준다",
              "파일을 내려면" in panel and "이어받기" in panel,
              " ".join(panel.split())[-90:])

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("자동은 지선 · 합류는 이어받기 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
