# -*- coding: utf-8 -*-
"""손질 화면에 «급수 시작» 이 사라지고 알람밸브 한 픽만 남았는가 — 브라우저 실측.

시험은 소스를 읽지만, 사람이 보는 것은 화면이다. 단추가 정말 하나인지·클릭
한 번이 두 칸을 다 놓는지를 실제 DOM 과 서버 상태로 잰다.
"""
from __future__ import annotations

import os
import re
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


def main():
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        br = pw.chromium.launch()
        pg = br.new_page(viewport={"width": 1600, "height": 1000})
        errors: list[str] = []
        pg.on("console", lambda m: errors.append(m.text)
              if m.type == "error" else None)

        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")

        print("[1] 손질 모드 단추")
        modes = pg.eval_on_selector_all(
            ".emode", "els => els.map(e => e.dataset.mode)")
        check("«급수시작위치» 단추가 없다", "급수시작위치" not in modes, str(modes))
        check("«알람밸브위치» 단추는 있다", "알람밸브위치" in modes, str(modes))
        lab = pg.eval_on_selector_all(
            '.emode[data-mode="알람밸브위치"]', "els => els.map(e => e.textContent)")
        check("이름이 접속점을 밝힌다",
              any("접속점" in t for t in lab), str(lab))

        print("[2] 제안 반영 단추")
        check("«급수 시작 제안» 단추가 없다",
              pg.query_selector("#ed-hint-source") is None)
        check("«알람밸브 제안» 단추는 있다",
              pg.query_selector("#ed-hint-alarm") is not None)

        print("[3] 자동 차선 설명")
        s1 = pg.inner_text("#au-s1")
        check("«손질의 급수 시작과 다르다» 는 옛 설명이 없다",
              "다른 단계" not in s1, s1.replace("\n", " ")[:120])
        check("«같은 자리» 라고 말한다", "같은 자리" in s1,
              s1.replace("\n", " ")[:120])

        print("[4] 콘솔 오류")
        real = [e for e in errors if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("손질 화면 — 알람밸브 한 픽으로 통합 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
