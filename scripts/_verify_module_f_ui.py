# -*- coding: utf-8 -*-
"""module_f 화면을 **실제 브라우저에서** 연다 — 단계바 리팩터 회귀 방벽.

`node --check` 도 스코프 검사도 «파일이 성립하는가» 까지다. 페이지가 뜰 때
정말 아무 것도 안 터지는지, 단계바가 슬롯에 맞게 그려지는지는 브라우저에서만
드러난다(실측 회귀: 함수-지역 헬퍼를 다른 스코프에서 불러 클릭하는 순간
ReferenceError 로 죽은 적이 있다).

운영 서버(5051)를 쓰지 않는다 — 로그인 실패가 쌓이면 IP 가 잠긴다. 여기서는
**임시 포트에 제 서버를 띄워** 그것만 본다.

    python scripts/_verify_module_f_ui.py
"""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# ASCII 로 둔다 — Windows 서브프로세스 env 는 비ASCII 를 코드페이지로 넘겨
# 서버가 받은 값과 여기서 타이핑한 값이 달라진다(실측으로 로그인이 실패했다).
PASSWORD = "ui-verify-pw-2026"
FAILS: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def _free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    from playwright.sync_api import sync_playwright

    port = _free_port()
    env = dict(os.environ)
    env.update({"PORT": str(port), "HOST": "127.0.0.1",
                "LOGIN_PASSWORD": PASSWORD,
                "DESIGN_WORKBENCH_ENABLED": "1",
                "PYTHONIOENCODING": "utf-8"})
    print(f"임시 서버 :{port} 기동 중…")
    proc = subprocess.Popen([sys.executable, "serve.py"], cwd=str(ROOT),
                            env=env, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL)
    base = f"http://127.0.0.1:{port}"
    try:
        for _ in range(60):
            try:
                socket.create_connection(("127.0.0.1", port), 0.5).close()
                break
            except OSError:
                time.sleep(1)
        else:
            check("임시 서버 기동", False, "포트가 안 열림")
            return 1
        check("임시 서버 기동", True, f":{port}")

        errors: list[str] = []
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.on("console", lambda m: errors.append(m.text)
                    if m.type == "error" else None)
            page.on("pageerror", lambda e: errors.append(str(e)))

            page.goto(f"{base}/module-f", wait_until="load")
            if page.query_selector("input[type=password]"):
                page.fill("input[type=password]", PASSWORD)
                with page.expect_navigation(wait_until="load"):
                    page.click("button[type=submit]")
                page.goto(f"{base}/module-f", wait_until="load")
            if "login" in page.url:
                # 왜 막혔는지 화면 문구를 그대로 싣는다 — 비번인지 잠금인지.
                msg = (page.inner_text("body") or "")[:160].replace("\n", " ")
                check("로그인 통과", False, msg)
                browser.close()
                return 1
            check("로그인 통과", True, page.url)

            page.wait_for_timeout(900)

            # ── 단계바가 슬롯 흐름대로 그려졌나 (평면도가 기본)
            steps = page.eval_on_selector_all(
                "#steps div", "els => els.map(e => e.textContent.trim())")
            want = ["도면 열기", "찍기", "손질", "변환", "수리계산", "통합"]
            check("단계바 = 평면도 흐름", steps == want, " · ".join(steps))

            # ── 슬롯 탭이 그려졌나 (세션 전에는 안내문)
            slots = page.inner_text("#slots").strip()
            check("슬롯 자리 있음", bool(slots), slots[:50])

            # ── 새로 만든 패널들이 DOM 에 있나
            for pid in ("panel-sub", "panel-merge", "panel-design",
                        "ed-zone-arm", "ed-zones", "dg-bore-legend"):
                check(f"#{pid} 존재", page.query_selector(f"#{pid}") is not None)

            # ── 단계바 클릭이 터지지 않나 (갈 수 없는 곳은 막혀야 한다)
            page.click("#steps div:nth-child(2)")     # 찍기 — 재료 없음
            page.wait_for_timeout(300)
            after = page.eval_on_selector_all(
                "#steps div.on", "els => els.map(e => e.textContent.trim())")
            check("재료 없는 단계로는 안 넘어간다", after == ["도면 열기"],
                  " · ".join(after))

            # ── 콘솔 오류 0
            check("콘솔 오류 없음", not errors,
                  " | ".join(errors[:3]) if errors else "")

            page.screenshot(path=str(ROOT / "data" / "_ui_verify.png"),
                            full_page=False)
            browser.close()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 화면이 브라우저에서 정상으로 뜬다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
