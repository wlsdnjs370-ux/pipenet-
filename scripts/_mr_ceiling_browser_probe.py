# -*- coding: utf-8 -*-
"""기계실 천장고 UI 브라우저 검증.

구문검사로는 안 잡히는 스코프 사고를 잡으려면 실제 Chromium 에서 그 코드가
돌아야 한다. 여기서는 천장고 입력칸이 실제로 존재하는지, 결과 카드의 표고
분기(천장고 있음 / 없음)가 예외 없이 렌더되는지를 본다.

사용: python scripts/_mr_ceiling_browser_probe.py --port 5051 --password ...
"""
from __future__ import annotations

import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright  # noqa: E402

MR_WITH = {
    "nodes": [{"label": "m1"}, {"label": "m2"}],
    "pipes": [{"label": "M1"}],
    "total_pipe_length_m": 12.3,
    "machine_room_ceiling_m": 4.5,
    "elev_sources": {"pipes": {"user_confirmed": 1}},
}
MR_WITHOUT = {
    "nodes": [{"label": "m1"}, {"label": "m2"}],
    "pipes": [{"label": "M1"}],
    "total_pipe_length_m": 12.3,
    "machine_room_ceiling_m": None,
    "elev_sources": {"pipes": {"unresolved": 1}},
}


def _session_cookie(base, password):
    import requests
    r = requests.post(f"{base}/login", data={"password": password},
                      headers={"X-Forwarded-Proto": "https"},
                      timeout=30, allow_redirects=False)
    value = r.cookies.get("session")
    if not value:
        raise SystemExit(f"로그인 실패 — status={r.status_code} (비번 확인)")
    return {"name": "session", "value": value, "domain": "127.0.0.1", "path": "/",
            "httpOnly": True, "secure": False, "sameSite": "Lax"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=5051)
    ap.add_argument("--password", default=os.environ.get("LOGIN_PASSWORD", ""))
    ap.add_argument("--headed", action="store_true")
    a = ap.parse_args()

    base = f"http://127.0.0.1:{a.port}"
    errors: list[str] = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not a.headed,
                                    args=[f"--explicitly-allowed-ports={a.port}"])
        page = browser.new_page(viewport={"width": 1600, "height": 950})
        page.context.add_cookies([_session_cookie(base, a.password)])
        page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
                if m.type == "error" else None)
        page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))

        page.goto(f"{base}/remote30-prototype", wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        if page.locator("#mode-btn-machineroom").count() == 0:
            raise SystemExit(f"프로토타입 진입 실패 — url={page.url}")

        page.click("#mode-btn-machineroom")
        page.wait_for_timeout(500)

        has_input = page.locator("#mr-ceiling-h").count() == 1
        visible = page.locator("#mr-ceiling-h").is_visible() if has_input else False
        print(f"천장고 입력칸 존재={has_input} 보임={visible}")
        if has_input:
            page.fill("#mr-ceiling-h", "4.5")
            print(f"입력값 왕복 = {page.input_value('#mr-ceiling-h')!r}")

        # 결과 카드 두 갈래를 실제 함수로 렌더 — 스코프 밖이면 여기서 잡힌다.
        for label, mr in (("천장고 있음", MR_WITH), ("천장고 없음", MR_WITHOUT)):
            out = page.evaluate(
                """(mr) => {
                    const fn = window._updateMachineroomResultCard;
                    if (typeof fn !== 'function') return {err: 'not in global scope'};
                    try { fn(mr); } catch (e) { return {err: String(e)}; }
                    const el = document.getElementById('mr-result-box');
                    return {html: el ? el.innerHTML : '(카드 요소 못찾음)'};
                }""", mr)
            if out.get("err"):
                print(f"[{label}] 렌더 실패 — {out['err']}")
            else:
                txt = out["html"]
                mark = "표고 미확정" if "표고 미확정" in txt else ("표고" if "표고" in txt else "없음")
                print(f"[{label}] 렌더 OK — 표고 블록: {mark}")

        page.screenshot(path="data/_mr_ceiling_probe.png")
        browser.close()

    print(f"콘솔/페이지 오류 {len(errors)}건")
    for e in errors[:10]:
        print("  ", e[:200])


if __name__ == "__main__":
    main()
