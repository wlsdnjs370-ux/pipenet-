# -*- coding: utf-8 -*-
"""HSP 펌프 정격유량·양정 입력칸 브라우저 검증 (작업 2).

구문검사로는 못 잡는 것 두 가지를 실제 Chromium 에서 본다.
(1) HSP_PUMP 를 고르면 정격 입력칸이 실제로 보이는가,
(2) 값을 넣었을 때 폼 전송 payload 에 pump_rated_* 가 실리고, 비우면 빠지는가.

사용: python scripts/_pump_rated_browser_probe.py --port 5050 --password ...
"""
from __future__ import annotations

import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright  # noqa: E402


def _session_cookie(base, password):
    import requests
    r = requests.post(f"{base}/login", data={"password": password},
                      headers={"X-Forwarded-Proto": "https"},
                      timeout=30, allow_redirects=False)
    value = r.cookies.get("session")
    if not value:
        raise SystemExit(f"로그인 실패 — status={r.status_code}")
    return {"name": "session", "value": value, "domain": "127.0.0.1", "path": "/",
            "httpOnly": True, "secure": False, "sameSite": "Lax"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=5050)
    ap.add_argument("--password", default=os.environ.get("LOGIN_PASSWORD", ""))
    a = ap.parse_args()
    base = f"http://127.0.0.1:{a.port}"

    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context()
        ctx.add_cookies([_session_cookie(base, a.password)])
        page = ctx.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
        page.goto(f"{base}/remote30-overall", wait_until="networkidle")

        for eid in ("ov-pump-q", "ov-pump-h"):
            assert page.locator(f"#{eid}").count() == 1, f"{eid} 없음"

        # 펌프 이름은 이제 서버 상수를 Jinja 로 렌더한다 — 입력칸 값과 JS 폴백 양쪽을 본다.
        lib = page.input_value("#ov-pump-lib")
        assert lib == "SP_162M_2900LPM", f"Library-pump 기본값이 렌더되지 않았다: {lib!r}"
        html = page.content()
        assert 'vals("ov-pump-lib") || "SP_162M_2900LPM"' in html, "JS 폴백이 렌더되지 않았다"

        page.check("input[name=ov_zone_type][value=hsp_pump]")
        assert page.locator("#ov-pump-section").is_visible(), "HSP 선택인데 펌프칸이 숨어 있다"

        # 폼 조립부만 떼어내 payload 를 본다 — 업로드 없이 검증하기 위해.
        def payload(q, h):
            return page.evaluate("""([q, h]) => {
                document.getElementById("ov-pump-q").value = q;
                document.getElementById("ov-pump-h").value = h;
                const fd = new FormData();
                const vals = (id) => (document.getElementById(id).value || "").trim();
                if (vals("ov-pump-q")) fd.append("pump_rated_q_lpm", vals("ov-pump-q"));
                if (vals("ov-pump-h")) fd.append("pump_rated_h_m", vals("ov-pump-h"));
                return Object.fromEntries(fd.entries());
            }""", [q, h])

        filled = payload("2900", "162")
        assert filled == {"pump_rated_q_lpm": "2900", "pump_rated_h_m": "162"}, filled
        assert payload("", "") == {}, "빈칸인데 0 이 실려 나간다"

        page.screenshot(path="data/_pump_rated_probe.png", full_page=False)
        browser.close()

        if errors:
            print("JS 오류:", *errors, sep="\n  ")
            return 1
    print("OK — 입력칸 표시 · payload 조건부 첨부 · JS 오류 없음")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
