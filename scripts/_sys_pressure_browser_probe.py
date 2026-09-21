# -*- coding: utf-8 -*-
"""계통도 압력표 입력 UI 브라우저 검증.

구문검사로는 스코프 사고를 못 잡는다. 실제 Chromium 에서 (1) 압력표 칸이 있고
(2) 붙여넣기 상태줄이 실제로 갱신되며 (3) 결과 카드의 표고 출처 블록이 예외
없이 렌더되는지를 본다.

사용: python scripts/_sys_pressure_browser_probe.py --port 5053 --password ...
"""
from __future__ import annotations

import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright  # noqa: E402

_BASE_RISER = {
    "title": "라이저 (probe)",
    "nodes": [{"label": "1", "elevation": 5.8, "io_node": "Input"},
              {"label": "10", "elevation": 0.0}],
    "pipes": [{"label": "P1"}],
    "total_pipe_length_m": 30.0,
    "floor_matching": {"label_count": 3, "floor_height_mm": 2900.0,
                       "av_floor_idx": 16, "av_floor_name": "16층",
                       "nodes_with_floor": 2},
}
RISER_USER = dict(_BASE_RISER, elev_sources={"nodes": {"user_confirmed": 2},
                                             "pipes": {"user_confirmed": 1}})
RISER_EST = dict(_BASE_RISER, elev_sources={"nodes": {"drawing_estimated": 2},
                                            "pipes": {"drawing_estimated": 1}})

PT_GOOD = '[{"floor_label":"18층","head_drop_m":13.8},{"floor_label":"17층","head_drop_m":16.7}]'
PT_PARTIAL = '[{"floor_label":"18층","head_drop_m":13.8},{"floor_label":"17층"}]'
PT_BROKEN = '[{"floor_label":'
PT_NOT_LIST = '{"floor_label":"18층"}'


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
    ap.add_argument("--port", type=int, default=5053)
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
        if page.locator("#mode-btn-system").count() == 0:
            raise SystemExit(f"프로토타입 진입 실패 — url={page.url}")

        page.click("#mode-btn-system")
        page.wait_for_timeout(400)

        has = page.locator("#sys-pressure-json").count() == 1
        print(f"압력표 입력칸 존재={has} 보임={page.locator('#sys-pressure-json').is_visible() if has else False}")

        for label, raw in (("정상", PT_GOOD), ("일부결측", PT_PARTIAL),
                           ("깨진JSON", PT_BROKEN), ("배열아님", PT_NOT_LIST), ("비움", "")):
            page.fill("#sys-pressure-json", raw)
            page.wait_for_timeout(120)
            print(f"[{label}] 상태줄 = {page.inner_text('#sys-pressure-status').strip()[:110]}")

        # 결과 카드 표고 출처 블록 — 스코프 밖이면 여기서 잡힌다.
        for label, riser in (("압력표 사용", RISER_USER), ("도면 추정", RISER_EST)):
            out = page.evaluate(
                """(riser) => {
                    if (typeof _updateSystemResultCard !== 'function') return {err: 'not in scope'};
                    try { _updateSystemResultCard({riser, algorithm: 'v1'}); }
                    catch (e) { return {err: String(e)}; }
                    const el = document.getElementById('sys-result-box');
                    return {txt: el ? el.innerText : '(카드 요소 못찾음)'};
                }""", riser)
            if out.get("err"):
                print(f"[{label}] 렌더 실패 — {out['err']}")
            else:
                line = next((l for l in out["txt"].splitlines() if "확정" in l or "추정" in l), "(표고 줄 없음)")
                print(f"[{label}] 렌더 OK — {line.strip()[:110]}")

        page.screenshot(path="data/_sys_pressure_probe.png")
        browser.close()

    print(f"콘솔/페이지 오류 {len(errors)}건")
    for e in errors[:10]:
        print("  ", e[:200])


if __name__ == "__main__":
    main()
