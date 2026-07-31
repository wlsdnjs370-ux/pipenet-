"""계통도·기계실 패널의 점진 렌더 브라우저 검증.

로컬 dev 서버에 실제 Chromium 으로 붙어 DXF 를 업로드하고, 스트리밍 도중
캔버스 픽셀이 실제로 바뀌는지(스크린샷 해시 변화 횟수)를 표본으로 확인한다.
평면도(기준)와 계통도·기계실을 같은 방식으로 재서 비교한다. 콘솔 error /
pageerror 는 전부 수집한다 — 함수-지역 헬퍼 스코프 사고(ReferenceError)는
구문검사로 안 잡히므로 이 경로로만 검출된다.

사용: python scripts/_browser_verify_stream.py --port 5061 --dxf samples/dxf/LH306동.dxf
"""
from __future__ import annotations

import argparse
import hashlib
import sys
import time

from playwright.sync_api import sync_playwright

PANELS = {
    "plane": ("#mode-btn-plane", "#wb-dxf", "#wb-overlay-info"),
    "system": ("#mode-btn-system", "#sys-dxf", "#sys-file-status"),
    "machineroom": ("#mode-btn-machineroom", "#mr-dxf", "#mr-file-status"),
}


def _probe(page, mode, dxf, timeout_s, interval_ms):
    btn_sel, input_sel, status_sel = PANELS[mode]
    page.click(btn_sel)
    page.wait_for_timeout(400)
    page.set_input_files(input_sel, dxf)

    hashes = []
    box = page.locator("canvas").first.bounding_box()
    # 전체 캔버스 스크린샷은 1장에 수 초 걸려 250ms 스로틀 리페인트를 놓친다.
    # 중앙 일부만 잘라 찍어 표본 주기를 렌더 주기보다 짧게 만든다.
    clip = {"x": box["x"] + box["width"] / 2 - 200, "y": box["y"] + box["height"] / 2 - 150,
            "width": 400, "height": 300}
    t0 = time.time()
    stable = 0
    while time.time() - t0 < timeout_s:
        try:
            shot = page.screenshot(clip=clip)
        except Exception:
            page.wait_for_timeout(interval_ms)
            continue
        h = hashlib.sha1(shot).hexdigest()
        if hashes and h == hashes[-1]:
            stable += 1
        else:
            stable = 0
        hashes.append(h)
        # 내용이 한 번이라도 그려진 뒤 5회 연속 동일 → 완료로 간주
        if stable >= 5 and len(set(hashes)) > 1:
            break
        page.wait_for_timeout(interval_ms)

    changes = sum(1 for a, b in zip(hashes, hashes[1:]) if a != b)
    try:
        status = (page.locator(status_sel).first.inner_text() or "").strip()
    except Exception:
        status = "(상태줄 없음)"
    return {
        "samples": len(hashes),
        "distinct": len(set(hashes)),
        "changes": changes,
        "elapsed_s": round(time.time() - t0, 1),
        "status": status[:200],
    }


def run(port, dxf, password, modes, timeout_s, interval_ms, headless):
    base = f"http://127.0.0.1:{port}"
    errors: list[str] = []

    with sync_playwright() as p:
        # 5060/5061 은 Chromium 이 SIP 포트로 막는다(ERR_UNSAFE_PORT) — 명시 허용.
        browser = p.chromium.launch(
            headless=headless, args=[f"--explicitly-allowed-ports={port}"])
        page = browser.new_page(viewport={"width": 1600, "height": 950})
        page.on("console",
                lambda m: errors.append(f"console.{m.type}: {m.text}") if m.type == "error" else None)
        page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))

        page.goto(f"{base}/login", wait_until="domcontentloaded")
        page.fill("input[name=password]", password)
        page.press("input[name=password]", "Enter")
        page.wait_for_load_state("domcontentloaded")

        page.goto(f"{base}/remote30-prototype", wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        out = {}
        for mode in modes:
            out[mode] = _probe(page, mode, dxf, timeout_s, interval_ms)
        page.screenshot(path="data/_browser_verify_last.png")
        browser.close()

    for mode, r in out.items():
        print(f"[{mode}] 표본 {r['samples']}회 / 서로다른화면 {r['distinct']} / "
              f"화면변화 {r['changes']}회 / {r['elapsed_s']}s")
        print(f"         상태줄: {r['status']}")
    print(f"콘솔 오류 {len(errors)}건")
    for e in errors[:20]:
        print("  !", e)
    return 1 if errors else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=5061)
    ap.add_argument("--dxf", default="samples/dxf/LH306동.dxf")
    ap.add_argument("--password", default="devlocal-throwaway-9f3a")
    ap.add_argument("--modes", default="plane,system,machineroom")
    ap.add_argument("--timeout", type=float, default=180.0)
    ap.add_argument("--interval", type=int, default=500)
    ap.add_argument("--headed", action="store_true")
    a = ap.parse_args()
    sys.exit(run(a.port, a.dxf, a.password, a.modes.split(","),
                 a.timeout, a.interval, headless=not a.headed))
