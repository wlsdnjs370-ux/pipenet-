# -*- coding: utf-8 -*-
"""레이어 정리 화면 업로드 개조 런타임 검증 — 일회용.

구문 검사로는 스코프 밖 참조를 못 잡는다. 실제 브라우저로 DXF·DWG·대용량 도면을
올려 진행률 단계와 레이어 목록까지 확인하고, 서버가 실제로 받은 바이트 수를 재
압축이 걸렸는지 본다. 돌고 있는 서버(5051)는 건드리지 않고 별도 포트로 띄운다.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5095
SHOT = ROOT / "data" / "_upload_verify.png"

CASES = [
    ("DXF 13MB", ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"),
    ("DWG 0.3MB", ROOT / "data" / "reference_library"
     / "1. 저수조_아성다이소 양주허브센터" / "도면" / "MF-009 옥외 소화설비 평면도.dwg"),
    ("DXF 110MB", ROOT / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
]

os.environ["LOGIN_PASSWORD"] = "upverify"
sys.path.insert(0, str(ROOT))

spec = importlib.util.spec_from_file_location("daejo_server", ROOT / "대조 서버.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

# 서버가 실제로 받은 요청 바이트 — 압축이 걸렸는지는 화면 문구가 아니라 이 값이 증거다.
received: dict[str, int] = {}


@mod.app.before_request
def _measure(*_a, **_k):
    from flask import request
    if request.path.startswith("/api/remote30/"):
        received[request.path] = request.content_length or 0


threading.Thread(
    target=lambda: mod.app.run(host="127.0.0.1", port=PORT, threaded=True),
    daemon=True,
).start()

from playwright.sync_api import sync_playwright  # noqa: E402

errors: list[str] = []
phases: list[str] = []

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1600, "height": 950})
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))

    base = f"http://127.0.0.1:{PORT}"
    page.goto(f"{base}/login", wait_until="domcontentloaded")
    page.fill("input[type=password]", "upverify")
    page.press("input[type=password]", "Enter")
    page.wait_for_load_state("networkidle")

    for label, path in CASES:
        if not path.is_file():
            print(f"[{label}] 파일 없음 — 건너뜀: {path}")
            continue
        page.goto(f"{base}/remote30-workbench", wait_until="networkidle")
        page.wait_for_selector("#wb-load-status", timeout=15_000)

        # 진행률 문구는 순식간에 지나간다 — 상태줄 변화를 MutationObserver 로 모은다.
        page.evaluate("""() => {
            window.__phases = [];
            const el = document.getElementById('wb-load-status');
            new MutationObserver(() => {
                const t = el.textContent;
                if (!window.__phases.length || window.__phases.at(-1) !== t) {
                    window.__phases.push(t);
                }
            }).observe(el, {childList: true, characterData: true, subtree: true});
        }""")

        received.clear()
        t0 = time.time()
        page.set_input_files("#wb-dxf", str(path))
        try:
            page.wait_for_selector("#wb-layer-list .layer-row", timeout=600_000)
        except Exception as exc:
            print(f"[{label}] 실패: {exc}")
            print(f"[{label}] 상태줄: {page.inner_text('#wb-load-status')}")
            continue
        elapsed = time.time() - t0

        seen = page.evaluate("() => window.__phases")
        phases.append(f"[{label}] " + " → ".join(
            s for s in seen if "%" in s or "분석" in s or "로딩" in s)[:400])

        src_mb = path.stat().st_size / 1048576
        up_mb = received.get("/api/remote30/upload", 0) / 1048576
        print(f"[{label}] {elapsed:5.1f}s  원본 {src_mb:6.1f} MB "
              f"→ 전송 {up_mb:6.1f} MB "
              f"({src_mb / up_mb:.1f}배 절감)" if up_mb else f"[{label}] 전송량 미측정")
        print(f"[{label}] 상태줄: {page.inner_text('#wb-load-status')}")
        print(f"[{label}] 레이어 행: {page.locator('#wb-layer-list .layer-row').count()}"
              f" / 재출력 버튼 활성: {page.is_enabled('#wb-export-btn')}")
        page.screenshot(path=str(SHOT.with_name(
            f"_upload_verify_{label.split()[0].lower()}_{int(src_mb)}.png")))

    browser.close()

print("\n── 진행 단계 ──")
for line in phases:
    print(line)
print("\n── 콘솔 오류 ──")
print("\n".join(errors) if errors else "없음")
