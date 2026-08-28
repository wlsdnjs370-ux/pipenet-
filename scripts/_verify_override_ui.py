# -*- coding: utf-8 -*-
"""[§18] 「직접 입력」을 **브라우저에서 실제로** 채워 본다.

서버 쪽은 `_probe_override_fills.py` 가 이미 쟀다. 여기서는 사람이 하는 그대로
— 목록에서 종류를 고르고 사유를 적고 저장 단추를 누르고 — 값이 산출까지 가는지
본다. 구문 검사나 DOM 존재 확인만으로는 「고른 값이 엉뚱한 자리로 간다」 같은
결함을 못 잡는다.

    python scripts/_verify_override_ui.py [--headed]
"""
from __future__ import annotations

import argparse
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
SHOTS = ROOT / "data" / "_ov_shots"
PASSWORD = "ov-verify-pw"
FAILS: list[str] = []


def check(label, cond, detail=""):
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""), flush=True)
    return cond


def free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--headed", action="store_true")
    a = ap.parse_args()
    from playwright.sync_api import sync_playwright

    if not DXF.is_file():
        print("도면 없음:", DXF)
        return 1
    SHOTS.mkdir(parents=True, exist_ok=True)
    port = free_port()
    env = dict(os.environ)
    env.update({"PORT": str(port), "HOST": "127.0.0.1",
                "LOGIN_PASSWORD": PASSWORD, "PYTHONIOENCODING": "utf-8"})
    proc = subprocess.Popen([sys.executable, "serve.py"], cwd=str(ROOT),
                            env=env, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL)
    base = f"http://127.0.0.1:{port}"
    try:
        for _ in range(90):
            try:
                socket.create_connection(("127.0.0.1", port), 0.5).close()
                break
            except OSError:
                time.sleep(1)
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=not a.headed)
            page = browser.new_page(viewport={"width": 1680, "height": 1000})
            errors: list[str] = []
            page.on("console", lambda m: errors.append(m.text)
                    if m.type == "error" else None)
            page.on("pageerror", lambda e: errors.append(str(e)))
            seen: dict = {}

            def grab(resp):
                if "/api/module-f/slot/open" in resp.url:
                    try:
                        seen["sid"] = (resp.json() or {}).get("sid")
                    except Exception:  # noqa: BLE001
                        pass
            page.on("response", grab)

            page.goto(f"{base}/module-f", wait_until="load")
            if page.query_selector("input[type=password]"):
                page.fill("input[type=password]", PASSWORD)
                with page.expect_navigation(wait_until="load"):
                    page.click("button[type=submit]")
                page.goto(f"{base}/module-f", wait_until="load")
            page.wait_for_timeout(800)

            # ── 손질까지 (기본 흐름 — 질문 0)
            page.set_input_files("#dxf", str(DXF))
            page.click("#btn-open")
            ok = False
            for _ in range(1800):
                page.wait_for_timeout(200)
                if page.is_visible("#panel-edit") and page.is_hidden("#busy"):
                    ok = True
                    break
            check("손질까지 도달", ok)
            sid = seen.get("sid")

            # ── 알람밸브 원클릭 (헤드에 가까운 배관 자리)
            st = page.evaluate(
                """async (sid) => (await (await fetch(
                    '/api/module-f/edit/state?sid=' + sid)).json()).state""", sid)
            pts = []
            for g in st["body_groups"]:
                s2 = g["segs"]
                for i in range(0, len(s2) - 3, 4):
                    pts.append((float(s2[i]), float(s2[i + 1])))
            hx, hy = float(st["heads"][0][0]), float(st["heads"][0][1])
            px, py = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            bb, box = st["bounds"], page.eval_on_selector(
                "#cv", "el => { const r = el.parentElement.getBoundingClientRect();"
                       " return {x:r.x, y:r.y, w:r.width, h:r.height}; }")
            bw = max(1e-6, bb["maxx"] - bb["minx"])
            bh = max(1e-6, bb["maxy"] - bb["miny"])
            sc = min(box["w"] / bw, box["h"] / bh) * 0.92
            ox = bb["minx"] - (box["w"] / sc - bw) / 2
            oy = bb["miny"] - (box["h"] / sc - bh) / 2
            page.mouse.click(box["x"] + (px - ox) * sc,
                             box["y"] + box["h"] - (py - oy) * sc)
            for _ in range(900):
                page.wait_for_timeout(200)
                if page.is_hidden("#busy"):
                    break

            # ── 표 확정
            page.click("#steps div:has-text('수리계산')")
            page.wait_for_selector("#panel-design:not(.hidden)", timeout=30000)
            page.click("#dg-build")
            for _ in range(1800):
                page.wait_for_timeout(200)
                if page.is_hidden("#busy"):
                    break
            page.wait_for_timeout(1200)
            chip = page.inner_text("#dg-issues-n").strip()
            check("확인할 것이 채워졌다", chip not in ("", "—"), chip)
            # 이상이 있으면 접힘이 «한 번» 저절로 펴져야 한다 — 접힌 채면
            # 「표시해서 확인하고 수정한다」가 성립하지 않는다(전사 27:41).
            opened = page.evaluate(
                """() => !document.getElementById('dg-issues-body')
                     .classList.contains('hidden')""")
            check("이상이 있으면 접힘이 저절로 펴진다", opened)
            if not opened:
                page.click('h2.fold[data-fold="dg-issues-body"]')

            # ── 채울 칸이 실제로 서 있나
            n_fields = page.eval_on_selector_all(".ovk", "els => els.length")
            check("부속 자리마다 «고르는 칸» 이 붙는다", n_fields > 0,
                  f"{n_fields}칸")
            opts = page.eval_on_selector_all(
                ".ovk:first-of-type option", "els => els.map(e => e.value)")
            check("고를 수 있는 종류를 서버가 준다",
                  "none" in opts and "elbow" in opts, " · ".join(opts))
            page.screenshot(path=str(SHOTS / "1_채우기칸.png"))

            before = page.evaluate(
                """async (sid) => {
                  const j = await (await fetch(
                    '/api/module-f/design/preview?sid=' + sid)).json();
                  const m = Object.fromEntries(j.tables.meta);
                  return {kind: m['부속 판정 불가'], ov: m['직접 입력 — 부속 판정']};
                }""", sid)

            # ── 사람이 하는 그대로 — 고르고, 사유 적고, 저장
            page.select_option(".ovk >> nth=0", "none")
            page.fill(".ovn >> nth=0", "현장 확인 — 직선이다")
            page.wait_for_timeout(200)
            check("채운 칸 수가 화면에 뜬다",
                  "1" in page.inner_text("#dg-ov-n"),
                  page.inner_text("#dg-ov-n").strip())
            page.click("#dg-ov-save")
            for _ in range(1800):
                page.wait_for_timeout(200)
                if page.is_hidden("#busy"):
                    break
            page.wait_for_timeout(1500)

            after = page.evaluate(
                """async (sid) => {
                  const j = await (await fetch(
                    '/api/module-f/design/preview?sid=' + sid)).json();
                  const m = Object.fromEntries(j.tables.meta);
                  return {kind: m['부속 판정 불가'], ov: m['직접 입력 — 부속 판정'],
                          applied: (j.tables.unresolved||{}).applied || []};
                }""", sid)
            check("★채우면 미해결이 줄고 「직접 입력」이 는다",
                  int(after["kind"]) < int(before["kind"])
                  and int(after["ov"]) > int(before["ov"]),
                  f"미해결 {before['kind']}→{after['kind']} · "
                  f"직접 입력 {before['ov']}→{after['ov']}")
            ap0 = (after["applied"] or [{}])[0]
            check("사유가 함께 남는다", "직선" in str(ap0.get("note") or ""),
                  str(ap0.get("note"))[:40])
            body = page.inner_text("#dg-issues")
            check("채운 자리가 「직접 입력」으로 목록에 남는다",
                  "직접 입력" in body, body.strip().replace("\n", " ")[:70])
            page.screenshot(path=str(SHOTS / "2_채운뒤.png"))
            check("콘솔 오류 없음", not errors,
                  " | ".join(errors[:2]) if errors else "")
            browser.close()
    finally:
        proc.terminate()

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  -", f)
        return 1
    print("PASS — 브라우저에서 직접 입력이 산출까지 간다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
