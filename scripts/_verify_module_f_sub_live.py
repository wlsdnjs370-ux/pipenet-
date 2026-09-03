# -*- coding: utf-8 -*-
"""[H-2 · H-3] 계통도 — 두 점을 찍는 동안 «선이 따라오나».

구문 검사로는 «따라온다» 를 말할 수 없다. 실제 브라우저에서 캔버스 화소를
세어 확인한다:

  ① 첫 점을 찍기 전에는 미리보기 선이 없다
  ② 첫 점을 찍고 마우스를 옮기면 그 색 선이 생긴다
  ③ 마우스를 다른 데로 옮기면 선이 **달라진다**(따라온다)
  ④ 레이어를 고르면 경로 그래프가 실제로 줄어든다
  ⑤ 미리보기와 추출이 같은 그래프를 쓴다 — 뽑은 경로가 미리보기와 어긋나지 않는다

실행:
    MF_BASE=http://127.0.0.1:5065 LOGIN_PASSWORD=… \
        python scripts/_verify_module_f_sub_live.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
SHOT = ROOT / "data" / "_mf_shots"
SHOT.mkdir(parents=True, exist_ok=True)
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5065")
PASSWORD = os.environ["LOGIN_PASSWORD"]
DXF = Path(os.environ.get(
    "MF_SUB_DXF", ROOT / "data/uploads/1. 입력도면 대명동 단위세대 계통도.dxf"))
# 첫 업로드는 반드시 평면도 슬롯으로 간다 — 슬롯 탭을 띄우기 위한 한 장.
PLAN = Path(os.environ.get(
    "MF_PLAN_DXF",
    ROOT / "routes/제출용[최종]/1. 입력도면 대명동 단위세대 평면도.dxf"))

# 미리보기 선 색 — module_f.js 의 drawSubPreview 와 같은 값이어야 한다.
PREVIEW = (0xFF, 0x2D, 0x2D)

problems: list[str] = []


def bad(m: str) -> None:
    problems.append(m)
    print("   !!", m)


def check(name: str, ok: bool, detail: str = "") -> bool:
    print(f"  [{'OK  ' if ok else 'FAIL'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        problems.append(name + (f" · {detail}" if detail else ""))
    return ok


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: bad(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad(f"pageerror: {e}"))
    # ★«500 이 났다» 만으로는 못 고친다 — 어느 요청인지 + 서버가 뭐라 했는지.
    failed = []
    page.on("response", lambda r: failed.append(r) if r.status >= 400 else None)

    def dump_failed():
        for r in failed:
            try:
                body = r.text()[:400]
            except Exception as exc:            # noqa: BLE001
                body = f"(본문 못 읽음: {exc})"
            bad(f"HTTP {r.status} {r.url[-60:]} — {body}")
        failed.clear()

    def px_count(rgb, tol=26):
        return page.evaluate("""([rgb, tol]) => {
          const cv = document.getElementById('cv');
          const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
          let n = 0, sx = 0, sy = 0;
          for (let y = 0; y < cv.height; y++)
            for (let x = 0; x < cv.width; x++) {
              const i = (y * cv.width + x) * 4;
              if (Math.abs(d[i]-rgb[0]) <= tol && Math.abs(d[i+1]-rgb[1]) <= tol
                  && Math.abs(d[i+2]-rgb[2]) <= tol) { n++; sx += x; sy += y; }
            }
          return {n, cx: n ? sx / n : 0, cy: n ? sy / n : 0};
        }""", [list(rgb), tol])

    print("[0] 로그인 · 계통도 슬롯으로 올리기")
    page.goto(f"{BASE}/module-f", wait_until="load")
    if "login" in page.url or page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    page.wait_for_selector("#slots", timeout=30_000)
    # ★슬롯 탭은 «도면을 연 뒤에야» 생긴다(첫 업로드는 늘 평면도 슬롯으로
    #   간다). 그래서 사람이 하는 그대로 평면도를 먼저 올리고, 계통도 슬롯으로
    #   바꾼 뒤 계통도를 올린다.
    page.set_input_files("#dxf", str(PLAN))
    page.click("#btn-open")
    page.wait_for_function(
        "() => document.querySelector('#busy').classList.contains('hidden')",
        timeout=600_000)
    page.wait_for_selector('#slots button', timeout=60_000)
    page.click('#slots button:has-text("계통도")')
    page.wait_for_timeout(1200)
    page.set_input_files("#dxf", str(DXF))
    page.click("#btn-open")
    try:
        page.wait_for_selector("#panel-sub:not(.hidden)", timeout=120_000)
    except Exception:                            # noqa: BLE001
        dump_failed()
        print("   상태줄:", page.inner_text("#status")[:120])
        browser.close()
        raise SystemExit(1)
    page.wait_for_function(
        "() => document.querySelector('#busy').classList.contains('hidden')",
        timeout=300_000)
    page.wait_for_timeout(1200)
    print("   상태줄:", page.inner_text("#status")[:80])

    print("[1] 경로 그래프를 받았나")
    box = page.eval_on_selector("#cv", """e => {
      const r = e.getBoundingClientRect();
      return {x: r.x, y: r.y, w: r.width, h: r.height};
    }""")
    page.click('h2.fold[data-fold="sub-layers-body"]')
    page.wait_for_timeout(400)
    info = page.inner_text("#sub-layers")[:120]
    check("경로 그래프 요약이 뜬다", "절점" in info, info.replace("\n", " ")[:90])
    n_lay = page.eval_on_selector_all(
        "#sub-layers input[data-lay]", "e => e.length")
    check("레이어 목록이 있다", n_lay > 0, f"{n_lay}종")

    print("[2] 첫 점 찍기 전에는 미리보기가 없다")
    before = px_count(PREVIEW)
    check("미리보기 선 없음", before["n"] < 30, f"{before['n']} 화소")

    print("[3] ① 을 찍고 마우스를 옮기면 선이 따라온다")
    page.click("#sub-pick-a")
    page.wait_for_timeout(200)
    # 도면 한가운데를 ① 로 찍는다.
    page.mouse.click(box["x"] + box["w"] * 0.5, box["y"] + box["h"] * 0.5)
    page.wait_for_timeout(300)
    page.mouse.move(box["x"] + box["w"] * 0.30, box["y"] + box["h"] * 0.30)
    page.wait_for_timeout(300)
    a = px_count(PREVIEW)
    check("마우스를 따라 선이 그려진다", a["n"] > 100, f"{a['n']} 화소")

    page.mouse.move(box["x"] + box["w"] * 0.72, box["y"] + box["h"] * 0.74)
    page.wait_for_timeout(300)
    bpx = px_count(PREVIEW)
    check("마우스를 옮기면 선이 달라진다",
          abs(bpx["n"] - a["n"]) > 20
          or abs(bpx["cx"] - a["cx"]) > 12 or abs(bpx["cy"] - a["cy"]) > 12,
          f"화소 {a['n']}→{bpx['n']} · 무게중심 "
          f"({a['cx']:.0f},{a['cy']:.0f})→({bpx['cx']:.0f},{bpx['cy']:.0f})")

    print("[4] ② 를 찍어 경로를 확정하고 추출")
    page.mouse.click(box["x"] + box["w"] * 0.72, box["y"] + box["h"] * 0.74)
    page.wait_for_timeout(400)
    fixed = px_count(PREVIEW)
    check("찍은 뒤에도 경로가 남는다", fixed["n"] > 100, f"{fixed['n']} 화소")
    page.screenshot(path=str(SHOT / "sub_live.png"))

    page.click("#sub-extract")
    page.wait_for_timeout(2500)
    summary = page.inner_text("#sub-summary").replace("\n", " ")
    print("   추출 결과:", summary[:110])
    check("추출이 성공했다", "실패" not in summary, summary[:80])

    print("[5] 레이어를 고르면 그래프가 실제로 줄어든다")
    n_before = page.evaluate(
        "() => (document.getElementById('sub-layers').innerText.match"
        "(/절점 (\\d+)/) || [0, 0])[1]")
    # 첫 레이어 하나를 끈다.
    page.eval_on_selector("#sub-layers input[data-lay]",
                          "e => { e.checked = false; e.dispatchEvent("
                          "new Event('change', {bubbles: true})); }")
    page.wait_for_timeout(1500)
    n_after = page.evaluate(
        "() => (document.getElementById('sub-layers').innerText.match"
        "(/절점 (\\d+)/) || [0, 0])[1]")
    check("레이어를 끄면 절점 수가 달라진다", str(n_before) != str(n_after),
          f"{n_before} → {n_after}")

    check("콘솔 오류 0",
          not [p for p in problems if p.startswith(("console", "pageerror"))])
    browser.close()

print()
if problems:
    print(f"실패 {len(problems)}건")
    for p in problems:
        print("  -", p)
    raise SystemExit(1)
print("모두 통과.")
