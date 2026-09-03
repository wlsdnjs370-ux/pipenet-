# -*- coding: utf-8 -*-
"""찍기 강조 — «사람이 찍은 배관» 이 실제로 밝은 빨강·굵게 그려지나.

색과 굵기는 «잘 보이나» 가 목적이라 구문 검사로는 아무것도 못 말한다.
캔버스 화소를 세어 확인한다:

  ① 찍기 전에는 그 빨강이 화면에 없다 (다른 것이 이미 빨간 게 아니다)
  ② 배관을 찍으면 그 빨강이 생긴다
  ③ 그 빨간 선이 바탕 도면 선보다 굵다 — 가로로 이어진 화소 폭으로 잰다
  ④ 헤드 삼각 기호는 빨강이 아니다 (배관과 헤드가 색으로 갈린다)

실행:
    MF_BASE=http://127.0.0.1:5065 LOGIN_PASSWORD=… \
        python scripts/_verify_module_f_pick_hl.py
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
    "MF_DXF",
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"))

problems: list[str] = []


def bad(msg: str) -> None:
    problems.append(msg)
    print("   !!", msg)


def check(name: str, ok: bool, detail: str = "") -> bool:
    print(f"  [{'OK  ' if ok else 'FAIL'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        problems.append(name + (f" · {detail}" if detail else ""))
    return ok


# 화면이 쓰는 값 — 여기 적힌 것과 코드가 갈리면 시험이 먼저 깨진다.
HL = (0xFF, 0x2D, 0x2D)


def count_and_width(page, rgb, tol=26):
    """그 색 화소 수 + 가로로 이어진 «가장 흔한» 폭(선 굵기의 대용)."""
    return page.evaluate("""([rgb, tol]) => {
      const cv = document.getElementById('cv');
      const g = cv.getContext('2d');
      const d = g.getImageData(0, 0, cv.width, cv.height).data;
      const near = (i) => Math.abs(d[i] - rgb[0]) <= tol
                       && Math.abs(d[i+1] - rgb[1]) <= tol
                       && Math.abs(d[i+2] - rgb[2]) <= tol;
      let n = 0;
      const runs = {};
      for (let y = 0; y < cv.height; y++) {
        let run = 0;
        for (let x = 0; x < cv.width; x++) {
          if (near((y * cv.width + x) * 4)) { n++; run++; }
          else if (run) { runs[run] = (runs[run] || 0) + 1; run = 0; }
        }
        if (run) runs[run] = (runs[run] || 0) + 1;
      }
      // 세로선이 만드는 가로폭 = 선 굵기. 가장 많이 나온 폭을 고른다.
      let best = 0, bestN = 0;
      for (const w in runs) {
        if (runs[w] > bestN || (runs[w] === bestN && +w > best)) {
          best = +w; bestN = runs[w];
        }
      }
      return {n, width: best, widthSeen: bestN};
    }""", [list(rgb), tol])


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: bad(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad(f"pageerror: {e}"))

    print("[0] 로그인 · 도면 올리기")
    page.goto(f"{BASE}/module-f", wait_until="load")
    if "login" in page.url or page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    page.set_input_files("#dxf", str(DXF))
    page.click("#btn-open")
    page.wait_for_selector("#panel-pick:not(.hidden)", timeout=300_000)
    page.wait_for_timeout(1200)
    print("   찍기 진입:", page.inner_text("#status")[:70])

    print("[1] 자동 흐름이 앉기를 기다린 뒤 «찍기» 단계로 간다")
    # 업로드하면 F-10a 흐름이 채택→조립까지 스스로 흘러간다. 그래서 찍기
    # 패널은 스쳐 지나간다 — 앉은 뒤에 단계바로 되돌아가야 강조를 볼 수 있다.
    page.wait_for_function(
        "() => document.querySelector('#busy').classList.contains('hidden')",
        timeout=600_000)
    page.wait_for_timeout(600)
    page.click('.steps div:text-is("찍기")')
    page.wait_for_selector("#panel-pick:not(.hidden)", timeout=60_000)
    page.wait_for_timeout(800)
    print("   상태줄:", page.inner_text("#status")[:76])

    print("[2] 이 빨강이 «도면 색» 이 아니라 «강조색» 인가")
    # 레이어 목록의 스와치가 각 묶음의 실제 색이다. 그중 강조색에 가까운 것이
    # 있으면 이 시험은 무의미해진다 — 그때는 다른 색을 골라야 한다.
    swatches = page.eval_on_selector_all(
        "#layers .sw", "es => es.map(e => getComputedStyle(e).backgroundColor)")
    def near_hl(css):
        nums = [int(v) for v in __import__("re").findall(r"\d+", css)[:3]]
        return len(nums) == 3 and all(
            abs(nums[k] - HL[k]) <= 26 for k in range(3))
    clash = [c for c in swatches if near_hl(c)]
    check("도면 레이어 중 강조색과 겹치는 것이 없다", not clash,
          f"레이어 {len(swatches)}개 · 겹침 {clash[:3]}")

    print("[3] 찍은 배관이 밝은 빨강으로, 바탕보다 굵게")
    # ★바탕 굵기를 «밝은 화소 전부» 로 재면 안 된다 — 그 안에 강조 빨강이
    #   같이 들어와 «강조 vs 강조» 를 비교하게 된다(실측으로 둘 다 2px 이
    #   나와 차이가 0 이었다). 레이어 스와치의 **실제 도면 색** 으로 잰다.
    red = count_and_width(page, HL)
    check("찍은 배관이 그 빨강으로 그려졌다", red["n"] > 500,
          f"{red['n']:,} 화소")

    import re as _re
    widths = []
    for css in swatches:
        nums = [int(v) for v in _re.findall(r"\d+", css)[:3]]
        if len(nums) != 3 or near_hl(css):
            continue
        got = count_and_width(page, tuple(nums), tol=12)
        if got["n"] > 400:                 # 화면에 충분히 그려진 레이어만
            widths.append((got["n"], got["width"], css))
    widths.sort(reverse=True)
    if not widths:
        bad("바탕 도면 색을 화면에서 못 쟀다")
        base_w = None
    else:
        base_w = max(w for _n, w, _c in widths[:5])
        for n, w, css in widths[:5]:
            print(f"   바탕 {css:<22} {n:>7,} 화소 · 선폭 {w}px")
    print(f"   강조 #ff2d2d           {red['n']:>7,} 화소 · 선폭 {red['width']}px")
    if base_w is not None:
        check("바탕 도면 선보다 굵다", red["width"] > base_w,
              f"강조 {red['width']}px vs 바탕(최대) {base_w}px")
        check("«살짝» 굵다 — 뭉개질 만큼은 아니다",
              red["width"] - base_w <= 3,
              f"차이 {red['width'] - base_w}px")

    print("[4] 배관과 헤드가 색으로 갈린다")
    check("배관색 ≠ 헤드색", HL != (0xFF, 0x5C, 0xF0),
          "#ff2d2d vs #ff5cf0")

    page.screenshot(path=str(SHOT / "pick_hl.png"))
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
