# -*- coding: utf-8 -*-
"""모듈 F 실제 브라우저 런타임 검증.

구문 검사만으로는 함수-지역 헬퍼를 다른 스코프에서 부르는 종류의 회귀를
못 잡는다(전례 있음). 콘솔 오류·pageerror 를 한 건이라도 잡으면 실패로 본다.
캔버스 좌표 왕복(세계 ↔ 화면)도 실제 클릭으로 확인한다.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
SHOT = ROOT / "data" / "_mf_shots"
SHOT.mkdir(parents=True, exist_ok=True)
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5061")
PASSWORD = os.environ["LOGIN_PASSWORD"]
SMALL_DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
SAVED_KEY = "B1F 현장조사 소화설비 평면도"

problems: list[str] = []


def bad(msg: str) -> None:
    problems.append(msg)
    print("   !!", msg)


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: bad(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad(f"pageerror: {e}"))

    print("[0] 로그인")
    page.goto(f"{BASE}/module-f", wait_until="load")
    if "login" in page.url or page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    print("   url:", page.url)

    print("[1] 초기 화면")
    print("   단계표:", page.eval_on_selector_all(
        ".steps div", "els => els.map(e => e.textContent)"))
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=15_000)
    opts = page.eval_on_selector_all("#saved option", "e => e.map(o => o.value)")
    print("   저장된 찍기:", len(opts), "→", opts[:3])
    if SAVED_KEY not in opts:
        bad(f"저장 목록에 {SAVED_KEY} 가 없다")
    page.screenshot(path=str(SHOT / "1_open.png"))

    print("[2] DXF 업로드 → 찍기 단계")
    page.set_input_files("#dxf", str(SMALL_DXF))
    page.click("#btn-open")
    page.wait_for_function(
        "document.querySelector('#st-pick').classList.contains('on')",
        timeout=180_000)
    print("   상태줄:", page.inner_text("#status"))
    info = page.inner_text("#pk-info").replace("\n", " · ")
    print("   찍기 패널:", info)
    n_layers = page.eval_on_selector_all("#layers label", "e => e.length")
    print("   레이어 묶음:", n_layers)
    if not n_layers:
        bad("레이어 목록이 비었다")
    drawn = page.evaluate("""() => {
      const c = document.getElementById('cv');
      const g = c.getContext('2d');
      const d = g.getImageData(0, 0, c.width, c.height).data;
      let on = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] || d[i+1] || d[i+2]) on++;
      return on;
    }""")
    print("   캔버스에 찍힌 픽셀:", drawn)
    if drawn < 500:
        bad(f"도면이 캔버스에 그려지지 않았다 (픽셀 {drawn})")
    page.screenshot(path=str(SHOT / "2_pick_loaded.png"))

    print("[2-A] 모듈 A 레이어 자동 추천")
    cats = page.inner_text("#pk-cats").replace("\n", " ")
    print("   카테고리 막대:", cats)
    if "PIPE" not in cats:
        bad("레이어 카테고리 요약이 그려지지 않았다")
    badges = page.eval_on_selector_all(
        "#layers .cat", "e => e.map(x => x.textContent)")
    print("   레이어 뱃지:", len(badges), "종", sorted(set(badges)))
    if not badges:
        bad("레이어 목록에 카테고리 뱃지가 없다")
    page.click("#pk-auto-pipe")
    page.wait_for_timeout(1500)
    print("   추천 일괄:", page.inner_text("#status")[:110])
    n_mat = page.evaluate("() => window.__mf.pick.materials.length")
    print("   잡힌 재료 묶음:", n_mat)
    if n_mat < 1:
        bad("배관 추천 일괄 찍기가 아무것도 잡지 못했다")
    page.screenshot(path=str(SHOT / "2b_auto_pick.png"))
    # 다음 검사가 깨끗한 판에서 시작하도록 찍은 만큼 전부 되돌린다.
    for _ in range(n_mat):
        page.click("#pk-undo")
        page.wait_for_timeout(350)
    left = page.evaluate("() => window.__mf.pick.materials.length")
    if left:
        bad(f"추천 일괄을 되돌리지 못했다 (남은 {left}묶음)")

    print("[3] 캔버스 클릭 — 세계↔화면 좌표 왕복")
    # 가장 선이 많은 묶음의 첫 선분 중점을 화면좌표로 환산해 그 자리를 클릭한다.
    spot = page.evaluate("""() => {
      const S = window.__mf;
      const b = S.world.bundles.reduce((a, x) => x.n_seg > a.n_seg ? x : a);
      const s = b.segs;
      const mx = (s[0] + s[2]) / 2, my = (s[1] + s[3]) / 2;
      return { mx, my, px: S.toScreenX(mx), py: S.toScreenY(my),
               layer: b.layer, name: b.name };
    }""")
    print("   목표 묶음:", spot["layer"], "×", spot["name"],
          "| 세계", round(spot["mx"], 1), round(spot["my"], 1),
          "→ 화면", round(spot["px"], 1), round(spot["py"], 1))
    box = page.query_selector("#cv").bounding_box()
    page.mouse.click(box["x"] + spot["px"], box["y"] + spot["py"])
    page.wait_for_timeout(1200)
    print("   클릭 결과:", page.inner_text("#status"))
    picked = page.evaluate("() => window.__mf.pick.materials.length")
    print("   잡힌 재료 묶음:", picked)
    if picked < 1:
        bad("캔버스 클릭이 재료를 잡지 못했다 — 좌표 환산 의심")
    page.screenshot(path=str(SHOT / "3_pick_clicked.png"))

    print("[4] 되돌리기 · 선택완료 · 헤드 칸")
    page.click("#pk-undo")
    page.wait_for_timeout(700)
    if page.evaluate("() => window.__mf.pick.materials.length") != 0:
        bad("되돌리기가 재료를 되돌리지 못했다")
    page.mouse.click(box["x"] + spot["px"], box["y"] + spot["py"])
    page.wait_for_timeout(1000)
    page.click("#pk-done")
    page.wait_for_timeout(700)
    print("   선택완료 후:", page.inner_text("#status"))
    if not page.evaluate("() => window.__mf.pick.mat_done"):
        bad("선택완료가 반영되지 않았다")
    page.click('.slot[data-slot="상하향"]')
    page.wait_for_timeout(600)
    print("   헤드 칸:", page.evaluate("() => window.__mf.pick.head_label"))
    if page.evaluate("() => document.getElementById('pk-next').disabled"):
        bad("다음 단추가 여전히 잠겨 있다")

    print("[5] 저장본으로 손질 단계 열기")
    page.goto(f"{BASE}/module-f", wait_until="load")
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=15_000)
    page.select_option("#saved", SAVED_KEY)
    page.click("#btn-reopen")
    page.wait_for_function(
        "document.querySelector('#st-edit').classList.contains('on')",
        timeout=300_000)
    print("   상태줄:", page.inner_text("#status"))
    print("   손질 패널:", page.inner_text("#ed-info").replace("\n", " · "))
    drawn = page.evaluate("""() => {
      const c = document.getElementById('cv');
      const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
      let on = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] || d[i+1] || d[i+2]) on++;
      return on;
    }""")
    print("   캔버스 픽셀:", drawn)
    if drawn < 500:
        bad(f"손질망이 그려지지 않았다 (픽셀 {drawn})")
    page.screenshot(path=str(SHOT / "4_edit.png"))

    print("[6] 손질 모드 전환 · 물흐름")
    for mode in ("삭제", "급수시작위치", "알람밸브위치", "이음"):
        page.click(f'.emode[data-mode="{mode}"]')
        page.wait_for_timeout(400)
        got = page.evaluate("() => window.__mf.edit.mode")
        if got != mode:
            bad(f"모드 전환 실패: {mode} → {got}")
    bodies = page.evaluate("() => window.__mf.edit.body_groups.length")
    print("   모드 전환 후 덩이 유지:", bodies)
    if bodies < 1:
        bad("모드만 바꿨는데 망 표시가 사라졌다")
    page.click("#ed-flow")
    page.wait_for_function(
        "document.querySelector('#status').textContent.includes('물 닿은')",
        timeout=120_000)
    print("   물흐름:", page.inner_text("#status"))
    wet = page.evaluate("() => window.__mf.edit.wet_pipes.length")
    print("   젖은 배관:", wet)
    if wet < 1:
        bad("물흐름 후 물길이 화면에 남지 않았다")
    page.screenshot(path=str(SHOT / "5_flow.png"))

    # 값만 들고 있는 것과 실제로 그려지는 것은 다르다 — 급수원 둘레를 확대해
    # 물길 색(#22b573) 픽셀이 캔버스에 실제로 찍혔는지 센다.
    page.evaluate("""() => {
      const S = window.__mf, s = S.edit.sources[0];
      S.view.scale = 0.06;
      S.view.ox = s[0] - 800 / S.view.scale / 2;
      S.view.oy = s[1] - 700 / S.view.scale / 2;
    }""")
    page.evaluate("() => window.dispatchEvent(new Event('resize'))")
    page.wait_for_timeout(600)
    wetpx = page.evaluate("""() => {
      const c = document.getElementById('cv');
      const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
      let n = 0;
      for (let i = 0; i < d.length; i += 4) {
        if (Math.abs(d[i] - 34) < 26 && Math.abs(d[i+1] - 181) < 26
            && Math.abs(d[i+2] - 115) < 26) n++;
      }
      return n;
    }""")
    print("   물길 픽셀:", wetpx)
    if wetpx < 200:
        bad(f"물길이 캔버스에 그려지지 않았다 (픽셀 {wetpx})")
    page.screenshot(path=str(SHOT / "8_flow_zoom.png"))
    page.click("#btn-fit")
    page.wait_for_timeout(400)

    print("[6-A] Remote 30 최불리 헤드")
    # 경로는 급수원에서 100 m 넘게 뻗으므로 근접 확대로는 못 잰다.
    # 화면 맞춤 상태에서 «선정 전 → 후» 흰 픽셀 증가분으로 그려졌는지 본다.
    page.click("#btn-fit")
    page.wait_for_timeout(400)
    WHITE = """() => {
      const c = document.getElementById('cv');
      const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
      let n = 0;
      for (let i = 0; i < d.length; i += 4) {
        if (Math.min(d[i], d[i+1], d[i+2]) > 235) n++;
      }
      return n;
    }"""
    before = page.evaluate(WHITE)
    page.click("#ed-worst")
    page.wait_for_function(
        "document.querySelector('#status').textContent.includes('최불리')",
        timeout=120_000)
    print("   선정:", page.inner_text("#status"))
    w = page.evaluate("() => window.__mf.edit.worst")
    print(f"   헤드 {w['k']} · 경로 {len(w['path'])} 간선 ·"
          f" 최원 {w['far_m']} m · 끝 {w['near_m']} m")
    if w["k"] != 30 or not w["path"]:
        bad(f"최불리 선정 결과가 비었다: {w}")
    if not page.is_checked("#cv-remote"):
        bad("최불리 선정 후 «최불리 30만 변환» 이 자동으로 켜지지 않았다")
    page.wait_for_timeout(500)
    after = page.evaluate(WHITE)
    print(f"   흰 픽셀 {before} → {after} (증가 {after - before})")
    if after - before < 400:
        bad(f"최불리 경로가 캔버스에 그려지지 않았다 ({before} → {after})")
    page.screenshot(path=str(SHOT / "9_worst30.png"))

    print("[7] 변환 단계")
    page.click("#ed-next")
    page.wait_for_function(
        "document.querySelector('#st-conv').classList.contains('on')", timeout=20_000)
    # 폼은 setStage 뒤 비동기로 채워진다 — 다 그려질 때까지 기다린다.
    page.wait_for_function(
        "document.querySelectorAll('#conv-fields input').length >= 12", timeout=30_000)
    n_fields = page.eval_on_selector_all("#conv-fields input", "e => e.length")
    print("   변환 폼 칸:", n_fields)
    if n_fields < 12:
        bad(f"변환 폼이 덜 그려졌다 ({n_fields}칸)")
    conv_px = page.evaluate("""() => {
      const c = document.getElementById('cv');
      const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
      let on = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] || d[i+1] || d[i+2]) on++;
      return on;
    }""")
    print("   변환 단계 캔버스 픽셀:", conv_px)
    if conv_px < 500:
        bad(f"변환 단계에서 망이 사라졌다 (픽셀 {conv_px})")
    page.screenshot(path=str(SHOT / "6_convert_form.png"))
    page.click("#btn-convert")
    page.wait_for_function(
        "!document.querySelector('#btn-download').disabled "
        "|| document.querySelector('#conv-info').textContent.includes('막힘')",
        timeout=300_000)
    print("   변환 결과:", page.inner_text("#conv-info").replace("\n", " · ")[:300])
    if page.evaluate("() => document.getElementById('btn-download').disabled"):
        bad("변환이 완료되지 않았다")
    page.screenshot(path=str(SHOT / "7_converted.png"))

    conv = page.inner_text("#conv-info").replace("\n", " ")
    if "최불리 30 헤드" not in conv:
        bad(f"최불리 범위로 변환되지 않았다: {conv[:160]}")
    for bid in ("btn-download-sdf", "btn-download-set"):
        if page.evaluate(f"() => document.getElementById('{bid}').disabled"):
            bad(f"{bid} 가 잠긴 채로 남았다")

    print("[8] 내려받기 — .kfp / .sdf / 한 벌")
    sid = page.evaluate("() => window.__mf.sid")
    for what, head in (("kfp", b"{"), ("sdf", b"<"), ("set", b"PK")):
        got = page.request.get(f"{BASE}/api/module-f/download?sid={sid}&what={what}")
        body = got.body()
        print(f"   {what:4s} HTTP {got.status} · {len(body):,} bytes · {body[:2]!r}")
        if got.status != 200 or not body.startswith(head):
            bad(f"{what} 내려받기 실패: HTTP {got.status} {len(body)}B")

    browser.close()

print()
if problems:
    print(f"실패 {len(problems)}건")
    for p in problems:
        print("  -", p)
    sys.exit(1)
print("브라우저 검증 통과 · 스크린샷:", SHOT)
