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
# 5060·5061 은 크로미움이 막는 포트다(SIP) — goto 가 ERR_UNSAFE_PORT 로 죽는다.
# 검증용 서버는 안전한 포트로 띄운다.
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5065")
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

    print("[5-A] 헤드 종류 그림 표기 (모듈 E 의 그 PNG)")
    figs = page.evaluate("""() => {
      const out = [];
      for (const b of document.querySelectorAll('.kinds .ekind')) {
        const im = b.querySelector('img.diagram');
        out.push({
          kind: b.dataset.kind,
          w: im ? im.naturalWidth : 0,
          box: im ? Math.round(im.getBoundingClientRect().height) : 0,
          dot: getComputedStyle(b.querySelector('.dot')).backgroundColor,
          cnt: b.querySelector('.cnt').textContent,
        });
      }
      return out;
    }""")
    for f in figs:
        print(f"   {f['kind']}: 원본 {f['w']}px · 표시 {f['box']}px"
              f" · 점 {f['dot']} · {f['cnt']}")
    if len(figs) != 3:
        bad(f"헤드 종류 그림이 3장이 아니다 ({len(figs)})")
    for f in figs:
        if f["w"] < 10:
            bad(f"{f['kind']} 그림을 못 받아왔다 (naturalWidth {f['w']})")
        if f["box"] < 20:
            bad(f"{f['kind']} 그림이 화면에서 접혔다 (높이 {f['box']}px)")
        if f["dot"] in ("rgba(0, 0, 0, 0)", ""):
            bad(f"{f['kind']} 종류색 점이 안 칠해졌다")
    # 점 색은 캔버스 헤드 색과 같은 표에서 와야 한다 — 붙박이면 여기서 갈린다.
    same = page.evaluate("""() => {
      const pal = window.__mf.edit.palette.kinds, out = [];
      const hex = c => { const m = c.match(/\\d+/g);
        return '#' + m.slice(0, 3).map(v => (+v).toString(16).padStart(2, '0')).join(''); };
      for (const b of document.querySelectorAll('.kinds .ekind')) {
        out.push([b.dataset.kind, hex(getComputedStyle(
          b.querySelector('.dot')).backgroundColor), pal[b.dataset.kind]]);
      }
      return out;
    }""")
    for kind, shown, want in same:
        if shown.lower() != str(want).lower():
            bad(f"{kind} 점 색이 캔버스 색표와 다르다: {shown} ≠ {want}")
    print("   점 색 = 캔버스 색표:", same)
    page.screenshot(path=str(SHOT / "4b_kind_figures.png"))

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

    print("[5-A2] keep 규약 — 망이 안 바뀌는 동작 뒤에도 사본이 살아 있나")
    # 서버는 «안 바뀜» 이면 body_groups/heads/wet_pipes 를 비워 보낸다. 화면이
    # 사본을 못 지키면 망이 통째로 사라진다 — 픽셀로 확인한다.
    PIX = """() => {
      const c = document.getElementById('cv');
      const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
      let n = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] || d[i+1] || d[i+2]) n++;
      return n;
    }"""
    px_before = page.evaluate(PIX)
    page.click('.emode[data-mode="삭제"]')
    page.wait_for_timeout(500)
    page.click('.emode[data-mode="이음"]')
    page.wait_for_timeout(500)
    keep = page.evaluate("() => window.__mf.edit.keep")
    nb = page.evaluate("() => window.__mf.edit.body_groups.length")
    nh = page.evaluate("() => window.__mf.edit.heads.length")
    px_after = page.evaluate(PIX)
    print(f"   keep={keep} · 덩이 {nb} · 헤드점 {nh} · 픽셀 {px_before} → {px_after}")
    if not nb or not nh:
        bad(f"안 바뀐 응답 뒤 사본이 비었다 (덩이 {nb} · 헤드 {nh})")
    if abs(px_after - px_before) > max(200, px_before * 0.02):
        bad(f"안 바뀐 동작인데 그림이 달라졌다 ({px_before} → {px_after})")
    if "body_groups" not in (keep or []):
        bad(f"모드 전환인데 서버가 망 도형을 다시 보냈다 (keep={keep})")

    print("[5-A3] Ctrl+Z 한 박자 되돌리기")
    # ① 되돌릴 것이 없을 때 — 핸들러가 붙어 있으면 안내가 뜬다.
    #    실제 키를 눌러 확인한다(핸들러 유무는 런타임에서만 드러난다).
    page.keyboard.press("Control+z")
    page.wait_for_timeout(700)
    msg_empty = page.inner_text("#status")
    print("   되돌릴 것 없을 때:", msg_empty[:46])
    if "되돌" not in msg_empty:
        bad(f"Ctrl+Z 가 되돌리기로 가지 않았다: {msg_empty[:60]}")

    # ② 입력칸 안에서는 브라우저의 «글자 되돌리기» 를 빼앗으면 안 된다.
    hijack = page.evaluate("""() => {
      const inp = document.createElement('input');
      inp.type = 'text'; inp.value = 'abc';
      document.body.appendChild(inp); inp.focus();
      const ev = new KeyboardEvent('keydown', {key: 'z', ctrlKey: true,
                                               bubbles: true, cancelable: true});
      window.dispatchEvent(ev);
      const p = ev.defaultPrevented;
      inp.remove();
      return p;
    }""")
    print("   입력칸 안에서 가로챘나:", hijack)
    if hijack:
        bad("입력칸 안에서 Ctrl+Z 가 글자 되돌리기를 빼앗았다")

    # ③ 진짜로 한 박자 되돌아가나 — 급수원을 껐다가 Ctrl+Z 로 되살린다.
    #    (망 도형을 안 건드리므로 뒤 단계가 흔들리지 않고, 스스로 복구된다)
    src0 = page.evaluate("() => window.__mf.edit.sources.length")
    page.click('.emode[data-mode="급수시작위치"]')
    page.wait_for_timeout(400)
    spot = page.evaluate("""() => {
      const S = window.__mf, s = S.edit.sources[0];
      return { px: S.toScreenX(s[0]), py: S.toScreenY(s[1]) };
    }""")
    cbox = page.query_selector("#cv").bounding_box()
    page.mouse.click(cbox["x"] + spot["px"], cbox["y"] + spot["py"])
    page.wait_for_timeout(900)
    src1 = page.evaluate("() => window.__mf.edit.sources.length")
    page.keyboard.press("Control+z")
    page.wait_for_timeout(900)
    src2 = page.evaluate("() => window.__mf.edit.sources.length")
    print(f"   급수원 {src0} → 클릭 {src1} → Ctrl+Z {src2}")
    # 클릭이 급수원을 «켜는지 끄는지» 는 스냅이 어느 노드를 잡느냐에 달렸다
    # (실측: 화면맞춤 배율에서 옆 노드가 잡혀 하나 더 찍혔다). 방향은 상관없다 —
    # 되돌리기가 확인해야 할 것은 «클릭 직전 상태로 정확히 돌아가는가» 다.
    if src1 == src0:
        bad("클릭이 급수원을 바꾸지 못해 되돌리기를 시험하지 못했다")
    elif src2 != src0:
        bad(f"Ctrl+Z 가 한 박자 되돌리지 못했다 ({src1}→{src2}, 기대 {src0})")
    page.click('.emode[data-mode="이음"]')
    page.wait_for_timeout(400)


    print("[5-B] 자동 이음 — A 의 실측 · E 의 판정 · 점선 미리보기")
    stat = page.evaluate("() => window.__mf.edit.body_stat")
    print("   덩이·도달:", stat)
    page.click("#ed-aj-scan")
    page.wait_for_function(
        "window.__mf.edit.autojoin !== null "
        "&& window.__mf.edit.autojoin !== undefined", timeout=120_000)
    aj = page.evaluate("() => window.__mf.edit.autojoin")
    print(f"   여유 {aj['eps_mm']}mm(실측 {aj['auto_eps_mm']}) · 후보 {aj['n']}곳"
          f" {aj['by_kind']} · 관끝 {aj['ends']} · 방향맞음 {aj['kept']}/{aj['near']}")
    if not aj["n"]:
        bad("자동 이음 후보를 하나도 못 찾았다")
    if len(aj["lines"]) != aj["n"]:
        bad(f"후보 점선 좌표가 개수와 안 맞는다 ({len(aj['lines'])} ≠ {aj['n']})")
    n_eps = page.eval_on_selector_all("#ed-aj-eps option", "e => e.length")
    print("   여유 사다리 칸:", n_eps)
    if n_eps != 12:
        bad(f"여유 사다리가 12칸이 아니다 ({n_eps})")
    # 후보는 «아직 배관이 아니다» — 찾기만 해서는 망이 바뀌면 안 된다.
    edges_scan = page.evaluate("() => window.__mf.edit.counts.edges")
    page.screenshot(path=str(SHOT / "4c_autojoin_preview.png"))
    if page.evaluate("() => document.getElementById('ed-aj-apply').disabled"):
        bad("후보를 찾았는데 «모두 잇기» 가 잠긴 채다")
    page.click("#ed-aj-apply")
    # 가림막은 «화면 전체» 를 덮어야 한다. 캔버스만 덮으면 옆 패널 단추가
    # 작업 중에도 눌려 같은 작업이 두 번 돈다(서버도 막지만 화면이 1차 방벽).
    page.wait_for_timeout(400)
    cover = page.evaluate("""() => {
      const btn = document.getElementById('ed-aj-apply');
      const r = btn.getBoundingClientRect();
      const hit = document.elementFromPoint(r.left + r.width / 2,
                                            r.top + r.height / 2);
      const busy = document.getElementById('busy');
      return { hidden: busy.classList.contains('hidden'),
               covered: !!hit && (hit === busy || busy.contains(hit)),
               hit: hit ? (hit.id || hit.className || hit.tagName) : null };
    }""")
    print("   작업 중 가림막:", cover)
    if not cover["hidden"] and not cover["covered"]:
        bad(f"작업 중인데 «모두 잇기» 단추가 노출돼 있다 (hit={cover['hit']})")
    page.wait_for_function(
        "document.querySelector('#status').textContent.includes('자동 이음 —')",
        timeout=600_000)
    print("   적용:", page.inner_text("#status"))
    rep = page.evaluate("() => window.__mf.edit.autojoin_report")
    print("   결과:", rep)
    if not rep or not rep.get("made"):
        bad(f"자동 이음이 한 곳도 못 붙였다: {rep}")
    if rep and rep["bodies_after"] >= rep["bodies_before"]:
        bad(f"덩이가 줄지 않았다: {rep['bodies_before']} → {rep['bodies_after']}")
    edges_after = page.evaluate("() => window.__mf.edit.counts.edges")
    if edges_after <= edges_scan:
        bad(f"이었는데 간선이 안 늘었다 ({edges_scan} → {edges_after})")
    if page.evaluate("() => window.__mf.edit.autojoin"):
        bad("붙인 뒤에도 후보 점선이 남아 있다")
    page.screenshot(path=str(SHOT / "4d_autojoin_applied.png"))
    # 물흐름을 다시 돌려 성과를 눈으로 확인한다.
    page.click("#ed-flow")
    page.wait_for_function(
        "document.querySelector('#status').textContent.includes('물 닿은')",
        timeout=180_000)
    print("   붙인 뒤 물흐름:", page.inner_text("#status"))

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
    print(f"   설계면적 {w['k']}개 · corridor {len(w['corridor'])} 간선 ·"
          f" 앵커 {w['far_m']} m · 폭 {w['span_m']} m ·"
          f" 연장 {w['total_m']} m · 주배관 {w['max_load']}개 담당")
    if w["k"] != 30 or not w["corridor"]:
        bad(f"최불리망 결과가 비었다: k={w['k']} corridor={len(w.get('corridor', []))}")
    # 앵커(가장 불리한 지점)가 실려 화면에 그려져야 한다.
    if not w.get("anchor"):
        bad("앵커 헤드가 실리지 않았다")
    # corridor 간선마다 담당 헤드 수(load)가 붙고, 주배관은 여러 개를 먹인다.
    loads = [c[4] for c in w["corridor"]]
    if not loads or max(loads) != w["max_load"] or max(loads) < 2:
        bad(f"담당 헤드 수(load)가 corridor 에 안 실렸다: max={w['max_load']}")
    print(f"   load 분포: 최대 {max(loads)} · load=1 가지 {sum(1 for x in loads if x == 1)}개")
    if not page.is_checked("#cv-worst-kfp"):
        bad("최불리 선정 후 «최불리 .kfp» 가 자동으로 켜지지 않았다")
    page.wait_for_timeout(500)
    after = page.evaluate(WHITE)
    print(f"   흰 픽셀 {before} → {after} (증가 {after - before})")
    if after - before < 400:
        bad(f"최불리망이 캔버스에 그려지지 않았다 ({before} → {after})")
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
    # 「① (m)」 만으로는 어느 토막인지 못 읽는다 — 묶음 그림이 곧 이름표다.
    grpfigs = page.evaluate("""() => Array.from(
      document.querySelectorAll('#conv-fields .grpfig img'),
      im => [im.alt, im.naturalWidth,
             Math.round(im.getBoundingClientRect().height)])""")
    print("   변환 폼 그림:", grpfigs)
    if len(grpfigs) != 5:
        bad(f"변환 폼 그림이 5장이 아니다 ({len(grpfigs)})")
    for alt, w, h in grpfigs:
        if w < 10 or h < 40:
            bad(f"변환 폼 그림 «{alt}» 이 안 떴다 (원본 {w}px · 표시 {h}px)")
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
    # [F-4] 이 시점엔 설계 표가 아직이라 «최불리 .sdf» 는 끈다 — 켠 채 누르면
    # worst_required 안내로 수리계산 패널로 이동한다(그 흐름은 [9-A]가 본다).
    if page.is_checked("#cv-worst-sdf"):
        page.uncheck("#cv-worst-sdf")
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
    if "최불리 .kfp" not in conv or "K30" not in conv:
        bad(f"최불리 .kfp 가 변환되지 않았다: {conv[:160]}")
    if "전체망 .kfp" not in conv:
        bad(f"전체망 .kfp 가 변환되지 않았다: {conv[:160]}")
    for bid in ("btn-download-worst", "btn-download-set"):
        if page.evaluate(f"() => document.getElementById('{bid}').disabled"):
            bad(f"{bid} 가 잠긴 채로 남았다")

    print("[8] 내려받기 — 전체망 .kfp / 최불리 .kfp / 한 벌")
    sid = page.evaluate("() => window.__mf.sid")
    for what, head in (("kfp", b"{"), ("worst-kfp", b"{"), ("set", b"PK")):
        got = page.request.get(f"{BASE}/api/module-f/download?sid={sid}&what={what}")
        body = got.body()
        print(f"   {what:4s} HTTP {got.status} · {len(body):,} bytes · {body[:2]!r}")
        if got.status != 200 or not body.startswith(head):
            bad(f"{what} 내려받기 실패: HTTP {got.status} {len(body)}B")

    print("[9] 수리계산 입력 패널 (F-3)")
    page.click("#btn-to-design")
    page.wait_for_function(
        "document.querySelector('#st-design').classList.contains('on')",
        timeout=10_000)
    page.click("#dg-build")
    page.wait_for_function(
        "() => !document.getElementById('dg-emit').disabled",
        timeout=300_000)
    summary = page.inner_text("#dg-summary").replace("\n", " · ")
    print("   요약:", summary[:220])
    for word in ("설계면적", "앵커", "관경 근거", "제외 헤드"):
        if word not in summary:
            bad(f"요약에 «{word}» 가 없다")

    # 캔버스에 실제로 그려졌는가 — 검은 화면이면 미리보기가 아니다.
    drawn = page.evaluate("""() => {
      const c = document.getElementById('cv');
      const g = c.getContext('2d');
      const d = g.getImageData(0, 0, c.width, c.height).data;
      let on = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] || d[i+1] || d[i+2]) on++;
      return on;
    }""")
    print("   미리보기 픽셀:", drawn)
    if drawn < 1000:
        bad(f"설계 미리보기가 안 그려졌다 — 켜진 픽셀 {drawn}")
    page.screenshot(path=str(SHOT / "9_design_iso.png"))

    # 보기 설정 변경 — 표 값은 그대로, 그림만 바뀐다(표시 전용 증명).
    tbl_before = page.inner_text("#dg-grid")[:4000]
    page.fill("#dg-canvas", "6000")
    page.dispatch_event("#dg-canvas", "change")
    page.wait_for_timeout(700)
    tbl_after = page.inner_text("#dg-grid")[:4000]
    if tbl_before != tbl_after:
        bad("보기 설정 변경이 표 값을 바꿨다 — 표시 전용이 아니다")
    page.fill("#dg-canvas", "3000")
    page.dispatch_event("#dg-canvas", "change")
    page.wait_for_timeout(700)

    # 표 4종 전환 + 배관 표의 관경 근거 열
    page.select_option("#dg-table", "pipes")
    head_row = page.inner_text("#dg-grid table thead")
    if "관경 근거" not in head_row:
        bad(f"배관 표에 관경 근거 열이 없다: {head_row[:120]}")
    for which in ("nodes", "nozzles", "fittings"):
        page.select_option("#dg-table", which)
        n_rows = page.eval_on_selector_all("#dg-grid tbody tr", "e => e.length")
        print(f"   표 {which}: {n_rows}행")
        if n_rows == 0:
            bad(f"표 {which} 가 비었다")
    page.select_option("#dg-table", "pipes")
    page.screenshot(path=str(SHOT / "9_design_table.png"))

    # 저장 → 내려받기 (zip 한 벌)
    page.click("#dg-emit")
    page.wait_for_function(
        "() => !document.getElementById('dg-download').disabled",
        timeout=120_000)
    got = page.request.get(
        f"{BASE}/api/module-f/download?sid={sid}&what=design")
    body = got.body()
    print(f"   design zip HTTP {got.status} · {len(body):,} bytes")
    if got.status != 200 or not body.startswith(b"PK"):
        bad(f"design 내려받기 실패: HTTP {got.status}")

    print("[9-A] 표 확정 뒤 — «최불리 .sdf» 체크가 변환에서 통한다")
    page.click("#dg-back")      # 변환 패널로
    page.wait_for_function(
        "document.querySelector('#st-conv').classList.contains('on')",
        timeout=10_000)
    page.uncheck("#cv-full-kfp")
    page.uncheck("#cv-worst-kfp")
    page.check("#cv-worst-sdf")
    page.click("#btn-convert")
    page.wait_for_function(
        "!document.querySelector('#btn-download-design').disabled "
        "|| document.querySelector('#conv-info').textContent.includes('막힘')",
        timeout=300_000)
    conv2 = page.inner_text("#conv-info").replace("\n", " ")
    if "최불리 .sdf" not in conv2:
        bad(f"최불리 .sdf 만 고른 변환이 실패: {conv2[:160]}")
    if page.evaluate("() => document.getElementById('btn-download').disabled") is False:
        bad(".sdf 만 골랐는데 전체망 .kfp 버튼이 켜졌다")
    got = page.request.get(f"{BASE}/api/module-f/download?sid={sid}&what=design")
    if got.status != 200 or not got.body().startswith(b"PK"):
        bad(f"design 내려받기 실패: HTTP {got.status}")
    page.screenshot(path=str(SHOT / "10_outputs.png"))

    browser.close()

print()
if problems:
    print(f"실패 {len(problems)}건")
    for p in problems:
        print("  -", p)
    sys.exit(1)
print("브라우저 검증 통과 · 스크린샷:", SHOT)
