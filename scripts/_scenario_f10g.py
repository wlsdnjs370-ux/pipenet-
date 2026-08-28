# -*- coding: utf-8 -*-
"""[F-10g] 상무 시나리오를 B1F 로 «그대로» 한 번 완주하고 잰다.

지시서 §0.1 의 흐름을 순서대로 밟는다:

    업로드 → (자동) 손질 진입 → 알람밸브 클릭 → corridor
          → 흐린 배관 2건 살리기 → 다시 계산
          → 아이소 → 평면 밑그림에서 1건 살리기 → 다시 계산 → 산출

단계마다 **사람 클릭 수 · 걸린 시간 · 화면 캡처** 를 남긴다. 클릭은 전부
실제 마우스 클릭이다 — API 를 대신 부르면 「사람이 몇 번 눌렀나」가 거짓이 된다.

캔버스 좌표는 화면의 `fit()` 과 같은 식으로 계산한다(그 함수가 IIFE 안이라
밖에서 못 부른다). 그리고 누른 뒤 **서버가 실제로 받았는지** 로 확인한다 —
빗나간 클릭을 성공으로 세지 않기 위해서다.

    python scripts/_scenario_f10g.py [--dxf 경로] [--headed]
"""
from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SHOTS = ROOT / "data" / "_f10g_shots"
DEFAULT_DXF = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
PASSWORD = "f10gscenario"

CLICKS = 0
STEPS: list[dict] = []
PROBLEMS: list[str] = []


def click(page, sel_or_xy, note=""):
    """사람 클릭 한 번 — 셀렉터 또는 (x, y) 화면 좌표."""
    global CLICKS
    CLICKS += 1
    if isinstance(sel_or_xy, str):
        page.click(sel_or_xy)
    else:
        page.mouse.click(sel_or_xy[0], sel_or_xy[1])
    return CLICKS


def step(name, t0, page, shot=None, detail=""):
    el = time.perf_counter() - t0
    STEPS.append({"name": name, "clicks": CLICKS,
                  "sec": round(el, 1), "detail": detail})
    if shot:
        page.screenshot(path=str(SHOTS / shot))
    print(f"  [{len(STEPS):>2}] {name:<28} 누적클릭 {CLICKS:>3} · "
          f"{el:>6.1f}s  {detail}", flush=True)
    # ★단계마다 저장한다. 뒤에서 죽으면 앞의 실측까지 같이 잃는다 — 한 번
    #   도는 데 몇 분씩 걸리는 측정에서 그건 큰 손해다.
    save_record()


def save_record(total_min=None):
    SHOTS.mkdir(parents=True, exist_ok=True)
    (SHOTS / "_기록.json").write_text(json.dumps(
        {"clicks": CLICKS, "minutes": total_min,
         "steps": STEPS, "problems": PROBLEMS},
        ensure_ascii=False, indent=2), encoding="utf-8")


def free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


def api(page, path):
    return page.evaluate(
        "async (p) => (await (await fetch(p)).json())", path)


def wait_idle(page, limit=6000):
    """잡이 끝나고 가림막이 걷힐 때까지."""
    for _ in range(limit):
        page.wait_for_timeout(200)
        if page.is_hidden("#busy"):
            return True
    return False


def canvas_box(page):
    return page.eval_on_selector(
        "#cv", "el => { const r = el.parentElement.getBoundingClientRect();"
               " return {x:r.x, y:r.y, w:r.width, h:r.height}; }")


def to_screen(box, bounds, wx, wy):
    """화면의 `fit()` 과 같은 식 — sc = min(w/bw, h/bh)*0.92."""
    bw = max(1e-6, bounds["maxx"] - bounds["minx"])
    bh = max(1e-6, bounds["maxy"] - bounds["miny"])
    sc = min(box["w"] / bw, box["h"] / bh) * 0.92
    ox = bounds["minx"] - (box["w"] / sc - bw) / 2
    oy = bounds["miny"] - (box["h"] / sc - bh) / 2
    return (box["x"] + (wx - ox) * sc,
            box["y"] + box["h"] - (wy - oy) * sc)


def anchor_candidates(state, n=6):
    """알람밸브를 찍어 볼 자리들 — 큰 덩이부터, 헤드에 붙은 점.

    ★아무 배관 끝점이나 고르면 안 된다. B1F 는 배관이 306조각으로 흩어져 있어
      작은 조각에 놓으면 최불리가 «1개 · 0.18 m» 로 나온다 — 프로그램은 옳게
      돌았지만 시연으로는 아무 뜻이 없다. 그렇다고 «가장 큰 덩이» 가 곧
      급수 도달망인 것도 아니다(실측: 큰 덩이에 찍었더니 「닿는 헤드가
      없습니다」로 400). 그래서 **후보를 여럿 준비하고 결과를 보고 고른다** —
      사람도 결과가 이상하면 다시 찍는다. 그 클릭은 정직하게 센다.
    """
    heads = [(float(h[0]), float(h[1])) for h in (state.get("heads") or [])]
    if not heads:
        return []
    groups = sorted((g.get("segs") or [] for g in state.get("body_groups") or []),
                    key=len, reverse=True)
    out = []
    for s in groups[:n * 3]:
        pts = []
        for i in range(0, len(s) - 3, 4):
            pts.append((float(s[i]), float(s[i + 1])))
            pts.append((float(s[i + 2]), float(s[i + 3])))
        if not pts:
            continue
        best, bd = None, None
        for hx, hy in heads:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if bd is None or d < bd:
                best, bd = p, d
        if best is not None and bd <= 2000.0:
            out.append(best)
        if len(out) >= n:
            break
    return out


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--dxf", default=str(DEFAULT_DXF))
    ap.add_argument("--headed", action="store_true")
    a = ap.parse_args()
    from playwright.sync_api import sync_playwright

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1
    SHOTS.mkdir(parents=True, exist_ok=True)

    port = free_port()
    env = dict(os.environ)
    env.update({"PORT": str(port), "HOST": "127.0.0.1",
                "LOGIN_PASSWORD": PASSWORD, "PYTHONIOENCODING": "utf-8"})
    print(f"임시 서버 :{port} · 도면 {dxf.name} "
          f"({dxf.stat().st_size / 1e6:.1f} MB)\n")
    proc = subprocess.Popen([sys.executable, "serve.py"], cwd=str(ROOT),
                            env=env, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL)
    base = f"http://127.0.0.1:{port}"
    t_all = time.perf_counter()
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

            # sid 는 화면 JS 안(IIFE)에만 있다. 그것을 꺼내자고 프로덕션에
            # 전역을 심지 않는다 — 열기 응답에서 그대로 줍는다(기존 검증기 규약).
            seen: dict = {}

            def _grab(resp):
                if "/api/module-f/slot/open" not in resp.url:
                    return
                try:
                    seen["sid"] = (resp.json() or {}).get("sid")
                except Exception:  # noqa: BLE001
                    pass

            page.on("response", _grab)

            # 400/404 는 «어느 주소» 인지 알아야 고칠 수 있다 — 콘솔 문구만으론
            # 못 가린다.
            bad: list[str] = []

            def _bad(resp):
                if resp.status >= 400:
                    bad.append(f"{resp.status} {resp.url.split('/api/')[-1][:70]}")

            page.on("response", _bad)

            # 로그인 절차는 기존 검증기와 «같은» 길을 쓴다 — 다르게 하면
            # 왜 못 들어갔는지 가리는 데 시간을 쓰게 된다.
            page.goto(f"{base}/module-f", wait_until="load")
            if page.query_selector("input[type=password]"):
                page.fill("input[type=password]", PASSWORD)
                with page.expect_navigation(wait_until="load"):
                    page.click("button[type=submit]")
                page.goto(f"{base}/module-f", wait_until="load")
            if "login" in page.url:
                msg = (page.inner_text("body") or "")[:160].replace("\n", " ")
                print(f"★로그인을 못 넘었다 — {msg}")
                return 1
            page.wait_for_timeout(900)

            # ── ① 업로드 → 손질 진입 (질문 0)
            t0 = time.perf_counter()
            page.set_input_files("#dxf", str(dxf))   # 파일 고르기는 클릭 아님
            click(page, "#btn-open")
            # ★«잠깐 거쳐 가는» 화면을 도착으로 세면 안 된다. 채택이 끝나면
            #   화면은 찍기를 한 번 지나 조립으로 간다(adoptRun 의 순서).
            #   그래서 같은 화면이 «가만히 있는» 것을 확인하고 멈춘다.
            landed = None
            stable, last = 0, None
            for _ in range(4500):        # 200ms × 4500 = 15분
                page.wait_for_timeout(200)
                if not page.is_hidden("#busy"):
                    stable, last = 0, None
                    continue
                now = ("손질" if page.is_visible("#panel-edit")
                       else "찍기" if page.is_visible("#panel-pick") else None)
                if now is None:
                    stable, last = 0, None
                    continue
                stable = stable + 1 if now == last else 1
                last = now
                if now == "손질" or stable >= 25:      # 5초간 그대로면 도착
                    landed = now
                    break
            if landed != "손질":
                PROBLEMS.append(f"손질까지 못 갔다 — 도착 {landed}")
            note = page.inner_text("#start-note").strip().replace("\n", " ")
            step("① 업로드 → 손질 진입", t0, page, "1_손질진입.png",
                 f"도착 {landed} · {note[:44]}")

            # ── ② 알람밸브 원클릭 → corridor
            t0 = time.perf_counter()
            sid = seen.get("sid")
            if not sid:
                PROBLEMS.append("sid 를 못 얻었다")
                raise SystemExit(2)
            got = api(page, f"/api/module-f/edit/state?sid={sid}")
            if not got.get("state"):
                PROBLEMS.append(f"손질 상태를 못 받았다 — {str(got)[:120]}")
                raise SystemExit(2)
            state = got["state"]
            cands = anchor_candidates(state)
            box = canvas_box(page)
            w, tries = None, 0
            for p in cands:
                tries += 1
                click(page, to_screen(box, state["bounds"], p[0], p[1]))
                wait_idle(page)
                st2 = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
                w = st2.get("worst")
                # 「1개짜리 조각」에 걸린 것은 답이 아니다 — 다시 찍는다.
                if w and int(w.get("k") or 0) >= 2:
                    break
                w = None
            if not w:
                PROBLEMS.append(f"원클릭이 corridor 를 못 만들었다 "
                                f"(후보 {len(cands)}곳 시도)")
            step("② 알람밸브 원클릭 → corridor", t0, page, "2_corridor.png",
                 (f"최불리 {w['k']}개 · 최원 {w['far_m']} m · 찍은 횟수 {tries}"
                  if w else f"실패 ({tries}회 시도)"))

            # ── ③ 끊긴 곳 찾기 → 흐린 배관 2건 살리기
            t0 = time.perf_counter()
            click(page, "#ed-aj-scan")      # 끊긴 곳 찾기 — 손질 패널에 그대로 있다
            wait_idle(page)
            aj = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
            lines = ((aj.get("autojoin") or {}).get("lines")) or []
            # 이음 모드로 두 끝을 눌러 잇는다 — 사람이 하는 그 동작이다.
            click(page, '.emode[data-mode="이음"]')
            page.wait_for_timeout(400)
            mode_now = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]["mode"]
            ui_on = page.evaluate(
                """() => [...document.querySelectorAll('.emode')]
                     .filter(b => b.classList.contains('on'))
                     .map(b => b.dataset.mode).join(',')""")
            if mode_now != "이음" or ui_on != "이음":
                PROBLEMS.append(f"이음 모드가 안 걸렸다 — 서버 {mode_now} · "
                                f"화면 {ui_on!r}")
            joined = 0
            before = int(aj.get("edits_since_worst") or 0)
            trace = []
            for ln in lines[:6]:
                if joined >= 2:
                    break
                box = canvas_box(page)
                bb = aj["bounds"]
                for (wx_, wy_) in ((ln[0], ln[1]), (ln[2], ln[3])):
                    click(page, to_screen(box, bb, wx_, wy_))
                    page.wait_for_timeout(250)
                    # 클릭 하나하나 뒤에 픽이 살아 있는지 잰다 — 어느 클릭이
                    # 급수원을 지우는지 추측으로는 못 가린다.
                    q = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
                    trace.append(f"{len(q.get('sources') or [])}/"
                                 f"{len(q.get('valves') or [])}")
                cur = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
                if int(cur.get("edits_since_worst") or 0) > before:
                    joined += 1
                    before = int(cur.get("edits_since_worst") or 0)
            # ★픽이 살아 있는지 단계마다 잰다. 「급수 시작 위치를 먼저 찍어야」로
            #   끝나는 실패를 겪고 나서, 어디서 사라지는지 추측하지 않기 위해 넣었다.
            s3 = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
            step("③ 흐린 배관 살리기", t0, page, "3_살리기.png",
                 f"이은 곳 {joined}건 · 후보 {len(lines)}곳 · "
                 f"급수 {len(s3.get('sources') or [])} · "
                 f"밸브 {len(s3.get('valves') or [])} · "
                 f"클릭별 급수/밸브 {' '.join(trace)}")
            if joined < 1:
                PROBLEMS.append("한 곳도 못 이었다")

            # ── ④ 다시 계산
            t0 = time.perf_counter()
            badge = page.inner_text("#ed-edits").strip()
            click(page, "#ed-recalc")
            wait_idle(page)
            st4 = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
            w4 = st4.get("worst")
            step("④ 최불리 다시 계산", t0, page, "4_다시계산.png",
                 f"{badge} → 수정 {st4.get('edits_since_worst')}건 · "
                 + (f"최불리 {w4['k']}개" if w4 else "최불리 없음")
                 + f" · 급수 {len(st4.get('sources') or [])}")

            # ── ⑤ 수리계산 → 표 확정 (아이소)
            t0 = time.perf_counter()
            # 단계바로 간다 — 「수리계산 입력 →」 단추는 변환 패널에 있어서
            # 손질 화면에서는 안 보인다. 사람도 단계바를 쓴다.
            click(page, "#steps div:has-text('수리계산')")
            page.wait_for_selector("#panel-design:not(.hidden)", timeout=30000)
            click(page, "#dg-build")
            ok5 = wait_idle(page, 9000)
            summ = page.inner_text("#dg-summary").strip().replace("\n", " ")
            issues = page.inner_text("#dg-issues-n").strip()
            step("⑤ 표 확정 → 아이소", t0, page, "5_아이소.png",
                 f"확인할 것 {issues} · {summ[:40]}")

            # ── ⑥ 평면 밑그림에서 1건 살리기
            t0 = time.perf_counter()
            click(page, "#dg-plan")
            page.wait_for_timeout(600)
            click(page, '.dgmode[data-mode="이음"]')
            cur = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
            lines2 = ((cur.get("autojoin") or {}).get("lines")) or lines
            got6 = 0
            base6 = int(cur.get("edits_since_worst") or 0)
            for ln in lines2[:6]:
                if got6:
                    break
                box = canvas_box(page)
                for (wx_, wy_) in ((ln[0], ln[1]), (ln[2], ln[3])):
                    click(page, to_screen(box, cur["bounds"], wx_, wy_))
                    page.wait_for_timeout(250)
                c2 = api(page, f"/api/module-f/edit/state?sid={sid}")["state"]
                if int(c2.get("edits_since_worst") or 0) > base6:
                    got6 = 1
            step("⑥ 평면 밑그림에서 살리기", t0, page, "6_평면밑그림.png",
                 f"이은 곳 {got6}건")

            # ── ⑦ 다시 계산 → 아이소 갱신
            t0 = time.perf_counter()
            click(page, "#dg-recalc")
            wait_idle(page, 9000)
            step("⑦ 다시 계산 → 아이소 갱신", t0, page, "7_아이소갱신.png",
                 page.inner_text("#dg-issues-n").strip())

            # ── ⑧ 산출 — .sdf + .slf
            t0 = time.perf_counter()
            emit_disabled = page.is_disabled("#dg-emit")
            if not emit_disabled:
                click(page, "#dg-emit")
                wait_idle(page, 9000)
            step("⑧ 산출 (.sdf + .slf)", t0, page, "8_산출.png",
                 page.inner_text("#status").strip()[:50])

            if errors:
                # 잘라 적으면 무엇이 404 인지 못 가린다 — 주소까지 남긴다.
                PROBLEMS.append(f"콘솔 오류 {len(errors)}건")
                PROBLEMS.extend(f"  · {e}" for e in errors[:5])
            if bad:
                PROBLEMS.append(f"400 이상 응답 {len(bad)}건")
                PROBLEMS.extend(f"  · {b}" for b in dict.fromkeys(bad))
            browser.close()
    finally:
        proc.terminate()

    total = time.perf_counter() - t_all
    print(f"\n■ 완주 — 사람 클릭 **{CLICKS}회** · {total / 60:.1f}분")
    print(f"  캡처 {SHOTS}")
    if PROBLEMS:
        print("\n★문제")
        for p in PROBLEMS:
            print("  -", p)
    save_record(round(total / 60, 1))
    print(f"  기록 {SHOTS / '_기록.json'}")
    return 1 if PROBLEMS else 0


if __name__ == "__main__":
    sys.exit(main())
