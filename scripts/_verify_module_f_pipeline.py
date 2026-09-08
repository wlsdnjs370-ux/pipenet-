# -*- coding: utf-8 -*-
"""[전체 공정] 모듈 F 를 처음부터 끝까지 한 번에 태운다.

사용자 지적: 「평면도·계통도·기계실을 전부 통합하고 가압송수방식을 고르고
결합을 하면 아이소메트릭이 나와야 하는데, 애초에 결합 버튼이 활성화되지
않는다. 대대적으로 모듈 F 의 전체 공정이 돌아가야지.」

그래서 사람이 하는 순서 그대로 밟고, **어느 걸음에서 끊기는지** 를 찍는다.

  ① 평면도  올리기 → 찍기 → 배관망 구성 → 손질 → 알람밸브 → 최불리 → 표 확정
  ② 계통도  올리기 → 두 점 → 경로 추출
  ③ 기계실  올리기 → 두 점 → 경로 추출
  ④ 통합    급수방식 → **결합** → 아이소메트릭
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_D = os.path.join(_ROOT, "routes", "제출용[최종]")
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}"
          + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(f"{name} — {detail}")
    return ok


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def main() -> int:
    from playwright.sync_api import sync_playwright

    plan = os.path.join(_D, "1. 입력도면 대명동 단위세대 평면도.dxf")
    sysd = os.path.join(_D, "1. 입력도면 대명동 단위세대 계통도.dxf")
    mrd = os.path.join(_D, "1. 입력도면 대명동 단위세대 기계실.dxf")
    for p in (plan, sysd, mrd):
        if not os.path.isfile(p):
            print(f"표본 없음: {p}")
            return 0

    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        errs: list[str] = []
        pg.on("console", lambda m: errs.append(m.text)
              if m.type == "error" else None)
        pg.on("pageerror", lambda e: errs.append(f"pageerror: {e}"))
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(1000)

        def idle(n=6000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        def stage():
            return pg.evaluate("() => window.__mf.stage")

        def wait_stage(want, n=6000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy") and stage() == want:
                    return True
            return False

        def painted():
            """캔버스에 실제로 그려진 픽셀 수 — 빈 화면인지 세어 본다."""
            return pg.evaluate("""() => {
                const c = document.querySelector('#cv');
                const g = c.getContext('2d');
                const d = g.getImageData(0, 0, c.width, c.height).data;
                let n = 0;
                // 배경(거의 검정)과 다른 픽셀만 센다. 8픽셀 간격 표본.
                for (let i = 0; i < d.length; i += 4 * 8) {
                    if (d[i] > 40 || d[i + 1] > 40 || d[i + 2] > 40) n++;
                }
                return n;
            }""")

        def canvas_click_world(x, y):
            scr = pg.evaluate("""(q) => {
                const v = window.__mf.view;
                const r = document.querySelector('#cv').getBoundingClientRect();
                return [r.x + (q[0] - v.ox) * v.scale,
                        r.y + r.height - (q[1] - v.oy) * v.scale];
            }""", [x, y])
            pg.mouse.click(scr[0], scr[1])

        # ── ① 평면도 ────────────────────────────────────────────────
        print("[①] 평면도")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        check("찍기까지 흘러온다", wait_stage("pick"), stage())
        steps = pg.eval_on_selector_all("#steps div", "es => es.map(e => e.textContent)")
        check("단계 차례가 «손질 → 수리계산 → 변환»",
              steps == ["도면 열기", "찍기", "손질", "수리계산",
                        "수리계산 입력 변환"], str(steps))
        pg.wait_for_timeout(1200)
        print(f"      {pg.inner_text('#pk-count')[:60]}")
        pg.click("#pk-next")
        check("배관망 구성 → 손질", wait_stage("edit"), stage())
        pg.wait_for_timeout(1200)

        pt = pg.evaluate("""() => {
            for (const g of (window.__mf.edit.body_groups || [])) {
                const s = g.segs || [];
                if (s.length >= 4) return [s[0], s[1]];
            }
            return null;
        }""")
        canvas_click_world(pt[0], pt[1])
        idle()
        pg.wait_for_timeout(900)
        n_src = pg.evaluate("() => (window.__mf.edit.sources || []).length")
        check("알람밸브가 놓인다", n_src == 1, f"{n_src}곳")

        pg.click("#ed-worst")
        idle()
        pg.wait_for_timeout(1500)
        w = pg.evaluate("() => (window.__mf.edit || {}).worst || null")
        if not check("최불리가 나온다", bool(w),
                     pg.inner_text("#status")[:90]):
            br.close()
            return 1
        print(f"      최불리 {w['k']}개 · 최원 {w['far_m']} m")

        # 수리계산 표 확정 — 결합의 «평면도 재료» 가 여기서 난다.
        # (2026-09 순서 교정: 손질 다음이 수리계산이고, 변환은 그 뒤다.)
        pg.evaluate("() => { for (const d of "
                    "document.querySelectorAll('#steps div'))"
                    " { if (d.textContent.indexOf('수리계산') >= 0)"
                    " { d.click(); return; } } }")
        pg.wait_for_timeout(1500)
        check("수리계산 화면이 열린다", pg.is_visible("#dg-build"))
        # ★사용자 지적: 「손질까지 끝내고 수리계산 → 을 누르니 화면에 아무것도
        #   안 나온다」. 표를 확정하기 전에는 설계 좌표가 없어 캔버스가 검게
        #   비었다. 이제는 손질한 «평면» 을 먼저 보여 준다 — 실제로 무언가
        #   그려졌는지 픽셀로 센다(빈 화면은 «고장» 으로 읽힌다).
        check("표 확정 전에도 화면이 비지 않는다 (평면부터)",
              painted() > 500 and pg.is_checked("#dg-plan"),
              f"칠해진 픽셀 {painted()}")
        pg.click("#dg-build")
        idle()
        pg.wait_for_timeout(2500)
        got = pg.evaluate("() => !!(window.__mf.design "
                          "&& window.__mf.design.tables)")
        check("★표가 확정된다 (결합의 평면도 재료)", got,
              pg.inner_text("#status")[:90])

        # 확정 뒤 «평면에서 보기» 를 끄면 30° 아이소매트릭으로 바뀐다.
        pg.uncheck("#dg-plan")
        pg.wait_for_timeout(1200)
        iso = pg.evaluate("() => !!(window.__mf.design "
                          "&& window.__mf.design.view)")
        check("평면을 끄면 아이소매트릭이 나온다", iso and painted() > 500,
              f"view={iso} · 칠해진 픽셀 {painted()}")
        pg.check("#dg-plan")
        pg.wait_for_timeout(800)

        # ── ①-b 변환 — 이제 «표 다음» 이다. 재료가 갖춰졌으니 돌아야 한다.
        pg.evaluate("() => { for (const d of "
                    "document.querySelectorAll('#steps div'))"
                    " { if (d.textContent.indexOf('입력 변환') >= 0)"
                    " { d.click(); return; } } }")
        idle()
        pg.wait_for_timeout(1500)
        # 값 입력 창이 떠 있으면 기본값으로 닫는다(모듈 E 와 같은 자리).
        if pg.is_visible("#conv-cancel"):
            pg.click("#conv-cancel")
        why = pg.inner_text("#cv-why")
        check("변환이 «재료가 다 있다» 고 말한다", "모두 있습니다" in why, why[:70])
        pg.click("#btn-convert")
        idle()
        pg.wait_for_timeout(3000)
        info = " ".join(pg.inner_text("#conv-info")[:120].split())
        check("★변환이 돈다 (표 → 파일)",
              not pg.is_disabled("#btn-download")
              or not pg.is_disabled("#btn-download-design"), info)
        print(f"      변환: {info}")

        # ── ②③ 계통도 · 기계실 ──────────────────────────────────────
        for label, path in (("계통도", sysd), ("기계실", mrd)):
            print(f"[②③] {label}")
            pg.wait_for_selector("#slots button", timeout=120_000)
            pg.click(f'#slots button:has-text("{label}")')
            pg.wait_for_timeout(1200)
            pg.set_input_files("#dxf", path)
            pg.click("#btn-open")
            check(f"{label} 가 열린다", wait_stage("sub"), stage())
            pg.wait_for_timeout(1500)
            box = pg.eval_on_selector("#cv", """e => {
                const r = e.getBoundingClientRect();
                return {x: r.x, y: r.y, w: r.width, h: r.height};
            }""")
            pg.click("#sub-pick-a")
            pg.mouse.click(box["x"] + box["w"] * 0.5,
                           box["y"] + box["h"] * 0.82)
            pg.wait_for_timeout(400)
            pg.click("#sub-pick-b")
            pg.mouse.click(box["x"] + box["w"] * 0.5,
                           box["y"] + box["h"] * 0.18)
            pg.wait_for_timeout(400)
            # 두 점을 찍으면 «그 두 점이 있는 계통» 으로 좁혔는지 본다.
            g = pg.evaluate("() => { const g = window.__mf.subGraph || {};"
                            " return {chosen: g.chosen, auto: g.chosen_auto,"
                            " why: (g.narrowed || {}).reason,"
                            " autoLayers: g.auto_layers}; }")
            print(f"      계통 좁히기: {g}")
            pg.click("#sub-extract")
            idle()
            pg.wait_for_timeout(1500)
            s = pg.evaluate("() => (window.__mf.sub || {}).summary || null")
            check(f"{label} 경로가 뽑힌다", bool(s and s.get("pipes")),
                  pg.inner_text("#status")[:80])
            if s:
                print(f"      절점 {s['nodes']} · 배관 {s['pipes']}"
                      f" · 연장 {s['total_m']} m")

        # ── ④ 통합 ──────────────────────────────────────────────────
        # 통합은 단계바가 아니라 **머리말 단추** 다 — 세 슬롯이 같은 곳으로
        # 가므로 단계 끝에 되풀이하지 않는다.
        print("[④] 통합")
        check("단계바에 통합이 없다",
              pg.evaluate("() => [...document.querySelectorAll('#steps div')]"
                          ".every((d) => d.textContent.indexOf('통합') < 0)"))
        pg.click("#btn-merge")
        idle()
        pg.wait_for_timeout(1800)
        check("통합 화면이 열린다", pg.is_visible("#mg-build"))
        print(f"      재료: {pg.inner_text('#mg-ready')[:120]}")
        st = pg.evaluate("() => window.__mf.merge || null")
        print(f"      merge state: {st}")

        n_mode = pg.eval_on_selector_all("input[name=mg-mode]", "e => e.length")
        check("급수방식 목록이 있다", n_mode > 0, f"{n_mode}가지")
        if n_mode:
            pg.evaluate("() => document.querySelectorAll("
                        "'input[name=mg-mode]')[0].click()")
            idle()
            pg.wait_for_timeout(1500)
        dis = pg.is_disabled("#mg-build")
        check("★결합 단추가 켜진다", not dis,
              f"can_build={(pg.evaluate('() => (window.__mf.merge||{}).can_build'))}"
              f" · 재료 {pg.evaluate('() => (window.__mf.merge||{}).ready')}")
        if not dis:
            pg.click("#mg-build")
            idle()
            pg.wait_for_timeout(2500)
            merged = pg.evaluate("() => (window.__mf.merge || {}).merged")
            check("★결합이 돈다", bool(merged),
                  pg.inner_text("#status")[:100])
            # ★합친 것을 «보여» 주는가 — 숫자만으로는 세 도면이 제대로
            #   이어졌는지 사람이 판단할 수 없다.
            pg.wait_for_timeout(1200)
            mv = pg.evaluate("""() => { const v = window.__mf.mergeView;
                return v ? {n: v.nodes.length, p: v.pipes.length,
                            seam: v.pipes.filter(q => q.part === 'seam').length,
                            parts: [...new Set(v.nodes.map(q => q.part))]}
                         : null; }""")
            check("★결합망이 화면에 그려진다",
                  bool(mv and mv["n"] and painted() > 500), str(mv))
            # ★라이저 그림이 표를 따르는가 — 수직 · 길이 비례 (2026-09 교정).
            rz = pg.evaluate("""() => { const v = window.__mf.mergeView;
                if (!v) return null;
                const at = {}; for (const n of v.nodes) at[n.label] = n;
                const part = {}; for (const n of v.nodes) part[n.label] = n.part;
                const xs = new Set(v.nodes.filter(n => n.part === 'system')
                                   .map(n => Math.round(n.x)));
                const rows = v.pipes
                  .filter(p => part[p.a] === 'system' && part[p.b] === 'system')
                  .map(p => [p.len_m || 0, Math.hypot(at[p.a].x - at[p.b].x,
                                                      at[p.a].y - at[p.b].y)]);
                const t1 = rows.reduce((s, r) => s + r[0], 0) || 1;
                const t2 = rows.reduce((s, r) => s + r[1], 0) || 1;
                const bad = rows.filter(
                  r => Math.abs(r[0] / t1 - r[1] / t2) > 0.021).length;
                return {xkinds: xs.size, bad, n: rows.length}; }""")
            # [2026-09-08 · 사용자] 라이저 배치는 «균등 간격»(이전 디자인)으로
            # 되돌렸다 — 수직만 본다(비례는 규범이 아니다).
            check("라이저가 수직이다 (균등 간격)",
                  bool(rz and rz["xkinds"] == 1), str(rz))
            pg.check("#mg-iso")
            pg.wait_for_timeout(1200)
            rz2 = pg.evaluate("""() => { const v = window.__mf.mergeView;
                if (!v) return null;
                const xs = new Set(v.nodes.filter(n => n.part === 'system')
                                   .map(n => Math.round(n.x)));
                return xs.size; }""")
            check("아이소에서도 라이저가 수직이다", rz2 == 1, f"x 종류 {rz2}")
            # [D5] 결합 뒤 검사 — 화면 응답에 실려 오는가.
            ck = pg.evaluate("() => ((window.__mf.merge || {}).summary || {})"
                             ".checks || null")
            print(f"      결합 검사: {ck}")
            check("★결합 뒤 검사가 붙는다",
                  bool(ck and ck.get("components") == 1
                       and not ck.get("dangling_pipes_n")
                       and not ck.get("orphan_fittings")
                       and len(ck.get("inputs") or ()) == 1), str(ck))
            pg.uncheck("#mg-iso")
            pg.wait_for_timeout(800)
            print(f"      결합망: {mv} · 칠해진 픽셀 {painted()}")
            print(f"      범례: {' '.join(pg.inner_text('#mg-legend').split())}")
            pg.check("#mg-iso")
            pg.wait_for_timeout(1200)
            check("아이소로도 볼 수 있다",
                  pg.evaluate("() => !!window.__mf.mergeView")
                  and painted() > 500, f"칠해진 픽셀 {painted()}")
            pg.uncheck("#mg-iso")
            pg.wait_for_timeout(800)
            print(f"      요약: {pg.inner_text('#mg-summary')[:140]}")

        # ── 산출 — 결합 SDF 가 아이소 한 벌을 함께 내는가.
        if not pg.is_disabled("#mg-emit"):
            pg.click("#mg-emit")
            idle()
            pg.wait_for_timeout(2500)
            names = pg.evaluate("() => (window.__mf.mergeFiles || null)")
            print(f"      산출: {names}")
            check("★결합 산출에 아이소 한 벌이 있다",
                  bool(names and names.get("sdf_iso")), str(names))

        print("[콘솔]")
        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        pg.screenshot(path=os.path.join(_ROOT, "data", "_pipeline.png"))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("모듈 F 전 공정 — 화면에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
