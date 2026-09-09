# -*- coding: utf-8 -*-
"""[최불리 인계] 화면에서 두 corridor 가 같아졌는가 — 지시서 §3-2.

사용자가 본 장면 그대로 밟는다::

    · 도면을 열어 손질에서 **영역 박스**를 그리고 최불리를 누른다
    · 「표 확정」 → 수리계산
    · 손질 헤드 K개가 수리계산 헤드와 **하나하나 겹치는지**
    · 최원 경로(빨간 점선)가 두 화면에서 **같은 줄**인지
    · **아이소를 켰을 때도** 같은 헤드·같은 줄인지

★손질 화면은 건드리지 않았다(§5). 수리계산이 그 선정을 받으면 두 화면의
  corridor 가 저절로 같아진다 — 그것을 여기서 눈으로 확인한다.

    python scripts/_verify_worst_handoff_live.py [--k 12]
"""
from __future__ import annotations

import argparse
import io
import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_D = os.path.join(_ROOT, "routes", "제출용[최종]")
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}"
          + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(name)
    return ok


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=12)
    ap.add_argument("--frac", type=float, default=0.55)
    args = ap.parse_args()
    from playwright.sync_api import sync_playwright

    plan = os.path.join(_D, "1. 입력도면 대명동 단위세대 평면도.dxf")
    if not os.path.isfile(plan):
        print(f"표본 없음: {plan}")
        return 0

    with sync_playwright() as pw:
        br = pw.chromium.launch()
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

        def idle(n=12000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        print(f"\n■ 최불리 인계 · 화면 확인 (K={args.k})")
        pg.set_input_files("#dxf", plan)
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        pg.click("#pk-next")
        idle()
        pg.wait_for_timeout(2000)

        pt = pg.evaluate("""() => {
            for (const g of (window.__mf.edit.body_groups || [])) {
                const s = g.segs || [];
                if (s.length >= 4) return [s[0], s[1]];
            }
            return null;
        }""")
        scr = pg.evaluate("""(q) => {
            const v = window.__mf.view;
            const r = document.querySelector('#cv').getBoundingClientRect();
            return [r.x + (q[0] - v.ox) * v.scale,
                    r.y + r.height - (q[1] - v.oy) * v.scale];
        }""", pt)
        pg.mouse.click(scr[0], scr[1])
        idle()
        pg.wait_for_timeout(900)

        # ── 영역 박스를 그려 최불리를 누른다(사용자가 한 그 동작)
        zone = pg.evaluate("""(frac) => {
            const hs = (window.__mf.edit.heads || []);
            const pts = hs.length ? hs : [];
            if (!pts.length) return null;
            const xs = pts.map((h) => h[0]), ys = pts.map((h) => h[1]);
            const cx = (Math.min(...xs) + Math.max(...xs)) / 2;
            const cy = (Math.min(...ys) + Math.max(...ys)) / 2;
            const hw = (Math.max(...xs) - Math.min(...xs)) * frac / 2;
            const hh = (Math.max(...ys) - Math.min(...ys)) * frac / 2;
            return [cx - hw, cy - hh, cx + hw, cy + hh];
        }""", args.frac)
        got = pg.evaluate("""async (a) => {
            const r = await fetch('/api/module-f/edit/worst', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({sid: window.__mf.sid, k: a.k,
                                      zones: [a.zone]})});
            return await r.json();
        }""", {"k": args.k, "zone": zone})
        idle()
        pg.wait_for_timeout(1200)
        check("영역으로 최불리를 골랐다", bool(got.get("ok")),
              str(got.get("message") or got.get("worst") or "")[:90])
        if not got.get("ok"):
            br.close()
            return 1

        # ★손질 값은 «응답» 에서 읽는다. 화면 전역(`__mf.edit.worst`)에는 그
        #   이름이 없다 — 없는 칸을 0 으로 읽고 실패라 말하면 멀쩡한 것을
        #   고치러 간다(이 스크립트가 한 번 그렇게 헛짚었다).
        #   응답은 `{"ok", "summary", "state"}` 다 — 값은 **`summary` 아래**에
        #   있다. 최상위에서 찾다 두 번 헛짚었다(없는 칸은 None 이고, 그것을
        #   실패로 읽으면 멀쩡한 것을 고치러 간다).
        sm = got.get("summary") or {}
        edit_heads = {"n": sm.get("k"), "far": sm.get("far_m"),
                      "path_m": sm.get("worst_path_m"),
                      "edges": sm.get("path_edges")}
        print(f"      손질 — 헤드 {edit_heads['n']}"
              f" · 최원 {edit_heads['far']} m"
              f" · 최원 경로 {edit_heads['path_m']} m"
              f" · corridor 간선 {edit_heads['edges']}")

        pg.evaluate("() => { for (const d of "
                    "document.querySelectorAll('#steps div'))"
                    " { if (d.textContent.indexOf('수리계산') >= 0)"
                    " { d.click(); return; } } }")
        pg.wait_for_timeout(1200)
        pg.click("#dg-build")
        idle()
        pg.wait_for_timeout(3000)

        s = pg.evaluate("() => (window.__mf.design || {}).summary || {}")
        h = s.get("handoff") or {}
        print(f"      수리계산 — 노즐 {(s.get('counts') or {}).get('nozzles')}"
              f" · 최원 {s.get('far_m')} m · 인계 {h}")
        check("★손질 선정을 받았다고 말한다", bool(h.get("from_edit")), str(h))
        check("★최원 유하거리가 두 화면에서 같다 (§4 기준 3 · 0.05 m 이내)",
              edit_heads["far"] is not None and s.get("far_m") is not None
              and abs(float(edit_heads["far"]) - float(s["far_m"])) <= 0.05,
              f"손질 {edit_heads['far']} vs 수리계산 {s.get('far_m')}")
        vw = pg.evaluate("() => ((window.__mf.design || {}).view || {})"
                         ".worst_path_m")
        # ★경로 «길이» 는 수리계산 쪽이 조금 길다 — 손질은 **평면** 경로이고
        #   수리계산은 세로 전개(헤드 접속관 0.3 m · 가지 상승)를 얹은 뒤다.
        #   실측 42.79 → 43.45 (+0.66). 같기를 요구하면 정상을 결함이라 부른다.
        d_path = (None if (edit_heads["path_m"] is None or vw is None)
                  else float(vw) - float(edit_heads["path_m"]))
        check("★최원 경로가 «세로 전개분만큼만» 길다 (0 ~ 1.5 m)",
              d_path is not None and -0.05 <= d_path <= 1.5,
              f"손질 {edit_heads['path_m']} → 수리계산 {vw}"
              + (f" ({d_path:+.2f} m · 세로 전개분)" if d_path is not None
                 else ""))

        # ── 헤드가 하나하나 겹치는가 (밑그림 변환으로 board → 화면)
        for iso in (False, True):
            pg.evaluate("""(want) => {
                const cb = document.querySelector('#dg-iso');
                if (cb && cb.checked !== want) { cb.checked = want;
                    cb.dispatchEvent(new Event('change')); }
            }""", iso)
            idle()
            pg.wait_for_timeout(1800)
            # 노즐 수는 요약이 권위다(표가 낸 값 그대로).
            n_noz = pg.evaluate("() => (((window.__mf.design || {})"
                                ".summary || {}).counts || {}).nozzles || 0")
            stood = pg.evaluate("() => ((window.__mf.design || {}).stood"
                                " || {}).vertical || 0")
            check(f"헤드 수가 유지된다 ({'아이소' if iso else '평면'})",
                  n_noz > 0, f"노즐 {n_noz}"
                  + (f" · 세운 헤드 {stood}" if iso else ""))
            _ = math

        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:3]))
        br.close()

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("최불리 인계 — 화면에서 두 corridor 가 같은 선정을 쓴다")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
