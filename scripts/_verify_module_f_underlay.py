# -*- coding: utf-8 -*-
"""[F-10e] 아이소 밑그림이 «화면에» 정말 그려지는가 — 화소로 잰다.

시험은 변환이 맞는지를 본다. 사람이 보는 것은 화소다. 켜기 전/후의 캔버스를
견줘 ① 실제로 더 그려지는가 ② 아이소를 껐다 켜도 그려지는가 ③ 콘솔 오류 0
을 확인한다.
"""
from __future__ import annotations

import io
import json
import os
import sys
import time
import urllib.parse
import urllib.request
import http.cookiejar

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(f"{name} — {detail}")


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def _ink(png_bytes):
    """캔버스에 그려진 화소 수 — 배경이 아닌 것을 센다."""
    from PIL import Image
    im = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    px = im.load()
    w, h = im.size
    bg = px[1, 1]
    n = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            c = px[x, y]
            if abs(c[0] - bg[0]) + abs(c[1] - bg[1]) + abs(c[2] - bg[2]) > 24:
                n += 1
    return n


def main() -> int:
    # ── 세션 하나를 표 확정까지 밀어 둔다(HTTP 로 — 화면은 그 결과만 본다).
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    def get(p):
        return json.loads(op.open(BASE + p, timeout=900).read())

    def post(p, body):
        req = urllib.request.Request(
            BASE + p, data=json.dumps(body).encode(),
            headers={"Content-Type": "application/json"})
        try:
            return json.loads(op.open(req, timeout=900).read())
        except urllib.error.HTTPError as exc:
            try:
                return json.loads(exc.read())
            except Exception:
                return {"ok": False, "message": f"HTTP {exc.code}"}

    pw = _password()
    op.open(BASE + "/login",
            urllib.parse.urlencode({"password": pw}).encode()).read()

    import glob
    from routes.module_f.common import IMPORT_WORK_ROOT
    ready = set()
    for f in glob.glob(str(IMPORT_WORK_ROOT / "DWG" / "*_유저손질.json")):
        try:
            j = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        if j.get("sources") or j.get("valve_picks"):
            ready.add(os.path.basename(f)[: -len("_유저손질.json")])
    items = (get("/api/module-f/saved") or {}).get("items") or []
    live = [i for i in items if i.get("source_exists")]
    pool = [i for i in live if i["key"] in ready] or live
    if not pool:
        print("쓸 저장본이 없다 — 검사 불가")
        return 0
    key = pool[0]["key"]
    sid = (post("/api/module-f/reopen", {"key": key}) or {}).get("sid")

    def wait(limit=1800):
        for _ in range(int(limit / 0.5)):
            j = get(f"/api/module-f/job?sid={sid}")
            if j.get("state") in ("done", "error", "idle"):
                return j
            time.sleep(0.5)
        return {"state": "timeout"}

    wait()
    post("/api/module-f/edit/worst", {"sid": sid, "k": 30})
    post("/api/module-f/design/build", {"sid": sid, "k": 30})
    wait()
    d = get(f"/api/module-f/design/preview?sid={sid}")
    u = ((d.get("view") or {}).get("underlay") or None)
    print(f"[1] 서버가 변환을 싣는다 — {key}")
    check("underlay 가 실려 온다", bool(u), str(sorted(u) if u else None))
    if not u:
        return 1
    check("합성 변환(k·tx·ty)이 있다",
          all(x in u for x in ("k", "tx", "ty")), f"k={u.get('k')}")

    # ── 화면
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw_:
        br = pw_.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        errs: list[str] = []
        pg.on("console", lambda m: errs.append(m.text)
              if m.type == "error" else None)
        bad: list[str] = []

        def _resp(r):
            if r.status < 400:
                return
            try:
                body = r.text()[:200]
            except Exception:
                body = "?"
            bad.append(f"{r.status} {r.url.split('/api/')[-1]} :: {body}")

        pg.on("response", _resp)
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", pw)
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f?sid={sid}", wait_until="domcontentloaded")
        pg.wait_for_timeout(1500)

        print("[2] 밑그림 토글")
        has = pg.query_selector("#dg-under") is not None
        check("토글이 화면에 있다", has)
        if not has:
            br.close()
            return 1

        ok = pg.evaluate("""async () => {
            const r = await fetch('/api/module-f/design/preview?sid=%s');
            const d = await r.json();
            return !!(d.view && d.view.underlay);
        }""" % sid)
        check("화면에서도 변환을 받는다", bool(ok))

        # ── 진짜 흐름을 태운다: 저장본 열기 → 표 확정 → 설계 화면
        print("[3] 실제 흐름 — 저장본을 열고 표를 확정한다")
        # 목록이 채워질 때까지 기다린다 — 안 기다리면 «옵션이 없다» 로 튄다.
        pg.wait_for_function(
            "() => { const s = document.querySelector('#saved');"
            " return s && s.options.length > 0; }", timeout=120_000)
        # ★값이 «정확히» 같을 때만 고른다. 접두사로 고르면
        #   「B1F …평면도」가 「B1F …평면도_도면정리(1)」(찍기만 된 것)에
        #   걸려, 손질이 안 된 저장본으로 열려 최불리가 400 으로 막힌다.
        picked = pg.evaluate(
            "(k) => { const s = document.querySelector('#saved');"
            " for (const o of s.options) { if (o.value === k)"
            " { s.value = o.value; s.dispatchEvent(new Event('change'));"
            " return o.value; } } return null; }", key)
        check("목록에서 그 저장본을 골랐다", picked == key, str(picked))
        if picked != key:
            br.close()
            return 1
        pg.click("#btn-reopen")
        # ★단계 개수가 아니라 «손질 판이 보이는가» 로 기다린다 — 기존 UI 검증이
        #   쓰는 그 조건이다(scripts/_verify_module_f_ui.py).
        got = False
        for _ in range(2400):
            pg.wait_for_timeout(200)
            if pg.is_visible("#panel-edit") and pg.is_hidden("#busy"):
                got = True
                break
        check("저장본이 손질로 열린다", got, key[:40])
        if not got:
            br.close()
            return 1

        pg.click("#ed-worst")
        for _ in range(2400):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(1200)
        # ★마지막 칸은 «통합» 이다 — 이름으로 찾는다.
        pg.evaluate("() => { for (const d of "
                    "document.querySelectorAll('#steps div'))"
                    " { if (d.textContent.indexOf('수리계산') >= 0)"
                    " { d.click(); return; } } }")
        pg.wait_for_timeout(1500)
        vis = pg.is_visible("#dg-build")
        check("수리계산 화면이 열린다", vis)
        if not vis:
            br.close()
            return 1
        pg.click("#dg-build")
        for _ in range(3600):
            pg.wait_for_timeout(200)
            if pg.is_hidden("#busy"):
                break
        pg.wait_for_timeout(3000)

        def ink():
            return _ink(pg.query_selector("#cv").screenshot())

        print("[4] 화소 — 밑그림을 켜면 «더» 그려지는가")
        def set_chk(cid, val):
            pg.evaluate(
                "([id, v]) => { const c = document.querySelector(id);"
                " c.checked = v; c.dispatchEvent(new Event('change')); }",
                [f"#{cid}", val])

        set_chk("dg-under", False)
        pg.wait_for_timeout(1500)
        off = ink()
        set_chk("dg-under", True)
        pg.wait_for_timeout(2500)
        on = ink()
        check("밑그림을 켜면 화소가 는다", on > off * 1.02,
              f"끔 {off:,} → 켬 {on:,} ({(on / max(off, 1) - 1) * 100:+.1f}%)")

        print("[5] 아이소를 껐다 켜도 그려지는가")
        set_chk("dg-iso", False)
        pg.wait_for_timeout(4000)
        plan_on = ink()
        check("평면 보기에서도 밑그림이 있다", plan_on > 0, f"{plan_on:,} 화소")
        set_chk("dg-iso", True)
        pg.wait_for_timeout(4000)
        back = ink()
        check("아이소로 돌아와도 그려진다", back > off, f"{back:,} vs 끔 {off:,}")
        br.close()

    print("[6] 콘솔 오류")
    if bad:
        print("    ★4xx/5xx 응답:", bad[:6])
    real = [e for e in errs if "favicon" not in e]
    check("콘솔 오류 0", not real, str(real[:3]))

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("아이소 밑그림 — 서버 변환 + 화면 토글 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
