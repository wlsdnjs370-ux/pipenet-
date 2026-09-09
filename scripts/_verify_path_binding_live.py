# -*- coding: utf-8 -*-
"""[경로 상수 함수화] 실서버가 **같은 폴더**에 쓰고 읽는가 — 지시서 §3-3.

계산이 맞는 것과 파일이 실제로 그 자리에 나는 것은 다른 이야기다. 띄워 놓은
서버(5051)에 도면을 올려 한 바퀴 돌리고, 조치 전에 쓰던 그 두 폴더에 파일이
나는지를 본다.

    handoff (찍기→1단계 넘김)  →  docs/import/0단계_새찍기/*_stage1_world.sqlite3
    표시캐시                    →  docs/import/_edit_disp_cache_*.json

★모듈 F 는 «찍은스펙» 을 쓰지 않는다 — 읽기만 한다(쓰는 쪽은 데스크톱 G).
  그래서 쓰기는 위 둘로 보고, 읽기는 «기존 찍은스펙이 되살아나는가» 로 본다.

임시 이름으로 올려 검증 뒤 지운다 — 실작업 폴더를 어지르지 않는다.

    python scripts/_verify_path_binding_live.py
"""
from __future__ import annotations

import io
import os
import shutil
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
WORK = ROOT / "cad_project_editor_g" / "docs" / "import"
PICK = WORK / "0단계_새찍기"
PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
TAG = "_경로검증_임시도면"
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(name)


def _password():
    p = ROOT / ".env"
    if p.is_file():
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def _sweep():
    """검증이 만든 것만 지운다 — 키가 이름에 박혀 있어 골라낼 수 있다."""
    n = 0
    for d in (PICK, WORK):
        for p in d.glob(f"*{TAG}*"):
            try:
                p.unlink()
                n += 1
            except OSError:
                pass
    return n


def main() -> int:
    from playwright.sync_api import sync_playwright
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0

    print(f"\n■ 실서버 경로 검증 · {BASE}")
    print(f"  쓰기 루트 {WORK}")
    print(f"  기존 파일 · 새찍기 {len(list(PICK.glob('*')))}"
          f" · 표시캐시 {len(list(WORK.glob('_edit_disp_cache_*.json')))}")
    _sweep()

    tmp = Path(tempfile.mkdtemp(prefix="mf_pathprobe_"))
    up = tmp / f"{TAG}.dxf"
    shutil.copy2(PLAN, up)

    with sync_playwright() as pw:
        br = pw.chromium.launch()
        pg = br.new_page(viewport={"width": 1500, "height": 950})
        errs: list[str] = []
        pg.on("pageerror", lambda e: errs.append(f"pageerror: {e}"))
        pg.goto(f"{BASE}/login", wait_until="domcontentloaded")
        pg.fill("input[type=password]", _password())
        pg.click("button[type=submit]")
        pg.wait_for_load_state("domcontentloaded")
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(800)

        def idle(n=9000):
            for _ in range(n):
                pg.wait_for_timeout(200)
                if pg.is_hidden("#busy"):
                    return

        # ① 새 이름으로 열기 — handoff 가 새 폴더가 아니라 **그 폴더**에 나야 한다.
        pg.set_input_files("#dxf", str(up))
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        segs = pg.evaluate("() => (window.__mf.view ? 1 : 0)")
        check("도면이 열린다", bool(segs))

        # ② 손질까지 가면 표시캐시가 난다.
        pg.click("#pk-next")
        idle()
        pg.wait_for_timeout(2000)
        real = [e for e in errs if "favicon" not in e]
        check("콘솔 오류 0", not real, str(real[:2]))

        # ③ 읽기 — 기존 저장분(대명동 평면도)이 되살아나는가.
        pg.goto(f"{BASE}/module-f", wait_until="domcontentloaded")
        pg.wait_for_timeout(600)
        pg.set_input_files("#dxf", str(PLAN))
        pg.click("#btn-open")
        idle()
        pg.wait_for_timeout(1500)
        n_mat = pg.evaluate("() => (window.__mf.pick.materials || []).length")
        n_head = pg.evaluate("() => (window.__mf.pick.heads || []).length")
        check("★기존 찍은스펙을 그 폴더에서 읽는다", n_mat > 0 or n_head > 0,
              f"재료 {n_mat} · 헤드 {n_head}")
        br.close()

    made_pick = sorted(p.name for p in PICK.glob(f"*{TAG}*"))
    made_cache = sorted(p.name for p in WORK.glob(f"_edit_disp_cache_*{TAG}*"))
    check("★handoff 가 조치 전과 같은 폴더에 난다", bool(made_pick),
          f"{PICK.name}/ → {made_pick}")
    check("★표시캐시가 조치 전과 같은 폴더에 난다", bool(made_cache),
          f"{WORK.name}/ → {made_cache}")
    print(f"  [정리] 검증 산출물 {_sweep()}건 삭제")
    shutil.rmtree(tmp, ignore_errors=True)

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("실서버 경로가 조치 전과 같다 — 쓰기 두 곳 · 읽기 한 곳")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
