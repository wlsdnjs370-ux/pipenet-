# -*- coding: utf-8 -*-
"""큰 도면 찬 열기의 **함수 단위** 프로파일 — 54초가 어느 함수에 있나.

`_probe_big_dxf_time.py` 가 「① EditSession.open 이 98%」까지 좁혔다.
여기서는 그 안을 cProfile 로 연다. 누적시간(cumtime)이 아니라 **자기시간
(tottime)** 도 같이 본다 — 누적만 보면 맨 위 호출자가 100% 로 나와 아무것도
안 알려 준다.

    python scripts/_probe_big_dxf_profile.py [--key ...] [--top 30]
"""
from __future__ import annotations

import argparse
import contextlib
import cProfile
import io
import os
import pstats
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
for p in (ROOT, os.path.join(ROOT, "scripts"), os.path.join(ROOT, "tests")):
    if p not in sys.path:
        sys.path.insert(0, p)

BIG = "B1F 현장조사 소화설비 평면도_컨셉2"


def run(key, top) -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import IMPORT_WORK_ROOT, _boot
    _boot()
    cache = os.path.join(str(IMPORT_WORK_ROOT), f"_edit_disp_cache_{key}.json")
    moved = None
    if os.path.isfile(cache):
        moved = cache + ".prof_bak"
        shutil.move(cache, moved)

    buf = io.StringIO()
    pr = cProfile.Profile()
    try:
        from services.cad_import.edit.session import EditSession
        with contextlib.redirect_stdout(buf):
            pr.enable()
            es = EditSession.open(key, out_dir=None, load_saved=True,
                                  use_cache=True)
            pr.disable()
        n = (len(es.board.pts), len(es.board.edges), len(es.board.disks))
    finally:
        if moved and os.path.isfile(moved):
            shutil.move(moved, cache)

    s = io.StringIO()
    st = pstats.Stats(pr, stream=s)
    print(f"■ {key} — 찬 열기 프로파일 · 망 점 {n[0]:,} 간선 {n[1]:,} 헤드 {n[2]:,}")
    print(f"  전체 {st.total_tt:.1f}s\n")

    print("── 자기시간(tottime) 상위 — «이 함수 본문» 이 먹은 시간")
    st.sort_stats("tottime").print_stats(top)
    body = s.getvalue()
    keep = [ln for ln in body.splitlines()
            if ln.strip() and not ln.startswith(("   Ordered", "   List",
                                                 "         ", "Wed", "Thu"))]
    for ln in keep[:top + 4]:
        print("  " + ln)

    s2 = io.StringIO()
    st2 = pstats.Stats(pr, stream=s2)
    print("\n── 우리 코드만(cumtime) — services/routes/core 아래")
    st2.sort_stats("cumtime").print_stats(
        r"(cad_project_editor_g|routes|core)", top)
    for ln in s2.getvalue().splitlines()[:top + 8]:
        if ln.strip():
            print("  " + ln)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=BIG)
    ap.add_argument("--top", type=int, default=25)
    a = ap.parse_args()
    return run(a.key, a.top)


if __name__ == "__main__":
    raise SystemExit(main())
