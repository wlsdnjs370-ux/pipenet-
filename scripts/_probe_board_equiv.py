# -*- coding: utf-8 -*-
"""[속도 최적화] «같은 망이 나오는가» — HEAD 코드와 현재 코드를 나란히 돌린다.

격자 색인은 **후보를 줄이는** 변경이라 원리상 같은 답이 나와야 한다. 그러나
원리는 증명이 아니다. 여기서는 바꾼 파일을 잠시 HEAD 판으로 되돌려 같은
도면의 망을 만들고, 지금 판의 망과 **점·간선·헤드·젖은헤드까지** 대조한다.

    python scripts/_probe_board_equiv.py [--key ...]

★파일을 잠깐 바꾸므로 끝나면 반드시 되돌린다(finally). git stash 는 쓰지
  않는다 — 이 저장소에서 그것으로 수정본을 잃은 적이 있다.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

BIG = "B1F 현장조사 소화설비 평면도_컨셉2"
TOUCHED = [
    "cad_project_editor_g/services/cad_import/pipeline/flow.py",
    "cad_project_editor_g/services/cad_import/pipeline/stage45.py",
]

CHILD = r'''
import contextlib, hashlib, io, json, os, sys
sys.stdout.reconfigure(errors="replace")
ROOT = os.getcwd()
for p in (ROOT, os.path.join(ROOT, "scripts"), os.path.join(ROOT, "tests")):
    if p not in sys.path:
        sys.path.insert(0, p)
os.environ.setdefault("LOGIN_PASSWORD", "probe")
key = sys.argv[1]
from routes.module_f.common import _boot
_boot()
buf = io.StringIO()
import time
with contextlib.redirect_stdout(buf):
    from services.cad_import.edit.session import EditSession
    t0 = time.perf_counter()
    es = EditSession.open(key, out_dir=None, load_saved=True, use_cache=True)
    dt = time.perf_counter() - t0
b = es.board
def h(obj):
    return hashlib.sha256(
        json.dumps(obj, sort_keys=True, default=float).encode()).hexdigest()[:16]
out = {
    "secs": round(dt, 2),
    "n_pts": len(b.pts), "n_edges": len(b.edges), "n_disks": len(b.disks),
    "pts": h([[round(float(x), 4), round(float(y), 4)] for x, y in b.pts]),
    "edges": h(sorted(tuple(sorted(e)) for e in b.edges)),
    "disks": h([[round(float(v), 4) for v in d] for d in b.disks]),
    "kinds": h(list(getattr(b, "disk_kinds", []) or [])),
}
print("RESULT=" + json.dumps(out, ensure_ascii=False))
'''


def child(key):
    r = subprocess.run([sys.executable, "-c", CHILD, key], cwd=ROOT,
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=3600)
    for ln in (r.stdout or "").splitlines():
        if ln.startswith("RESULT="):
            return json.loads(ln[len("RESULT="):])
    raise SystemExit("자식이 결과를 못 냈다:\n"
                     + (r.stderr or r.stdout or "")[-900:])


def cold(key):
    """표시 캐시를 치워 찬 경로로 만든다. (되돌릴 경로, ) 를 돌려준다."""
    from routes.module_f.common import IMPORT_WORK_ROOT, _boot
    _boot()
    p = os.path.join(str(IMPORT_WORK_ROOT), f"_edit_disp_cache_{key}.json")
    if os.path.isfile(p):
        bak = p + ".equiv_bak"
        shutil.move(p, bak)
        return p, bak
    return p, None


def run(key) -> int:
    sys.path.insert(0, ROOT)
    tmp = tempfile.mkdtemp(prefix="equiv_")
    saved = {}
    live_p, live_bak = cold(key)
    try:
        print(f"■ 망 동치 — {key}")
        print("  [지금 판] 찬 열기…")
        now = child(key)
        print(f"    {now['secs']}s · 점 {now['n_pts']:,} · 간선 "
              f"{now['n_edges']:,} · 헤드 {now['n_disks']:,}")

        # HEAD 판으로 잠시 되돌린다
        for f in TOUCHED:
            saved[f] = os.path.join(tmp, os.path.basename(f))
            shutil.copy2(f, saved[f])
            blob = subprocess.run(["git", "show", f"HEAD:{f}"], cwd=ROOT,
                                  capture_output=True)
            if blob.returncode != 0:
                raise SystemExit(f"git show 실패: {f}")
            with open(f, "wb") as fh:
                fh.write(blob.stdout)
        # 캐시를 다시 치운다(지금 판이 만들어 놨다)
        if os.path.isfile(live_p):
            os.remove(live_p)
        print("  [HEAD 판] 찬 열기…")
        old = child(key)
        print(f"    {old['secs']}s · 점 {old['n_pts']:,} · 간선 "
              f"{old['n_edges']:,} · 헤드 {old['n_disks']:,}")
    finally:
        for f, bak in saved.items():
            shutil.copy2(bak, f)
        shutil.rmtree(tmp, ignore_errors=True)
        if os.path.isfile(live_p):
            os.remove(live_p)
        if live_bak and os.path.isfile(live_bak):
            shutil.move(live_bak, live_p)
        print("  (바꾼 파일과 표시 캐시를 되돌려 놓음)")

    print()
    same = True
    for k in ("n_pts", "n_edges", "n_disks", "pts", "edges", "disks", "kinds"):
        ok = old[k] == now[k]
        same = same and ok
        print(f"  [{'OK  ' if ok else 'FAIL'}] {k} · HEAD {old[k]} / 지금 {now[k]}")
    gain = (1 - now["secs"] / max(old["secs"], 1e-9)) * 100
    print(f"\n  찬 열기 {old['secs']}s → {now['secs']}s  ({gain:.0f}% 감소)")
    if not same:
        print("\n  ★★망이 달라졌다 — 이것은 최적화가 아니다")
        return 3
    print("  ★망은 한 글자도 안 달라졌다")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=BIG)
    a = ap.parse_args()
    return run(a.key)


if __name__ == "__main__":
    raise SystemExit(main())
