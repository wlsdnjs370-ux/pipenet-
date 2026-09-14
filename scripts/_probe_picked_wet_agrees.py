# -*- coding: utf-8 -*-
"""뽑힌 K개만 전개한 답이 전체망 전개와 **같은가** — 바꾸기 전에 확인한다.

  `/edit/worst` 의 전체망 전개(B1F 154초)를 「뽑힌 K개만 전개」(1.1초)로
  바꿨다. 빨라진 것은 좋지만, **답이 달라지면 아무 소용이 없다.**
  그래서 같은 판에서 둘 다 돌려 뽑힌 K 개에 대한 판정을 맞대 본다.

    python scripts/_probe_picked_wet_agrees.py [--key 저장본] [--k 30]
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DM_KEY = "1. 입력도면 대명동 단위세대 평면도"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import load_edits, open_board
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.restrict import attachable_heads
    from services.cad_import.design.worst import worst_k_heads
    from routes.module_f.attach import picked_heads_wet

    board = open_board(a.key)
    load_edits(board)
    if not board.sources:
        from services.cad_import.edit.board import body_seg_groups
        from services.cad_import.edit.session import MODE_SOURCE
        g = sorted((x[0] for x in body_seg_groups(
            board.pts, board.edges, board.bodies())), key=len, reverse=True)[0]
        e0 = EditSession(board, key=a.key)
        e0.set_mode(MODE_SOURCE)
        pa, pb = g[0]
        e0.click((float(pa[0]) + float(pb[0])) / 2.0,
                 (float(pa[1]) + float(pb[1])) / 2.0, 2000)
    es = EditSession(board, key=a.key)

    w = worst_k_heads(board.pts, board.edges, board.hnodes, board.sources,
                      k=a.k, only_heads=None,
                      source_index=0 if board.sources else None,
                      head_xy=board.disks)
    picked = [int(h) for h in (w.get("heads") or ())]

    t0 = time.perf_counter()
    small = picked_heads_wet(es, picked)
    t_small = time.perf_counter() - t0

    t0 = time.perf_counter()
    full = attachable_heads(es.convert_payload(), key=a.key)
    t_full = time.perf_counter() - t0

    fw = {int(i) for i in (full.get("wet") or ())} & set(picked)
    sw = {int(i) for i in (small.get("wet") or ())}
    print(f"\n■ 뽑힌 K개 판정 대조 · {a.key} · K={a.k}")
    print(f"\n  전체망 전개   {t_full:7.2f}s · 뽑힌 K 중 붙는 것 {len(fw)}개")
    print(f"  K개만 전개    {t_small:7.2f}s · 뽑힌 K 중 붙는 것 {len(sw)}개"
          f"   ({t_full / t_small:.0f}배 빠름)" if t_small else "")
    same = fw == sw
    print(f"\n  답이 같은가: {'예 — 그대로 써도 된다' if same else '★아니오'}")
    if not same:
        print(f"    전체망에만 {sorted(fw - sw)[:10]}")
        print(f"    K개만에만 {sorted(sw - fw)[:10]}")
    return 0 if same else 2


if __name__ == "__main__":
    raise SystemExit(main())
