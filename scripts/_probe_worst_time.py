# -*- coding: utf-8 -*-
"""최불리 선정이 어디서 오래 걸리나 — 단계마다 시계를 댄다.

  짐작하지 않는다. `/edit/worst` 가 실제로 부르는 순서대로 부르고 각각 잰다:

      ① wet_heads        전체망 전개 한 번 (붙는 헤드 판정)
      ② _worst_k_heads   Dijkstra + 순위
      ③ _rank_invariant  같은 후보로 K+1 개를 한 번 더 (먼 순서 검사)
      ④ 제한 전개        최불리 K 개만 다시 전개 (표가 쓰는 것)

    python scripts/_probe_worst_time.py [--key 저장본] [--k 30]
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


class T:
    def __init__(self, label, out):
        self.label, self.out = label, out

    def __enter__(self):
        self.t0 = time.perf_counter()
        return self

    def __exit__(self, *a):
        self.out.append((self.label, time.perf_counter() - self.t0))
        return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")

    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import load_edits, open_board
    from services.cad_import.design.restrict import (attachable_heads,
                                                     restrict_to_worst)
    from services.cad_import.design.worst import worst_k_heads
    from services.cad_import.convert.planar import build_planar_graph

    out: list = []
    with T("판 열기 (open_board + load_edits)", out):
        board = open_board(a.key)
        load_edits(board)
    if not board.sources:
        from services.cad_import.edit.board import body_seg_groups
        from services.cad_import.edit.session import MODE_SOURCE, EditSession
        g = sorted((x[0] for x in body_seg_groups(
            board.pts, board.edges, board.bodies())), key=len, reverse=True)[0]
        es = EditSession(board, key=a.key)
        es.set_mode(MODE_SOURCE)
        pa, pb = g[0]
        es.click((float(pa[0]) + float(pb[0])) / 2.0,
                 (float(pa[1]) + float(pb[1])) / 2.0, 2000)

    payload = dict(board.payload())
    payload["key"] = a.key
    payload["pts"] = [list(p) for p in board.pts]
    payload["edges"] = [list(e) for e in board.edges]
    payload["hcov"] = [list(d) for d in board.disks]
    payload["ups"] = [list(u) for u in board.ups]
    payload["disk_kinds"] = list(board.disk_kinds)
    payload["edge_len_mm"] = dict(getattr(board, "edge_len_mm", None) or {})

    src_index = 0 if board.sources else None
    with T("① wet_heads (전체망 전개)", out):
        probe = attachable_heads(payload, key=a.key)
    with T("② _worst_k_heads (Dijkstra + 순위)", out):
        w = worst_k_heads(board.pts, board.edges, board.hnodes, board.sources,
                          k=a.k, only_heads=None, source_index=src_index,
                          head_xy=board.disks)
    with T("③ _rank_invariant (K+1 한 번 더)", out):
        worst_k_heads(board.pts, board.edges, board.hnodes, board.sources,
                      k=a.k + 1, only_heads=None, source_index=src_index,
                      head_xy=board.disks)
    with T("④ 제한 전개 (최불리 K개만)", out):
        lim = restrict_to_worst(payload, board, w)
        build_planar_graph(a.key, write=False, pts=lim.get("pts"),
                           edges=lim.get("edges"), hcov=lim.get("hcov"),
                           ups=lim.get("ups"),
                           head_kinds=lim.get("head_kinds"),
                           user_sources=lim.get("sources"), ho=lim.get("ho"),
                           edge_len_mm=lim.get("edge_len_mm"))

    print(f"\n■ 최불리 시간 분해 · {a.key} · K={a.k}"
          f" · 노드 {len(board.pts)} · 헤드 {len(board.disks)}")
    tot = sum(t for _l, t in out)
    print()
    for lab, t in out:
        bar = "█" * int(round(40 * t / tot)) if tot else ""
        print(f"  {t:8.2f}s  {t / tot * 100:5.1f}%  {lab:<34} {bar}")
    print(f"  {tot:8.2f}s  100.0%  합계")
    print(f"\n  붙는 헤드 {len(probe.get('wet') or ())}"
          f" / 도면 {probe.get('total')}  ·  선정 {len(w.get('heads') or ())}개")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
