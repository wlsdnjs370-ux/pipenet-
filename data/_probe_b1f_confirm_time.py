# -*- coding: utf-8 -*-
"""[§21] B1F «표 확정» 이 아직 23분인가 — 재고, 무음 구간이 남았는지 본다.

§21 은 두 가지를 문제 삼았다:
  ① 표 확정 평균 **23분** (1,486 / 1,368 / 1,368 / 1,414 초)
  ② 그 사이 **한 줄도 안 나오는** 무음 구간 (「급수원 Z1 → …펌프」 다음부터
     「노드정리(SSOT)」까지)

둘 다 잰다. 출력 줄마다 시각을 찍어 «가장 긴 침묵» 을 함께 낸다 — 총 시간이
줄어도 침묵이 남아 있으면 ②는 안 풀린 것이다.
"""
from __future__ import annotations

import io
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")

from routes.module_f.common import _boot                           # noqa: E402
_boot()
from services.cad_import.design.anchor import valve_kfp_nodes      # noqa: E402
from services.cad_import.design.restrict import select_and_expand  # noqa: E402
from services.cad_import.design.tables import build_design_tables  # noqa: E402
from services.cad_import.design.worst import worst_k_heads         # noqa: E402
from services.cad_import.edit.session import EditSession           # noqa: E402


class _Stamp(io.TextIOBase):
    """출력 줄마다 시각을 찍어 «침묵» 을 잰다."""

    def __init__(self, real, t0):
        self.real, self.t0 = real, t0
        self.last = t0
        self.gaps: list = []

    def write(self, s):
        if s.strip():
            now = time.perf_counter()
            self.gaps.append((now - self.last, s.strip()[:70]))
            self.last = now
            self.real.write(f"  [{now - self.t0:7.1f}s] {s.rstrip()}\n")
        return len(s)

    def flush(self):
        self.real.flush()


es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
print(f"board — 노드 {len(b.pts):,} · 간선 {len(b.edges):,} · "
      f"헤드 {len(b.disks):,} · 접속점 {list(b.sources)}")

w = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=30)
print(f"최불리 — K {len(w['heads'])} · 최원 {w['far_m']} m")

real, t0 = sys.stdout, time.perf_counter()
tap = _Stamp(real, t0)
sys.stdout = tap
try:
    got = select_and_expand(es.convert_payload(), b, k=30)
    t_exp = time.perf_counter() - t0
    if not got.get("ok"):
        raise SystemExit(got.get("error"))
    av, _m = valve_kfp_nodes(got["kfp"].get("nodes_meta_runtime") or {},
                             b.pts, list(b.valves), got.get("origin_mm"))
    tbl = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                              board_pts=b.pts, valve_nodes=av,
                              tree_loads=got.get("tree_loads"),
                              origin_mm=got.get("origin_mm"))
    t_all = time.perf_counter() - t0
finally:
    sys.stdout = real

print(f"\n■ 표 확정 — 전개 {t_exp:.1f}s · 표까지 합계 {t_all:.1f}s")
print(f"    §21 실측은 평균 1,409초(23분)였다 — "
      f"지금은 {t_all:.1f}초 ({1409 / max(t_all, 1e-9):.0f}배 빠름)")
tap.gaps.sort(reverse=True)
print("\n■ 가장 긴 침묵 5개 (이 값이 크면 ②는 안 풀린 것)")
for g, line in tap.gaps[:5]:
    print(f"    {g:7.1f}s  ← 그 뒤 「{line}」")
print(f"\n    표 — 노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)} · "
      f"노즐 {len(tbl.nozzles)} · 기기 {len(tbl.equipment)}")
