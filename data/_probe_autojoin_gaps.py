# -*- coding: utf-8 -*-
"""자동 이음 — 관 끝 틈의 실제 분포를 잰다. 문턱을 무엇으로 고를지 정하려는 것."""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)

import routes.module_f as mf  # noqa: E402

mf._boot()
sys.path.append(str(mf.EDITOR_ROOT))
from services.cad_import.edit.session import EditSession  # noqa: E402

KEY = sys.argv[1] if len(sys.argv) > 1 else "B1F 현장조사 소화설비 평면도"
es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
print(f"{KEY} · 노드 {len(b.pts)} · 간선 {len(b.edges)} · 덩이 {len(b.bodies())}")

scan = mf._autojoin_scan(b)
print(f"\n관 끝 {scan['ends']} · 가까운 짝 {scan['near']} · 방향 맞는 짝 {scan['kept']}")
print(f"고른 여유 {scan['eps_mm']}mm · 후보 {len(scan['cands'])} · {scan['by_kind']}")

print("\n사다리:")
prev = 0
for t in scan["trials"]:
    gain = t["pairs"] - prev
    print(f"  {t['eps_mm']:6.0f}mm · 짝 {t['pairs']:5d} (+{gain:4d})"
          f" · 덩이 {t['bodies']:4d} · 최대조각 {t['largest']:6d}")
    prev = t["pairs"]

print("\n덩이 크기 상위 12:")
sizes = sorted((len(s) for s in b.bodies()), reverse=True)
print("  " + " ".join(str(s) for s in sizes[:12]) + f"  (총 {len(sizes)}덩이)")

# ── 진짜 성과 지표: 급수원에서 물이 닿는 헤드 수. 덩이 수는 대리지표일 뿐이다.
if b.sources:
    st0 = b.water_state()
    print(f"\n붙이기 전 물 닿는 헤드 {len(st0['wet_heads'])}/{st0['total_heads']}"
          f" · 젖은 간선 {len(st0['wet_edges'])}")
    rep = mf._autojoin_apply(b, scan)
    st1 = b.water_state()
    print(f"붙인 뒤   물 닿는 헤드 {len(st1['wet_heads'])}/{st1['total_heads']}"
          f" · 젖은 간선 {len(st1['wet_edges'])}")
    print(f"→ {rep}")
else:
    print("\n급수원이 없어 물흐름 비교는 건너뜁니다.")
