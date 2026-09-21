# -*- coding: utf-8 -*-
"""도면 범례 찾기 — 기호 옆 텍스트로 r=180 ARC 쌍의 정체를 확인."""
from __future__ import annotations

import math
import sys
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "data/uploads/B1F_.dxf"
bundle = R.parse_dxf_bundle_cached(DXF)

texts = [(en["p"][0], en["p"][1], en.get("v", ""), en["l"])
         for en in bundle.entities if en["t"] == "T" and len(en.get("p", [])) >= 2]
arcs180 = [(en["c"][0], en["c"][1], en["l"]) for en in bundle.entities
           if en["t"] == "A" and abs(float(en.get("r", 0) or 0) - 180) < 1]
print(f"텍스트 {len(texts)} · r≈180 ARC {len(arcs180)}")

KEY = ("범례", "기호", "티", "엘보", "레듀", "레듀샤", "교차", "관통", "분기",
       "밸브", "체크", "후렉", "플렉", "신축", "슬리브", "행거", "지지", "LEGEND",
       "TEE", "ELBOW", "CROSS", "입상", "입하")
hits = [t for t in texts if any(k in t[2] for k in KEY)]
print(f"\n=== 키워드 텍스트 {len(hits)}건 (앞 60) ===")
for x, y, v, ly in hits[:60]:
    print(f"  ({x:10.0f},{y:10.0f}) {ly:22s} {v[:52]!r}")

# 텍스트가 밀집한 곳 = 범례 후보
cell = 5000.0
dens = Counter()
for x, y, v, ly in texts:
    dens[(int(x // cell), int(y // cell))] += 1
print("\n=== 텍스트 밀집 상위 8 셀 (5m 격자) ===")
for (cx, cy), n in dens.most_common(8):
    print(f"  cell({cx*5:6.0f}m,{cy*5:6.0f}m) 텍스트 {n}")

# 각 r180 ARC 근처 텍스트
print("\n=== r180 ARC 최근접 텍스트 (앞 12개 ARC) ===")
for ax, ay, aly in arcs180[:12]:
    best = min(((math.hypot(x-ax, y-ay), v, ly) for x, y, v, ly in texts),
               default=(1e9, "", ""))
    print(f"  ARC({ax:10.0f},{ay:10.0f}) {aly:16s} → {best[0]:8.1f}mm {best[2]:16s} {best[1][:40]!r}")

# 범례 영역 추정: ARC 가 배관 없이 고립된 곳
print("\n=== 짧은 텍스트 빈도 상위 40 ===")
tc = Counter(v.strip() for _x, _y, v, _l in texts if 0 < len(v.strip()) <= 14)
for v, n in tc.most_common(40):
    print(f"  {n:4d}  {v!r}")
