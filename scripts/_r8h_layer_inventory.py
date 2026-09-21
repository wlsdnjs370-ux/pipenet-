# -*- coding: utf-8 -*-
"""레이어 전수 조사 — 숨김(frozen/off) 여부 + 엔티티 종류 분포.

파서가 hidden 레이어를 통째로 버리므로, 부속 기호가 거기 숨어 있는지 본다.
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

import ezdxf

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "data/uploads/B1F_.dxf"

doc = ezdxf.readfile(str(DXF))
msp = doc.modelspace()

state = {}
for ly in doc.layers:
    state[str(ly.dxf.name)] = ("OFF" if ly.is_off() else "") + ("FROZEN" if ly.is_frozen() else "")

types: dict = defaultdict(Counter)
depth_types: dict = defaultdict(Counter)


def walk(e, depth=0, override=None):
    own = getattr(e.dxf, "layer", "")
    layer = override if (override is not None and own in ("0", "")) else (own or override or "")
    types[layer][e.dxftype()] += 1
    depth_types[depth][e.dxftype()] += 1
    if e.dxftype() == "INSERT" and depth < 8:
        try:
            blk = doc.blocks[e.dxf.name]
        except Exception:
            return
        for sub in blk:
            walk(sub, depth + 1, layer)


for e in msp:
    walk(e)

print(f"레이어 {len(types)} · 블록 정의 {len(doc.blocks)}")
print("\n=== 레이어별 (엔티티 총수 상위 30) ===")
rows = sorted(types.items(), key=lambda kv: -sum(kv[1].values()))
for name, c in rows[:30]:
    st = state.get(name, "")
    top = ", ".join(f"{t}×{n}" for t, n in c.most_common(5))
    print(f"  {'[' + st + ']' if st else '     ':10s} {name:34s} {sum(c.values()):6d}  {top}")

hidden = [n for n, s in state.items() if s]
print(f"\n=== 숨김 레이어 {len(hidden)}개 ===")
for n in hidden:
    c = types.get(n, Counter())
    print(f"  [{state[n]:6s}] {n:34s} {sum(c.values()):6d}  "
          + ", ".join(f"{t}×{v}" for t, v in c.most_common(5)))

print("\n=== 중첩 깊이별 엔티티 종류 ===")
for d in sorted(depth_types):
    c = depth_types[d]
    print(f"  depth {d}: {sum(c.values()):7d}  "
          + ", ".join(f"{t}×{n}" for t, n in c.most_common(6)))

print("\n=== INSERT 블록 이름 상위 20 (전 깊이) ===")
names = Counter()


def walk_names(e, depth=0):
    if e.dxftype() == "INSERT":
        names[(depth, str(e.dxf.name))] += 1
        if depth < 8:
            try:
                blk = doc.blocks[e.dxf.name]
            except Exception:
                return
            for sub in blk:
                walk_names(sub, depth + 1)


for e in msp:
    walk_names(e)
for (d, n), v in names.most_common(20):
    print(f"  depth{d} {n:40s} ×{v}")
