# -*- coding: utf-8 -*-
"""실 편집 기하가 실도면에서 얼마나 통하는지 실측 — 일회용.

브라우저 검증에서 맞닿은 두 실의 합치기와 실 한가운데를 지나는 자르기가 모두
거절됐다. 고른 쌍이 나빴던 것인지, 기하 규칙이 실도면에서 사실상 안 되는 것인지
전체 조합으로 세어 본다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.design import room_edit as RE  # noqa: E402

SESS = ROOT / "data" / "design_sessions"
path = Path(sys.argv[1]) if len(sys.argv) > 1 else max(
    SESS.glob("*/building.json"), key=lambda p: p.stat().st_mtime)
rooms = json.loads(path.read_text(encoding="utf-8"))["rooms"]
rooms = [r for r in rooms if len(r.get("polygon") or []) >= 3]
print(path.parent.name, "— 실", len(rooms), "개")


def edges(room):
    poly = room["polygon"]
    out = set()
    for a, b in RE._ring_edges(poly):
        ka, kb = RE._key(a), RE._key(b)
        if ka != kb:
            out.add((ka, kb) if ka < kb else (kb, ka))
    return out


emap = {r["id"]: edges(r) for r in rooms}
pmap = {r["id"]: r["polygon"] for r in rooms}
ids = [r["id"] for r in rooms]

ok, fail = 0, {}
for i, a in enumerate(ids):
    for b in ids[i + 1:]:
        if not (emap[a] & emap[b]):
            continue
        try:
            RE.merge_polygons([pmap[a], pmap[b]])
            ok += 1
        except RE.RoomEditError as e:
            fail[str(e)[:30]] = fail.get(str(e)[:30], 0) + 1
print(f"\n[합치기] 변을 공유한 쌍 {ok + sum(fail.values())} — 성공 {ok}")
for msg, n in sorted(fail.items(), key=lambda kv: -kv[1]):
    print(f"   실패 {n:4d}  {msg}…")

# 자르기: 실 무게중심을 지나는 가로선·세로선. 사람이 가장 흔하게 긋는 두 선이다.
res = {"성공": 0}
for r in rooms:
    poly = pmap[r["id"]]
    cx = sum(p[0] for p in poly) / len(poly)
    cy = sum(p[1] for p in poly) / len(poly)
    for p1, p2 in (((cx - 1e6, cy), (cx + 1e6, cy)),
                   ((cx, cy - 1e6), (cx, cy + 1e6))):
        try:
            RE.split_polygon(poly, p1, p2)
            res["성공"] += 1
        except RE.RoomEditError as e:
            k = str(e)[:24]
            res[k] = res.get(k, 0) + 1
print(f"\n[자르기] 무게중심 가로·세로 {len(rooms) * 2} 회")
for msg, n in sorted(res.items(), key=lambda kv: -kv[1]):
    print(f"   {n:4d}  {msg}…")

# 변은 공유하지 않지만 붙어 있는 쌍이 얼마나 되는지 — 면 추출이 이웃 실의 꼭짓점을
# 내 변 위에 심어 두면(T 접합) 같은 변으로 세어지지 않는다.
nmap = {r["id"]: {k for e in emap[r["id"]] for k in e} for r in rooms}
share_edge = share_node = 0
for i, a in enumerate(ids):
    for b in ids[i + 1:]:
        if emap[a] & emap[b]:
            share_edge += 1
        elif len(nmap[a] & nmap[b]) >= 2:
            share_node += 1
print(f"\n[인접] 변 공유 {share_edge} 쌍 / 변은 다르나 꼭짓점 2개 이상 공유"
      f" {share_node} 쌍")

# 볼록·오목 분포 — 자르기 거절이 실 모양 탓인지 본다.
concave = 0
for r in rooms:
    poly = pmap[r["id"]]
    n = len(poly)
    signs = set()
    for i in range(n):
        ax, ay = poly[i]
        bx, by = poly[(i + 1) % n]
        cx2, cy2 = poly[(i + 2) % n]
        cross = (bx - ax) * (cy2 - by) - (by - ay) * (cx2 - bx)
        if abs(cross) > 1e-6:
            signs.add(cross > 0)
    if len(signs) > 1:
        concave += 1
print(f"\n오목한 실 {concave} / {len(rooms)}"
      f" — 꼭짓점 중앙값 {sorted(len(pmap[i]) for i in ids)[len(ids) // 2]}")
