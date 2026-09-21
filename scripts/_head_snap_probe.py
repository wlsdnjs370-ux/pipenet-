# -*- coding: utf-8 -*-
"""leaf 헤드 x,y 스냅(r30_combined.py 공통 1.5)이 헤드를 얼마나 옮기는지 실측."""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
for p in (BASE, BASE / "core"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))
sys.stdout.reconfigure(encoding="utf-8")

import importlib
import math

srv = importlib.import_module("대조 서버")
app = srv.app
app.before_request_funcs[None] = [
    f for f in app.before_request_funcs.get(None, [])
    if f.__name__ != "_require_login_gate"
]

PLANE = BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"
c = app.test_client()

with PLANE.open("rb") as fh:
    r = c.post("/api/remote30/prototype/run",
               data={"dxf_file": (fh, PLANE.name)},
               content_type="multipart/form-data")
job = r.get_json()["job_id"]
for _ in c.get(f"/api/remote30/prototype/stream/{job}").response:
    pass
c.post(f"/api/remote30/prototype/finalize/{job}")
for _ in c.get(f"/api/remote30/prototype/finalize_stream/{job}").response:
    pass
r = c.post("/api/remote30/system/extract",
           data={"use_legacy_template": "true", "pump_x": "0", "pump_y": "0",
                 "av_x": "0", "av_y": "-3000"})
riser = r.get_json()
assert riser["ok"], riser
r = c.post("/api/remote30/combined/build",
           json={"plane_job_id": job, "system_riser": riser["riser"],
                 "head_orientation": "pendent"})
d = r.get_json()
assert d["ok"], d.get("message")

g = d["geometry"]
node = {n["label"]: n for n in g["nodes"]}
heads = set(g["head_labels"])
nbrs: dict[str, set] = {}
for p in g["pipes"]:
    a, b = str(p["in"]), str(p["out"])
    if a in heads:
        nbrs.setdefault(a, set()).add(b)
    if b in heads:
        nbrs.setdefault(b, set()).add(a)

leaf = {h: next(iter(n)) for h, n in nbrs.items() if len(n) == 1}
print(f"헤드 {len(heads)}개 · leaf 헤드 {len(leaf)}개 (스냅 대상) · "
      f"통과 헤드 {len(heads) - len(leaf)}개")
dists = []
for h, t in leaf.items():
    hn, tn = node.get(h), node.get(t)
    if not hn or not tn:
        continue
    dists.append((math.hypot(hn["x"] - tn["x"], hn["y"] - tn["y"]), h, t))
dists.sort(reverse=True)
print(f"스냅으로 이동하는 평면거리 [도면단위]:")
for dist, h, t in dists[:10]:
    print(f"   {h} → {t} : {dist:,.0f}")
if dists:
    nz = [d for d, _, _ in dists if d > 1e-6]
    med = sorted(d for d, _, _ in dists)[len(dists) // 2]
    print(f"   0 아님 {len(nz)}/{len(dists)}개 · 최대 {dists[0][0]:,.0f} · 중앙 {med:,.0f}")
    xs = [n["x"] for n in g["nodes"]]
    ys = [n["y"] for n in g["nodes"]]
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys))
    drop = float(g.get("head_stub_len") or 0.0)
    print(f"\n평면 대각선 {diag:,.0f} · 니플 길이(head_stub_len) {drop:,.0f}")
    print(f"스냅 제거 시 드롭배관 기울기: 중앙 {math.degrees(math.atan2(med, drop)):.1f}° · "
          f"최대 {math.degrees(math.atan2(dists[0][0], drop)):.1f}° (0°=완전수직)")
