# -*- coding: utf-8 -*-
"""기준 헤드(최원단)를 kfp 노드로 되짚을 수 있는가 — 실측.

`design/tables._worst_head_node()` 는 무조건 None 을 돌려주는 스텁이다. 그 탓에
표의 「기준 헤드 노드」가 항상 '?' 이고, 최원 유하거리 «경로» 를 그리는 블록이
통째로 죽어 있다(BLOCKED §30).

포기 사유는 「board 헤드 번호와 전개 노드가 1:1 이 아니다」였다. 정말 그런지,
그렇다면 무엇으로 이을 수 있는지를 실제 도면에서 잰다.
"""
from __future__ import annotations

import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "1. 입력도면 대명동 단위세대 평면도")

from routes.module_f.common import _boot                            # noqa: E402
_boot()
from services.cad_import.edit.session import EditSession, MODE_VALVE  # noqa: E402
from services.cad_import.design.restrict import select_and_expand    # noqa: E402
from services.cad_import.design.worst import worst_k_heads           # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
print(f"board — 노드 {len(b.pts)} · 간선 {len(b.edges)} · 헤드 {len(b.disks)}"
      f" · 알람밸브 {list(b.valves)} · 접속점 {list(b.sources)}")

hn = b._head_nodes()
if not b.sources:
    # 아무 배관이나 찍으면 헤드망과 안 이어진 조각에 걸린다 — 헤드에 «닿는»
    # 자리를 찾을 때까지 찍어 본다(사람이 화면에서 하는 일과 같다).
    es.set_mode(MODE_VALVE)
    tried = 0
    for (u, v) in sorted(b.edges):
        mx = (b.pts[u][0] + b.pts[v][0]) / 2
        my = (b.pts[u][1] + b.pts[v][1]) / 2
        if not es.click(mx, my, 500.0):
            continue
        tried += 1
        probe = worst_k_heads(b.pts, b.edges, hn, b.sources, k=30)
        if probe.get("worst_head") is not None:
            print(f"  → 알람밸브를 찍었다({tried}번째 시도) · "
                  f"접속점 {list(b.sources)} · 닿는 헤드 {probe['reachable']}")
            break
        es.click(mx, my, 500.0)          # 토글로 되돌린다
        if tried >= 60:
            raise SystemExit("헤드에 닿는 자리를 못 찾았다")
    hn = b._head_nodes()

w = worst_k_heads(b.pts, b.edges, hn, b.sources, k=30)
wh = w["worst_head"]
path = w["worst_path"]
print(f"\n최불리 — 기준 헤드(디스크 번호) {wh} · 경로 절점 {len(path)}"
      f" · {w['worst_path_m']} m · 최원 {w['far_m']} m")
if wh is None:
    raise SystemExit("기준 헤드가 없다 — 이 도면으로는 못 잰다")

disk = b.disks[wh]
board_node = path[-1] if path else None
print(f"  디스크 중심 (board mm) = ({disk[0]:.1f}, {disk[1]:.1f}) r={disk[2]:.1f}")
print(f"  경로 끝 board 노드 = {board_node} "
      f"{b.pts[board_node] if board_node is not None else ''}")

payload = es.convert_payload()
got = select_and_expand(payload, b, k=30)
if not got.get("ok"):
    raise SystemExit(f"전개 실패: {got.get('error')}")
kfp = got["kfp"]
nodes = kfp["nodes_meta_runtime"]
origin = got.get("origin_mm")
node_ref = got.get("node_ref") or {}
print(f"\n전개 — kfp 노드 {len(nodes)} · origin_mm {origin} · "
      f"node_ref {len(node_ref)}")

heads = {nid: n for nid, n in nodes.items()
         if str(n.get("type_id")) == "head"}
print(f"  kfp 헤드 노드 {len(heads)}개")


def to_kfp(mx, my):
    """board mm → kfp m — main_walk.xf_mm_to_m 과 같은 규칙."""
    return ((mx - origin[0]) / 1000.0 + 1.0, (my - origin[1]) / 1000.0 + 1.0)


# ── ① node_ref 로 되짚기 (세로 전개 전 노드에만 있다)
rev = {}
for nid, vid in node_ref.items():
    rev.setdefault(vid, nid)
print(f"\n① node_ref 역방향 — board 노드 {board_node} → "
      f"kfp {rev.get(board_node, '없음')}")
hit = rev.get(board_node)
if hit is not None:
    print(f"   그 노드 type_id = {nodes.get(hit, {}).get('type_id')}"
          f" · coords {[round(float(c), 3) for c in (nodes[hit] or {}).get('coords') or []]}")

# ── ② 좌표로 맞대기 (세로 전개가 x,y 를 옮기는가)
tx, ty = to_kfp(disk[0], disk[1])
print(f"\n② 디스크 중심을 kfp 좌표로 = ({tx:.3f}, {ty:.3f})")
near = []
for nid, n in heads.items():
    c = n.get("coords") or [0, 0, 0]
    d = math.hypot(float(c[0]) - tx, float(c[1]) - ty)
    near.append((d, nid, float(c[2])))
near.sort()
for d, nid, z in near[:4]:
    print(f"   {nid}  거리 {d * 1000:.0f} mm · z {z:.3f}")

TOL = 0.30      # m — 전개가 x,y 를 옮기는 폭(상하향식 combo_2)보다 넉넉히
within = [t for t in near if t[0] <= TOL]
print(f"\n   허용 {TOL * 1000:.0f}mm 안 헤드 노드 {len(within)}개 "
      f"→ {'유일' if len(within) == 1 else '여러 개(어느 것이 그 헤드인지 모호)'}")

# ── ③ «닿는 헤드 전부» 에 대해 node_ref 되짚기가 얼마나 되는가
print("\n③ 최불리에 뽑힌 헤드마다 되짚기 성적")
picked = w["heads"]
hn_all = b._head_nodes()
ok_ref = miss_ref = stale = 0
ok_kind = other_kind = 0
dists = []
for hi in picked:
    reach = [n for n in hn_all[hi] if n < len(b.pts)]
    nid = None
    for n in reach:
        if n in rev:
            nid = rev[n]
            break
    if nid is None:
        miss_ref += 1
        continue
    if nid not in nodes:
        stale += 1
        continue
    ok_ref += 1
    if str(nodes.get(nid, {}).get("type_id")) == "head":
        ok_kind += 1
    else:
        other_kind += 1
    # board 노드 좌표와 kfp 노드 좌표가 얼마나 어긋나나
    bn = next(n for n in reach if n in rev)
    ex, ey = to_kfp(b.pts[bn][0], b.pts[bn][1])
    c = nodes[nid].get("coords") or [0, 0, 0]
    dists.append(math.hypot(float(c[0]) - ex, float(c[1]) - ey) * 1000.0)
print(f"   node_ref 로 찾음 {ok_ref} · 못 찾음 {miss_ref} · "
      f"★가리키는 노드가 최종 kfp 에 없음 {stale} (뽑힌 헤드 {len(picked)})")
print(f"   그중 type_id=head {ok_kind} · 다른 종류 {other_kind}")
if dists:
    dists.sort()
    print(f"   board 노드 ↔ kfp 노드 좌표차: 최소 {dists[0]:.0f} · "
          f"중앙 {dists[len(dists) // 2]:.0f} · 최대 {dists[-1]:.0f} mm")

# ── ③-2 좌표 매칭은 얼마나 안정적인가 (board 헤드부착노드 → kfp 헤드노드)
print("\n③-2 좌표 매칭 — 뽑힌 헤드 전수")
d1s, ratios, amb = [], [], 0
for hi in picked:
    reach = [n for n in hn_all[hi] if n < len(b.pts)]
    if not reach:
        continue
    # worst.py 와 같은 «부착 노드» 를 쓴다 — 그중 아무거나가 아니라, 헤드 원
    # 안의 노드 전부에 대해 가장 가까운 kfp 헤드 노드를 본다.
    best = None
    for n in reach:
        ex, ey = to_kfp(b.pts[n][0], b.pts[n][1])
        ds2 = sorted(math.hypot(float((m.get("coords") or [0, 0])[0]) - ex,
                                float((m.get("coords") or [0, 0])[1]) - ey)
                     for m in heads.values())
        if ds2 and (best is None or ds2[0] < best[0]):
            best = (ds2[0], ds2[1] if len(ds2) > 1 else float("inf"))
    if best is None:
        continue
    d1s.append(best[0] * 1000.0)
    ratios.append(best[1] / best[0] if best[0] > 1e-9 else float("inf"))
    if best[1] < 0.5:                 # 500mm 안에 둘 이상이면 모호
        amb += 1
d1s.sort()
if d1s:
    print(f"   최근접 거리(mm): 최소 {d1s[0]:.0f} · 중앙 {d1s[len(d1s)//2]:.0f} · "
          f"최대 {d1s[-1]:.0f}  (n={len(d1s)})")
    print(f"   500mm 안에 둘 이상 = {amb}건  ← 0 이어야 «유일» 로 이을 수 있다")
    rs = sorted(r for r in ratios if r != float("inf"))
    if rs:
        print(f"   2등/1등 거리비: 최소 {rs[0]:.1f}배  ← 클수록 안전")

# ── ④ 헤드 종류별로 전개 모양이 다른가
kinds = {}
for rec in (got.get("head_kinds") or ()):
    kinds[str(rec.get("kind"))] = kinds.get(str(rec.get("kind")), 0) + 1
print(f"\n④ 이 도면의 헤드 종류 분포 = {kinds}")
