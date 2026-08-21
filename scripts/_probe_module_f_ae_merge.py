# -*- coding: utf-8 -*-
"""모듈 F 개조 타당성 확인 — 모듈 A 의 장점 3종이 E 그래프 위에서 되는지.

UI 를 붙이기 전에 (1) 레이어 자동 추천 (2) 최불리 30헤드 (3) PIPENET SDF 출력
세 가지가 실도면에서 실제 값을 내는지 먼저 본다.
"""
from __future__ import annotations

import heapq
import math
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)

EDITOR = os.path.join(ROOT, "cad_project_editor")
if EDITOR not in sys.path:
    sys.path.append(EDITOR)

WORK = os.path.join(EDITOR, "docs", "import")
KEY = "B1F 현장조사 소화설비 평면도"


def boot():
    from services.cad_import.pipeline import disp_cache, handoff
    handoff.import_write_root = lambda: WORK
    handoff.OUT_DIR = handoff.pick_out_dir()
    disp_cache._DISP_CACHE_DIR = WORK


# ── (2) 최불리 K 헤드 — 급수원 Dijkstra, E 그래프 자료구조 위에서 ──────
def worst_k_heads(pts, edges, hnodes, sources, k=30):
    """급수원에서 가장 먼 헤드 K개와 그 최단경로 간선.

    모듈 A `select_worst30_heads` 의 규칙(급수원 기점 최원 경로)을 E 의
    pts/edges/hnodes 위에서 그대로 적용한다. 길이는 도면 실측 유클리드.
    """
    adj = {}
    for a, b in edges:
        d = math.dist(pts[a], pts[b])
        adj.setdefault(a, []).append((b, d))
        adj.setdefault(b, []).append((a, d))

    INF = float("inf")
    dist = {}
    prev = {}
    pq = []
    for s in sources:
        dist[s] = 0.0
        heapq.heappush(pq, (0.0, s))
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, INF):
            continue
        for v, w in adj.get(u, ()):
            nd = d + w
            if nd < dist.get(v, INF):
                dist[v] = nd
                prev[v] = u
                heapq.heappush(pq, (nd, v))

    scored = []
    for hi, nodes in enumerate(hnodes):
        best = min((dist[n] for n in nodes if n in dist), default=None)
        if best is None:
            continue
        node = min((n for n in nodes if n in dist), key=lambda n: dist[n])
        scored.append((best, hi, node))
    scored.sort(reverse=True)
    picked = scored[:k]

    keep_edges = set()
    keep_nodes = set()
    for _d, _hi, node in picked:
        cur = node
        keep_nodes.add(cur)
        while cur in prev:
            nxt = prev[cur]
            keep_edges.add((min(cur, nxt), max(cur, nxt)))
            keep_nodes.add(nxt)
            cur = nxt
    return {
        "picked": [(hi, d) for d, hi, _n in picked],
        "edges": keep_edges,
        "nodes": keep_nodes,
        "reachable_heads": len(scored),
        "far": picked[0][0] if picked else 0.0,
        "near": picked[-1][0] if picked else 0.0,
    }


def main():
    boot()
    from services.cad_import.edit.session import EditSession

    print("[A] 레이어 자동 추천 — 모듈 A 분류기를 E 세계에 적용")
    from remote30_prototype import _categorize_layer
    from services.cad_import.pick.session import PickSession
    spec_dxf = None
    import json
    sp = os.path.join(WORK, "0단계_새찍기", f"{KEY}_찍은스펙.json")
    with open(sp, encoding="utf-8") as f:
        spec_dxf = json.load(f).get("source_dxf")
    if spec_dxf and os.path.isfile(spec_dxf):
        t0 = time.perf_counter()
        ps = PickSession.open(spec_dxf)
        cats = {}
        for ly, c, _a, _b in ps.world.segs:
            cats.setdefault(_categorize_layer(ly), set()).add((ly, c))
        print(f"    파싱 {time.perf_counter()-t0:.1f}s · 묶음 {sum(len(v) for v in cats.values())}종")
        for cat in ("PIPE", "HEAD", "ALARM", "ARCH", "TEXT", "EXCLUDE", "OTHER"):
            got = cats.get(cat) or set()
            if got:
                sample = sorted({ly for ly, _ in got})[:3]
                print(f"    {cat:8s} {len(got):3d}묶음  {sample}")
        picked = sorted({ly for ly, _ in (cats.get("PIPE") or set())})
        print(f"    → 배관 추천 레이어 {len(picked)}개")
    else:
        print("    원본 DXF 없음 — 건너뜀")

    print("\n[B] 최불리 K 헤드 — E 그래프 위 Dijkstra")
    es = EditSession.open(KEY, load_saved=True, use_cache=True)
    b = es.board
    print(f"    노드 {len(b.pts)} · 간선 {len(b.edges)} · 헤드 {len(b.disks)}"
          f" · 급수원 {b.sources}")
    if not b.sources:
        print("    급수원이 없어 건너뜀")
        return
    t0 = time.perf_counter()
    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30)
    print(f"    {time.perf_counter()-t0:.2f}s · 도달 헤드 {w['reachable_heads']}개 중 30개 선정")
    print(f"    최원 {w['far']/1000:.1f} m · 30번째 {w['near']/1000:.1f} m"
          f" · 경로 간선 {len(w['edges'])} · 노드 {len(w['nodes'])}")

    print("\n[C] PIPENET SDF — 변환된 .kfp 를 SDF 로")
    from services.cad_import.convert.engine import convert_to_kfp, ensure_planar
    payload = ensure_planar(es.convert_payload())
    res = convert_to_kfp(payload, None)
    if not res["ok"]:
        print("    변환 막힘:", res["blockers"])
        return
    kfp = res["kfp"]
    print(f"    .kfp 노드 {len(kfp['nodes_meta_runtime'])} · 배관 {len(kfp['pipe_data'])}")
    t0 = time.perf_counter()
    from kfp_sdf_converter import emit_sdf_xml, kfp_dict_to_network
    net = kfp_dict_to_network(kfp)
    xml = emit_sdf_xml(net)
    print(f"    {time.perf_counter()-t0:.2f}s · SDF {len(xml):,} 자")
    print("    머리:", xml[:120].replace("\n", " "))
    for tag in ("<Node", "<Pipe", "<Nozzle", "<Equipment"):
        print(f"      {tag:12s} {xml.count(tag)}")
    try:
        from kfp_sdf_converter import _resolve_standard_slf
        slf = _resolve_standard_slf()
        print("    표준 SLF:", slf, "존재" if slf and os.path.isfile(slf) else "없음")
    except Exception as exc:  # noqa: BLE001
        print("    표준 SLF 확인 실패:", exc)


if __name__ == "__main__":
    main()
