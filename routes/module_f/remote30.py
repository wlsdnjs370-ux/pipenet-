# -*- coding: utf-8 -*-
"""모듈 A 에서 빌려온 것 — 최불리 K · 도면 장 나누기 · 범위 제한 · PIPENET."""
from __future__ import annotations

import heapq
import math
import os

from routes.module_f.common import REMOTE_K_DEFAULT, _r1

def _worst_k_heads(pts, edges, hnodes, sources, k=REMOTE_K_DEFAULT,
                   only_heads=None) -> dict:
    """앵커 기반 «최불리 배관망» 추출 — 수리계산의 설계면적 그 자체.

    ─ 왜 «먼 순서 K개» 가 아니라 앵커인가 ────────────────────────────
    NFPC 103 의 기준개수(K)는 «하나의 설계구역 안에서 동시에 방수되는 인접
    K개» 다. 급수원에서 먼 순서로 그냥 K개를 뽑으면 도면 곳곳의 막다른 헤드가
    섞여 뽑힌다 — B1F 실측: 먼 순서 30개는 대각 95.9m 로 흩어졌고, 앵커 방식은
    30.3m 로 한 구역에 뭉쳤다. 흩어진 30개로는 설계면적이 성립하지 않는다.

    ─ 세 단계 ────────────────────────────────────────────────────────
    ① 앵커 = 급수원에서 **배관 거리로** 가장 먼(가장 불리한) 헤드. 여기가
       기준압을 잡는 지점 — 급수원↔앵커 거리가 «최원 유하거리» 다.
    ② 설계면적 = 앵커에서 **배관 거리로** 가까운 K개(유클리드 아님 — 실제 물이
       같은 관을 타고 함께 흐르는 무리라야 한다).
    ③ corridor = 그 K개를 급수원까지 잇는 최단경로의 합집합. 각 간선의
       **담당 헤드 수(load)** 를 함께 낸다 — NFPC 별표1 이 최소 호칭경을 정할
       때 쓰는 바로 그 값이라, 이 최대값이 주배관 관경을 결정한다.

    `only_heads` : 도면이 여러 장일 때 한 장으로 범위를 좁힌다. 앵커도 그
        범위 안에서 고른다(장이 다르면 앵커가 남의 도면으로 튄다).
    """
    adj: dict[int, list[int]] = {}
    for a, b in edges:
        adj.setdefault(a, []).append(b)
        adj.setdefault(b, []).append(a)

    def dijkstra(seeds):
        INF = float("inf")
        dist: dict[int, float] = {}
        prev: dict[int, int] = {}
        pq: list[tuple[float, int]] = []
        for s in seeds:
            if isinstance(s, int) and 0 <= s < len(pts):
                dist[s] = 0.0
                heapq.heappush(pq, (0.0, s))
        while pq:
            d, u = heapq.heappop(pq)
            if d > dist.get(u, INF):
                continue
            for v in adj.get(u, ()):
                nd = d + math.dist(pts[u], pts[v])
                if nd < dist.get(v, INF):
                    dist[v] = nd
                    prev[v] = u
                    heapq.heappush(pq, (nd, v))
        return dist, prev

    # ① 급수원 기점 — 헤드마다 부착 노드·유하거리
    src_dist, prev = dijkstra(list(sources))
    head_node: dict[int, int] = {}
    head_far: dict[int, float] = {}
    for hi, nodes in enumerate(hnodes):
        if only_heads is not None and hi not in only_heads:
            continue
        reach = [n for n in nodes if n in src_dist]
        if not reach:
            continue
        node = min(reach, key=lambda n: src_dist[n])
        head_node[hi] = node
        head_far[hi] = src_dist[node]

    reachable = len(head_far)
    empty = {"heads": [], "anchor": None, "edges": set(), "nodes": set(),
             "loads": {}, "reachable": reachable, "unreachable": 0,
             "far_m": 0.0, "near_m": 0.0, "span_m": 0.0, "total_m": 0.0,
             "max_load": 0}
    if not head_far:
        return empty

    k = max(1, min(int(k), reachable))
    anchor = max(head_far, key=head_far.get)   # 가장 불리한 헤드

    # ② 앵커 기점 — 배관 거리로 가까운 K개 = 설계면적
    an_dist, _ = dijkstra([head_node[anchor]])
    ranked = sorted(head_node,
                    key=lambda hi: an_dist.get(head_node[hi], float("inf")))
    picked = ranked[:k]
    span = an_dist.get(head_node[picked[-1]], 0.0) if picked else 0.0

    # ③ corridor — K개 → 급수원 경로 합집합 + 담당 헤드 수
    loads: dict[tuple[int, int], int] = {}
    keep_nodes: set[int] = set()
    for hi in picked:
        cur = head_node[hi]
        keep_nodes.add(cur)
        while cur in prev:
            nxt = prev[cur]
            key = (min(cur, nxt), max(cur, nxt))
            loads[key] = loads.get(key, 0) + 1
            keep_nodes.add(nxt)
            cur = nxt

    total = sum(math.dist(pts[a], pts[b]) for a, b in loads)
    return {
        "heads": picked,
        "anchor": anchor,
        "dists": {hi: head_far[hi] for hi in picked},
        "edges": set(loads),
        "loads": loads,
        "nodes": keep_nodes,
        "reachable": reachable,
        "unreachable": 0,          # picked 는 전부 도달 헤드 중에서 골랐다
        "far_m": round(head_far[anchor] / 1000.0, 2),   # 앵커 = 최원 유하거리
        "near_m": round(min(head_far[hi] for hi in picked) / 1000.0, 2),
        "span_m": round(span / 1000.0, 2),              # 설계면적 폭(배관거리)
        "total_m": round(total / 1000.0, 2),            # corridor 총연장
        "max_load": max(loads.values(), default=0),     # 주배관 관경 결정값
    }


def _worst_view(sess: dict) -> dict | None:
    """화면용 — 최불리 배관망(corridor)·앵커·담당 헤드 수.

    corridor 간선은 좌표 4개 + load 를 함께 싣는다(화면이 굵기/색을 정한다).
    앵커는 «가장 불리한 지점» 이라 따로 강조한다.
    """
    w = sess.get("worst")
    if not w:
        return None
    b = sess["edit"].board
    pts = b.pts
    disks = b.disks
    an = w.get("anchor")
    return {
        "k": len(w["heads"]),
        "reachable": w["reachable"],
        "far_m": w["far_m"],
        "near_m": w["near_m"],
        "span_m": w.get("span_m", 0.0),
        "total_m": w.get("total_m", 0.0),
        "max_load": w.get("max_load", 0),
        "sheet": w.get("sheet"),
        "heads": [[_r1(disks[hi][0]), _r1(disks[hi][1]), _r1(disks[hi][2])]
                  for hi in w["heads"] if hi < len(disks)],
        "anchor": ([_r1(disks[an][0]), _r1(disks[an][1]), _r1(disks[an][2])]
                   if isinstance(an, int) and an < len(disks) else None),
        "corridor": [[_r1(pts[a][0]), _r1(pts[a][1]),
                      _r1(pts[c][0]), _r1(pts[c][1]), int(load)]
                     for (a, c), load in w.get("loads", {}).items()],
    }


# ────────────────────────────────────────────── 자동 이음 · 도면 장 · 덩이


def _sheet_frames(board) -> list[dict]:
    """한 파일에 도면이 여러 장 들어 있는지 — 모듈 A 의 규칙을 그대로 부른다.

    국내 도서는 도면 한 장이 곧 파일 하나가 아니다(A 실측 — 죽전 6장·청라
    포레스트 3장·대구오페라 단위세대 5장). 여러 장을 한 망으로 보면 최불리 30 이
    서로 다른 도면의 헤드를 섞어 뽑아 계산이 성립하지 않는다.

    A 의 `detect_sheet_frames` 는 헤드 좌표(`.pos`)만 본다 — 문턱도 상수가 아니라
    그 도면의 헤드 간격에서 잰다. 그래서 규칙을 베끼지 않고 그대로 호출한다.
    """
    disks = getattr(board, "disks", None) or ()
    if len(disks) < 24:
        return []

    class _Head:  # A 가 보는 것은 .pos 하나뿐이다
        __slots__ = ("pos",)

        def __init__(self, p):
            self.pos = p

    try:
        from remote30_prototype import detect_sheet_frames
    except Exception as exc:  # noqa: BLE001 — A 가 없어도 손질은 돌아야 한다
        print(f"[손질] 도면 장 나누기 건너뜀 — 모듈 A 미탑재: {exc}")
        return []
    try:
        return detect_sheet_frames(
            [_Head((float(d[0]), float(d[1]))) for d in disks])
    except Exception as exc:  # noqa: BLE001
        print(f"[손질] 도면 장 나누기 실패: {exc}")
        return []


def _restrict_to_worst(payload: dict, board, worst: dict) -> dict:
    """변환 대상을 최불리 K 헤드로 좁힌다 — 헤드만 지우고 배관은 안 자른다.

    간선을 직접 잘라내고 싶은 유혹이 있지만 그러면 안 된다. 모듈 E 의
    `build_planar_graph` 는 이미 «급수원에서 물 닿는 간선만 남기고, 헤드로
    가지 않는 막다른관을 쳐내는» 단계를 갖고 있다(실측 로그: 물길 필터 →
    막다른관 삭제). 그러니 남길 헤드만 남겨 두면 그 배관은 E 가 제 규칙으로
    정리한다. 손으로 자르면 E 가 지키는 불변식(티 겹침·노드정리)을 깬다.

    hcov / disk_kinds / head_kinds 는 같은 디스크 집합을 가리키므로 함께 건다.
    ups 는 좌표 집합으로만 쓰여 남아 있어도 해가 없다.
    """
    from services.cad_import.kinds import disk_key

    keep_idx = {int(i) for i in (worst or {}).get("heads") or ()}
    disks = list(board.disks)
    kept = [disks[i] for i in sorted(keep_idx) if 0 <= i < len(disks)]
    if not kept:
        return payload

    keys = {disk_key(d[0], d[1], d[2]) for d in kept}
    out = dict(payload)
    out["hcov"] = [list(d) for d in kept]
    dk = payload.get("disk_kinds") or []
    out["disk_kinds"] = [dk[i] for i in sorted(keep_idx) if 0 <= i < len(dk)]

    fresh = []
    for rec in payload.get("head_kinds") or ():
        if not isinstance(rec, dict) or "c" not in rec:
            continue
        c = rec["c"]
        r = rec.get("head_r")
        if r is None and rec.get("tri_side"):
            r = float(rec["tri_side"]) / math.sqrt(3.0)
        if r is None:
            continue
        if disk_key(c[0], c[1], r) in keys:
            fresh.append(dict(rec))
    out["head_kinds"] = fresh
    print(f"[변환] 최불리 {len(kept)} 헤드로 범위를 좁힘 "
          f"(도면 헤드 {len(disks)} · 종류표 {len(fresh)}행)")
    return out


def _emit_pipenet(sess: dict, kfp: dict, out_dir: Path) -> dict:
    """.kfp → PIPENET .sdf(+표준 .slf). 11번 모듈의 변환기를 그대로 쓴다.

    모듈 A 는 제 표(`PipeTables`)에서 SDF 를 직접 찍지만, 모듈 F 의 산출물은
    모듈 E 계열의 .kfp 다. 여기서 SDF 를 새로 짜면 규약이 셋으로 갈라지므로,
    이미 있는 KFP↔SDF 변환기(`kfp_sdf_converter`)를 태운다.
    SDF 는 SLF(라이브러리) 없이 열면 PIPENET 이 "pipe bore must be given" 을
    내므로 항상 한 세트로 묶는다.
    """
    info: dict = {"ok": False}
    try:
        from kfp_sdf_converter import emit_sdf_xml, kfp_dict_to_network
        net = kfp_dict_to_network(kfp)
        xml = emit_sdf_xml(net)
    except Exception as exc:  # noqa: BLE001 — SDF 실패가 .kfp 를 무르게 하면 안 된다
        print(f"[변환] SDF 생성 실패 — .kfp 는 정상입니다: {exc}")
        info["error"] = f"{type(exc).__name__}: {exc}"
        return info

    sdf_path = out_dir / f"{sess['id']}.sdf"
    sdf_path.write_text(xml, encoding="utf-8")
    sess["sdf_path"] = str(sdf_path)
    info.update({
        "ok": True, "bytes": sdf_path.stat().st_size,
        "nodes": xml.count("<Node "), "pipes": xml.count("<Pipe "),
        "nozzles": xml.count("<Nozzle "),
    })

    try:
        from kfp_sdf_converter import _resolve_standard_slf
        slf = _resolve_standard_slf()
    except Exception:  # noqa: BLE001
        slf = None
    if slf and os.path.isfile(str(slf)):
        sess["slf_path"] = str(slf)
        info["slf"] = os.path.basename(str(slf))
    else:
        info["slf"] = None
        print("[변환] 표준 SLF 를 찾지 못했습니다 — SDF 만 담습니다.")
    print(f"[변환] PIPENET SDF · 노드 {info['nodes']} · 배관 {info['pipes']} · "
          f"노즐 {info['nozzles']} · {info['bytes']:,} bytes")
    return info
