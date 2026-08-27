# -*- coding: utf-8 -*-
"""R7(헤드틈 접속)이 왜 0건인가 — 관문을 하나씩 센다.

`_join_head_gap_endpoints` 를 «같은 입력으로» 계측판으로 갈아끼워, 끝점쌍이
어느 조건에서 떨어지는지 센다. 실제 파이프라인을 그대로 태우므로 입력이
달라질 여지가 없다.

관문(소스 순서 그대로):
    ends      차수 1 인 끝점이 몇 개인가          ← 여기가 0 이면 나머지는 무의미
    near      이웃 칸에 헤드가 있는가
    gap       0 < 거리 <= max_gap
    axis      간극이 축에 평행한가
    wings     양쪽 인접 간선도 같은 축인가
    dir       두 간선이 간극 반대로 뻗는가(런의 연속)
    head      간극 «안» 에 그 선 위 헤드가 있는가
    joined    루프 금지·끝점 유지 통과 → 실제 접속

    python scripts/_probe_r7_gate.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
CNT = Counter()
GAPS: list = []


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    # ★상한을 올려 가며 «R7 의 자체 관문» 을 몇 건이 통과하는지 본다.
    #   틈 안에 배관이 있나 없나가 아니라, 동일선상 + 사이에 헤드 인가가 근거다.
    ap.add_argument("--sweep", default="400,600,700,800,1000,1200")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    import remote30_prototype as A
    from remote30_graph import HeadRegion

    real = A._join_head_gap_endpoints
    axis_index = A._axis_index

    def probe(graph, edge_len, head_pts,
              max_gap_mm=A.HEAD_GAP_JOIN_MAX_MM,
              tol_mm=A.HEAD_GAP_JOIN_TOL_MM, joins_out=None):
        """소스와 같은 순서로 관문을 세고, 마지막에 진짜를 부른다."""
        ends = [n for n, nb in graph.items() if len(nb) == 1]
        CNT["graph_nodes"] = len(graph)
        CNT["ends"] = len(ends)
        CNT["heads"] = len(head_pts or ())
        deg = Counter(len(nb) for nb in graph.values())
        CNT["deg1"] = deg[1]
        CNT["deg2"] = deg[2]
        CNT["deg3+"] = sum(v for k, v in deg.items() if k >= 3)
        if ends and head_pts:
            inv = 1.0 / max_gap_mm
            egrid = defaultdict(list)
            for n in ends:
                egrid[(int(math.floor(n[0] * inv)),
                       int(math.floor(n[1] * inv)))].append(n)
            hgrid = defaultdict(list)
            for hp in head_pts:
                hgrid[(int(math.floor(hp[0] * inv)),
                       int(math.floor(hp[1] * inv)))].append(hp)
            seen = set()
            for u in ends:
                wu = next(iter(graph[u]))
                cgx = int(math.floor(u[0] * inv))
                cgy = int(math.floor(u[1] * inv))
                cells = [(cgx + dx, cgy + dy)
                         for dx in (-1, 0, 1) for dy in (-1, 0, 1)]
                near_heads = [hp for c in cells for hp in hgrid.get(c, ())]
                if not near_heads:
                    CNT["reject_no_near_head"] += 1
                    continue
                CNT["pass_near"] += 1
                for v in (n for c in cells for n in egrid.get(c, ())):
                    key = (min(u, v), max(u, v))
                    if v == u or key in seen:
                        continue
                    seen.add(key)
                    CNT["pairs"] += 1
                    gx, gy = v[0] - u[0], v[1] - u[1]
                    g = math.hypot(gx, gy)
                    if not (0.0 < g <= max_gap_mm):
                        CNT["reject_gap"] += 1
                        continue
                    CNT["pass_gap"] += 1
                    GAPS.append(g)
                    axis = axis_index(u, v, tol_mm)
                    if axis < 0:
                        CNT["reject_axis"] += 1
                        continue
                    CNT["pass_axis"] += 1
                    wv = next(iter(graph[v]))
                    if (axis_index(u, wu, tol_mm) != axis
                            or axis_index(v, wv, tol_mm) != axis):
                        CNT["reject_wings"] += 1
                        continue
                    CNT["pass_wings"] += 1
                    if ((wu[0] - u[0]) * gx + (wu[1] - u[1]) * gy >= 0
                            or (wv[0] - v[0]) * gx + (wv[1] - v[1]) * gy <= 0):
                        CNT["reject_dir"] += 1
                        continue
                    CNT["pass_dir"] += 1
                    if not any(
                        0.0 < ((hp[0] - u[0]) * gx + (hp[1] - u[1]) * gy) < g * g
                        and abs((hp[0] - u[0]) * gy
                                - (hp[1] - u[1]) * gx) <= tol_mm * g
                        for hp in near_heads
                    ):
                        CNT["reject_head_between"] += 1
                        continue
                    CNT["pass_head_between"] += 1
        n = real(graph, edge_len, head_pts, max_gap_mm, tol_mm, joins_out)
        CNT["joined"] = n
        return n

    A._join_head_gap_endpoints = probe
    try:
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        cat = {}
        for nm in {str(e.get("l") or "0") for e in ents}:
            try:
                cat[nm] = A._categorize_layer(nm)
            except Exception:  # noqa: BLE001
                cat[nm] = "OTHER"
        heads = A.detect_heads(ents, cat)
        pts = [(h.pos[0], h.pos[1]) for h in heads]
        sheet = A.sheet_frame_at(pts)
        inside = pts
        if sheet is not None:
            x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
            inside = [q for q in pts if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
        cx = sum(q[0] for q in inside) / len(inside)
        cy = sum(q[1] for q in inside) / len(inside)
        alarm = min(inside, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
        zones = A.head_bbox_for_region(pts, alarm)
        A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
            head_region=HeadRegion.from_rects(zones), zones=zones, k=30)
    finally:
        A._join_head_gap_endpoints = real

    print(f"\n{dxf.name} · R7 관문 (상한 {A.HEAD_GAP_JOIN_MAX_MM:.0f}mm · "
          f"오차 {A.HEAD_GAP_JOIN_TOL_MM:.0f}mm)\n")
    print(f"  그래프 절점 {CNT['graph_nodes']:,} — "
          f"차수1 {CNT['deg1']:,} · 차수2 {CNT['deg2']:,} · "
          f"차수3+ {CNT['deg3+']:,}")
    print(f"  헤드 {CNT['heads']:,}\n")
    rows = [
        ("끝점(차수1)", CNT["ends"], None),
        ("  이웃칸에 헤드 있음", CNT["pass_near"], CNT["reject_no_near_head"]),
        ("훑은 끝점쌍", CNT["pairs"], None),
        ("  거리 통과", CNT["pass_gap"], CNT["reject_gap"]),
        ("  축 평행", CNT["pass_axis"], CNT["reject_axis"]),
        ("  양날개 같은 축", CNT["pass_wings"], CNT["reject_wings"]),
        ("  런의 연속(방향)", CNT["pass_dir"], CNT["reject_dir"]),
        ("  간극 안에 헤드", CNT["pass_head_between"],
         CNT["reject_head_between"]),
        ("실제 접속", CNT["joined"], None),
    ]
    for label, ok, bad in rows:
        s = f"  {label:<24} {ok:>8,}"
        if bad is not None:
            s += f"   (떨어짐 {bad:,})"
        print(s)
    if GAPS:
        GAPS.sort()
        print(f"\n  거리 통과분 중앙값 {GAPS[len(GAPS) // 2]:.0f}mm · "
              f"최대 {GAPS[-1]:.0f}mm")

    # ── 상한 스윕 — R7 을 그대로 다시 돌리되 상한만 바꾼다.
    #    그래프는 매번 새로 얻어야 한다(앞선 접속이 남으면 셈이 오염된다).
    print("\n■ 상한을 올리면 몇 건이 «R7 의 관문 전부» 를 통과하나\n")
    print(f"  {'상한':>7} {'접속':>7} {'틈 중앙값':>10} {'틈 최대':>9}")
    print("  " + "-" * 36)
    for lim in [float(s) for s in a.sweep.split(",") if s.strip()]:
        box: dict = {}

        def sweep(graph, edge_len, head_pts, *args, _lim=lim, _box=box, **kw):
            joins: list = []
            n = real(graph, edge_len, head_pts, _lim,
                     A.HEAD_GAP_JOIN_TOL_MM, joins)
            _box["n"] = n
            _box["gaps"] = sorted(float(j["gap_mm"]) for j in joins)
            return n

        A._join_head_gap_endpoints = sweep
        try:
            A.select_worst30_heads_anchored(
                pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
                head_region=HeadRegion.from_rects(zones), zones=zones, k=30)
        except Exception as exc:  # noqa: BLE001
            print(f"  {lim:>7.0f} 실패 — {type(exc).__name__}: {exc}")
            continue
        finally:
            A._join_head_gap_endpoints = real
        g = box.get("gaps") or []
        print(f"  {lim:>7.0f} {box.get('n', 0):>7,} "
              f"{(g[len(g) // 2] if g else 0):>10.0f} "
              f"{(g[-1] if g else 0):>9.0f}")
    print(f"\n  (현재 상한 {A.HEAD_GAP_JOIN_MAX_MM:.0f}mm · "
          f"동일선상 오차 {A.HEAD_GAP_JOIN_TOL_MM:.0f}mm)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
