"""직각화 사전 실측 — 최종 배관망의 각도 분포 · 사이클 유무.

usage:  python scripts/_ortho_probe.py
"""
from __future__ import annotations

import io
import math
import sys
from collections import defaultdict
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import remote30_prototype as rp  # noqa: E402

WEST_UNIT_POLY = [
    (244500.0, -243500.0), (253500.0, -243500.0),
    (253500.0, -221500.0), (244500.0, -221500.0),
]
TOL = 1e-6


def run(load_mode: bool, on_residual_cycle: str = "preserve", ortho: bool = True):
    bundle = rp.parse_dxf_bundle(ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf")
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    region = rp.HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(ents, layer_cat, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    audit: dict = {}
    sel = rp.select_worst30_heads_anchored(
        ents, layer_cat, alarm_xy=(cx, cy), head_region=region, audit_out=audit,
        load_mode=load_mode, on_residual_cycle=on_residual_cycle, ortho=ortho)
    return sel, audit


def report(tag: str, sel, audit) -> None:
    diag = []
    for a, b, L in sel.edges:
        dx, dy = abs(b[0] - a[0]), abs(b[1] - a[1])
        if dx > TOL and dy > TOL:
            diag.append((a, b, L, math.degrees(math.atan2(dy, dx))))
    nodes = {n for a, b, _L in sel.edges for n in (a, b)}
    cyc = len(sel.edges) - (len(nodes) - 1)          # 연결그래프 가정 시 독립 사이클 수
    adj: dict = defaultdict(set)
    for a, b, _L in sel.edges:
        adj[a].add(b)
        adj[b].add(a)
    seen, comps = set(), 0
    for n in nodes:
        if n in seen:
            continue
        comps += 1
        stack = [n]
        seen.add(n)
        while stack:
            u = stack.pop()
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
    print(f"\n=== {tag} ===")
    print(f"pipe {len(sel.edges)} / node {len(nodes)} / 연결성분 {comps}")
    print(f"독립 사이클 수 = E-V+C = {len(sel.edges) - len(nodes) + comps}")
    print(f"비직각(대각) 간선 : {len(diag)} / {len(sel.edges)}"
          f"  ({100*len(diag)/max(1,len(sel.edges)):.1f}%)")
    if diag:
        diag.sort(key=lambda t: -t[2])
        print(f"  대각 총연장 {sum(d[2] for d in diag):.1f}mm"
              f" / 전체 {sum(L for _a,_b,L in sel.edges):.1f}mm")
        print("  상위 8개 (길이, 각도°, dx, dy):")
        for a, b, L, ang in diag[:8]:
            print(f"    {L:9.1f}  {ang:5.1f}°  dx={b[0]-a[0]:9.1f} dy={b[1]-a[1]:9.1f}")
        angs = sorted(round(d[3], 1) for d in diag)
        print(f"  각도 범위 {angs[0]}° ~ {angs[-1]}°")


def histogram(sel) -> None:
    """짧은 성분(minor) 크기 분포 — 미세 어긋남 vs 진짜 대각 구분용."""
    minors = []
    for a, b, L in sel.edges:
        dx, dy = abs(b[0] - a[0]), abs(b[1] - a[1])
        if dx > TOL and dy > TOL:
            minors.append((min(dx, dy), max(dx, dy), L))
    if not minors:
        print("\n  minor 분포: 대각 간선 없음")
        return
    minors.sort()
    print("\n  minor(짧은 성분) 누적분포:")
    for thr in (1, 5, 10, 25, 50, 100, 200, 500, 1000, 10 ** 9):
        n = sum(1 for m, _M, _L in minors if m <= thr)
        print(f"    minor <= {thr:>10} mm : {n:3d} / {len(minors)}")
    print("  minor 비율(minor/major) 분포:")
    ratios = sorted(m / M for m, M, _L in minors)
    for q in (0.1, 0.25, 0.5, 0.75, 0.9, 1.0):
        print(f"    p{int(q*100):>3} = {ratios[min(len(ratios)-1, int(q*(len(ratios)-1)))]:.4f}")
    grow = sum((m + M) - L for m, M, L in minors)
    print(f"  L자 분해 시 연장 증가: +{grow:.1f}mm "
          f"({100*grow/sum(L for _a,_b,L in sel.edges):.2f}%)")


def raw_graph_check() -> None:
    """병합 이전 원시 그래프의 직각도 — 대각선이 도면 탓인지 병합 탓인지 판별."""
    bundle = rp.parse_dxf_bundle(ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf")
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    graph, edge_len = rp._build_graph(ents, layer_categories=layer_cat)
    diag = n = 0
    diag_len = tot = 0.0
    for (a, b), L in edge_len.items():
        n += 1
        tot += L
        if abs(b[0] - a[0]) > TOL and abs(b[1] - a[1]) > TOL:
            diag += 1
            diag_len += L
    print("\n=== 원시 그래프(병합 전) ===")
    print(f"간선 {n} / 비직각 {diag} ({100*diag/max(1,n):.1f}%)")
    print(f"비직각 연장 {diag_len:.1f} / 전체 {tot:.1f} mm ({100*diag_len/max(1,tot):.1f}%)")
    minors = sorted(min(abs(b[0]-a[0]), abs(b[1]-a[1])) for (a, b) in edge_len
                    if abs(b[0]-a[0]) > TOL and abs(b[1]-a[1]) > TOL)
    print("  minor 누적:", ", ".join(
        f"<={t}mm:{sum(1 for m in minors if m <= t)}" for t in (0.5, 1, 5, 25, 100)),
        f"/ {len(minors)}")
    angs = sorted(math.degrees(math.atan2(min(abs(b[1]-a[1]), abs(b[0]-a[0])),
                                          max(abs(b[1]-a[1]), abs(b[0]-a[0]))))
                  for (a, b) in edge_len
                  if abs(b[0]-a[0]) > TOL and abs(b[1]-a[1]) > TOL)
    print("  축과의 이탈각 누적:", ", ".join(
        f"<={t}°:{sum(1 for x in angs if x <= t)}" for t in (0.1, 1, 5, 12, 30, 45)))


def invariants(off, on, audit_on) -> None:
    lo = sum(L for _a, _b, L in off.edges)
    ln = sum(L for _a, _b, L in on.edges)
    print(f"  총연장 {lo:.3f} → {ln:.3f} mm  (오차 {100*abs(ln-lo)/lo:.6f}%)")
    print(f"  소스 좌표 불변: {off.source_pos == on.source_pos}")
    hp_off = [h.pos for h in off.heads]
    hp_on = [h.pos for h in on.heads]
    print(f"  헤드 {len(hp_off)}→{len(hp_on)}, 좌표 동일: {hp_off == hp_on}")
    print(f"  거리 동일: {off.distances == on.distances}")
    ea = sum(len(v) for v in (off.elbow_fittings or {}).values())
    eb = sum(len(v) for v in (on.elbow_fittings or {}).values())
    print(f"  elbow fitting 수 {ea} → {eb}")
    print(f"  ortho audit: {audit_on.get('ortho')}")


def main() -> None:
    raw_graph_check()
    for mode, pol in ((False, "preserve"), (True, "force_tree")):
        sel_off, _ = run(mode, pol, ortho=False)
        sel, audit = run(mode, pol)
        report(f"load_mode={mode} / {pol}", sel, audit)
        histogram(sel)
        print("\n  --- 불변식 (ortho off → on) ---")
        invariants(sel_off, sel, audit)


if __name__ == "__main__":
    main()
