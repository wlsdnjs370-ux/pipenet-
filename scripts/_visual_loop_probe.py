"""화면상 '루프처럼 보이는' 구간의 정체 실측.

그래프 사이클은 0인데 닫힌 영역이 보이는 원인 후보를 각각 센다.
usage:  python scripts/_visual_loop_probe.py
"""
from __future__ import annotations

import io
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
NEAR_MM = 60.0


def build(ortho: bool = True):
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
        ortho=ortho)
    return sel, audit


def _cross(p1, p2, p3, p4):
    """끝점을 공유하지 않는 두 세그먼트의 교차점. 없으면 None."""
    d1 = (p2[0] - p1[0], p2[1] - p1[1])
    d2 = (p4[0] - p3[0], p4[1] - p3[1])
    den = d1[0] * d2[1] - d1[1] * d2[0]
    if abs(den) < 1e-12:
        return None
    t = ((p3[0] - p1[0]) * d2[1] - (p3[1] - p1[1]) * d2[0]) / den
    u = ((p3[0] - p1[0]) * d1[1] - (p3[1] - p1[1]) * d1[0]) / den
    if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        return (p1[0] + t * d1[0], p1[1] + t * d1[1])
    return None


def _overlap(p1, p2, p3, p4):
    """같은 직선 위 두 세그먼트의 겹침 길이. 없으면 0."""
    d1 = (p2[0] - p1[0], p2[1] - p1[1])
    n1 = (d1[0] ** 2 + d1[1] ** 2) ** 0.5
    if n1 < 1e-12:
        return 0.0
    ux, uy = d1[0] / n1, d1[1] / n1
    for q in (p3, p4):
        if abs((q[0] - p1[0]) * uy - (q[1] - p1[1]) * ux) > 1e-6:
            return 0.0
    s3 = (p3[0] - p1[0]) * ux + (p3[1] - p1[1]) * uy
    s4 = (p4[0] - p1[0]) * ux + (p4[1] - p1[1]) * uy
    return max(0.0, min(n1, max(s3, s4)) - max(0.0, min(s3, s4)))


def main():
    ortho = "--no-ortho" not in sys.argv
    print(f"=== ortho={ortho} ===")
    sel, audit = build(ortho)
    edges = sel.edges
    nodes = sorted({n for a, b, _L in edges for n in (a, b)})
    adj = defaultdict(set)
    for a, b, _L in edges:
        adj[a].add(b)
        adj[b].add(a)

    print(f"간선 {len(edges)} · 노드 {len(nodes)}")

    # ── ① 거의 겹친 별개 노드 (비인접) ──────────────────────────────────
    near = []
    for i, p in enumerate(nodes):
        for q in nodes[i + 1:]:
            if abs(p[0] - q[0]) > NEAR_MM:
                continue
            if abs(p[1] - q[1]) > NEAR_MM:
                continue
            if q in adj[p]:
                continue
            d = ((p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2) ** 0.5
            if d <= NEAR_MM:
                near.append((round(d, 2), p, q))
    near.sort()
    print(f"\n① 비인접인데 {NEAR_MM}mm 이내로 붙은 노드쌍: {len(near)}")
    for d, p, q in near[:10]:
        print(f"   {d}mm  {p} ~ {q}")

    # ── ② 노드 없이 교차/겹치는 간선쌍 ──────────────────────────────────
    segs = [(a, b) for a, b, _L in edges]
    cross, overlap = [], []
    for i, (a1, b1) in enumerate(segs):
        for a2, b2 in segs[i + 1:]:
            if {a1, b1} & {a2, b2}:
                continue
            pt = _cross(a1, b1, a2, b2)
            if pt is not None:
                cross.append((pt, (a1, b1), (a2, b2)))
                continue
            ov = _overlap(a1, b1, a2, b2)
            if ov > 1e-9:
                overlap.append((round(ov, 2), (a1, b1), (a2, b2)))
    overlap.sort(reverse=True)
    def _kind(pt, e):
        d = min(((pt[0] - q[0]) ** 2 + (pt[1] - q[1]) ** 2) ** 0.5 for q in e)
        return "end" if d < 1e-6 else "mid"

    tee = [c for c in cross if "end" in (_kind(c[0], c[1]), _kind(c[0], c[2]))]
    ex = [c for c in cross if c not in tee]
    on_node = sum(1 for pt, _e1, _e2 in cross
                  if any(abs(pt[0] - n[0]) < 1e-6 and abs(pt[1] - n[1]) < 1e-6
                         for n in nodes))
    print(f"\n② 노드 없이 교차하는 간선쌍: {len(cross)}"
          f"  (T분기 {len(tee)} · X교차 {len(ex)} · 교차점이 기존 노드 {on_node})")
    for pt, e1, e2 in ex[:6]:
        print(f"   X {pt}\n      {e1}\n      {e2}")
    print(f"\n③ 같은 축선에서 구간이 겹치는 간선쌍: {len(overlap)}")
    for ov, e1, e2 in overlap[:10]:
        print(f"   겹침 {ov}mm\n      {e1}\n      {e2}")

    # ── ④ 프론트가 겹쳐 그리는 추정-edge 수 ─────────────────────────────
    for k in ("head_drops",):
        v = audit.get(k) or []
        print(f"\n④ {k}: {len(v)}")

    print(f"\northo audit: {audit.get('ortho')}")

    # ── ⑤ 불변식 ────────────────────────────────────────────────────────
    seen, comps = set(), 0
    for n in adj:
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
    print(f"\n⑤ 연결성분 {comps} · 독립 사이클 {len(edges) - len(nodes) + comps}")
    print(f"   총연장 {sum(L for _a, _b, L in edges):.3f} mm")
    print(f"   최소 세그먼트 {min(L for _a, _b, L in edges):.4f} mm")
    print(f"   비직각 {sum(1 for a, b, _L in edges if abs(a[0]-b[0]) > 1e-9 and abs(a[1]-b[1]) > 1e-9)}")
    print(f"   source {sel.source_pos} · 헤드 {len(sel.heads)} · 부속 {sum(len(v) for v in sel.elbow_fittings.values())}")
    missing = [h.pos for h in sel.heads if h.pos not in nodes]
    print(f"   최종망에 없는 헤드: {len(missing)}")


if __name__ == "__main__":
    main()
