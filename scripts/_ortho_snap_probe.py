"""축 스냅 충돌 원인 실측 — 왜 snap_aborted=True 인가.

usage:  python scripts/_ortho_snap_probe.py
"""
from __future__ import annotations

import io
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import remote30_prototype as rp  # noqa: E402

WEST_UNIT_POLY = [
    (244500.0, -243500.0), (253500.0, -243500.0),
    (253500.0, -221500.0), (244500.0, -221500.0),
]


def main() -> None:
    bundle = rp.parse_dxf_bundle(ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf")
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    region = rp.HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(ents, layer_cat, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    sel = rp.select_worst30_heads_anchored(ents, layer_cat, alarm_xy=(cx, cy),
                                           head_region=region, ortho=False)
    edges = sel.edges
    anchors = {sel.source_pos} | {h.pos for h in sel.heads}

    aud: dict = {}
    _e, _f, nmap = rp.orthogonalize_edges(edges, anchors=anchors, audit=aud)
    print("audit:", aud["ortho"])

    # 충돌 노드 직접 재현
    nodes = {n for a, b, _L in edges for n in (a, b)}
    rev = defaultdict(list)
    for n, m in nmap.items():
        rev[m].append(n)
    print(f"노드 {len(nodes)} / 매핑 이미지 {len(set(nmap.values()))}")

    # snap 을 강제 적용해 충돌만 보기 위해 내부 로직 재현 대신,
    # 결과 세그먼트 길이 분포로 L자 분해 부작용을 본다.
    segs = sorted(L for _a, _b, L in _e)
    print(f"세그먼트 {len(segs)}  최소 {segs[0]:.4f}  p10 {segs[len(segs)//10]:.2f}"
          f"  중앙 {segs[len(segs)//2]:.1f}  최대 {segs[-1]:.1f} mm")
    for thr in (0.1, 1, 5, 10, 25, 50, 100):
        print(f"    L <= {thr:>6} mm : {sum(1 for s in segs if s <= thr):3d} / {len(segs)}")
    cnt = Counter(round(s, 3) for s in segs if s <= 25)
    print("  25mm 이하 세그먼트 길이:", dict(sorted(cnt.items())[:20]))

    pre = sorted(L for _a, _b, L in edges)
    print(f"\n직교화 전 간선 {len(pre)}  최소 {pre[0]:.3f}  중앙 {pre[len(pre)//2]:.1f} mm")
    for thr in (1, 5, 25, 100):
        print(f"    L <= {thr:>4} mm : {sum(1 for s in pre if s <= thr):3d} / {len(pre)}")
    # 직교화가 새로 만든 짧은 세그먼트만 분리
    born = [(a, b, L) for a, b, L in _e if L <= 25 and (a, b, L) not in set(edges)]
    print(f"  25mm 이하 중 직교화가 새로 만든 것: {len(born)}")
    for a, b, L in sorted(born, key=lambda t: t[2])[:6]:
        print(f"    L={L:8.3f}  dx={b[0]-a[0]:9.2f} dy={b[1]-a[1]:9.2f}")


if __name__ == "__main__":
    main()
