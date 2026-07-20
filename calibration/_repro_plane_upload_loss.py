"""평면도 업로드 누락 재현 — 드롭된 레이어를 UTF-8 리포트로 덤프.

"특정 레이어 통째" 누락 → 실제 배관망 레이어가 OTHER/ARCH/EXCLUDE 로 잘못 분류돼
Stage1 filter_pipenet_only 에서 통째 제거된 것. 드롭 레이어별 geometry 통계로 식별.
"""
import json
import math
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, _categorize_layer)

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"
OUT = Path(__file__).resolve().parent / "_plane_loss_report.json"


def _flat_xy(p):
    """p 를 [(x,y),...] 로 정규화 (flat [x,y,x,y] 또는 [[x,y],...] 둘 다 지원)."""
    if not p:
        return []
    if isinstance(p[0], (list, tuple)):
        return [(float(q[0]), float(q[1])) for q in p if len(q) >= 2]
    return [(float(p[i]), float(p[i + 1])) for i in range(0, len(p) - 1, 2)]


def _seg_len(e):
    """LINE/PL 근사 총 길이 — 배관망 여부 판정용."""
    t = e.get("t")
    if t not in ("L", "PL"):
        return 0.0
    pts = _flat_xy(e.get("p", []))
    tot = 0.0
    for i in range(len(pts) - 1):
        tot += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
    return tot


def main():
    bundle = parse_dxf_bundle(DXF)
    ents = bundle.entities
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)

    per_layer = {}
    for e in ents:
        L = e.get("l", "")
        d = per_layer.setdefault(L, {"count": 0, "types": Counter(),
                                     "line_len": 0.0,
                                     "xs": [], "ys": []})
        d["count"] += 1
        d["types"][e.get("t")] += 1
        d["line_len"] += _seg_len(e)
        # bbox 표본(모든 정점)
        for (qx, qy) in _flat_xy(e.get("p") or []):
            d["xs"].append(qx); d["ys"].append(qy)

    kept_layers = set(e.get("l", "") for e in pipe_ents)

    report = {"dxf": str(DXF), "entities": len(ents), "layers": len(bundle.layers),
              "after_filter": len(pipe_ents), "layer_detail": []}
    for L, d in sorted(per_layer.items(), key=lambda kv: -kv[1]["count"]):
        cat = layer_cat.get(L, "?")
        kept = L in kept_layers
        bbox = None
        if d["xs"]:
            bbox = [round(min(d["xs"])), round(min(d["ys"])),
                    round(max(d["xs"])), round(max(d["ys"]))]
        report["layer_detail"].append({
            "layer": L,
            "category": cat,
            "kept": kept,
            "count": d["count"],
            "types": dict(d["types"]),
            "line_len_mm": round(d["line_len"]),
            "bbox": bbox,
        })

    OUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print("wrote", OUT)
    # 드롭됐지만 배관망 의심(라인길이 큼) 레이어 요약(ascii-safe: 인덱스만)
    print("\nDROPPED layers with pipe-like geometry (line_len desc):")
    dropped = [x for x in report["layer_detail"] if not x["kept"]]
    dropped.sort(key=lambda x: -x["line_len_mm"])
    for i, x in enumerate(dropped[:12]):
        print(f"  #{i}: cat={x['category']:7s} count={x['count']:6d} "
              f"line_len={x['line_len_mm']:>10} types={x['types']}")


if __name__ == "__main__":
    main()
