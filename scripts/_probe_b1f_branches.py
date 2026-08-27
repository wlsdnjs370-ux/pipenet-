# -*- coding: utf-8 -*-
"""「자동이 가지배관을 못 잡는다」 — 어느 단계에서 빠지는지 잰다.

R7(헤드틈 접속)·R8(관통 티)·R9(중복 선분)은 `select_worst30_heads_anchored`
안에서 «호출은» 되고 있다(소스 확인). 그러니 문제는 «안 부른다» 가 아니라
«불렀는데 안 걸린다» 이거나, 걸린 뒤에 다른 곳에서 잘려 나가는 것이다.

여기서는 추측하지 않고 감사(ExtractionAudit)를 그대로 펼쳐 본다:

    tee_splits · covered_drops · head_gap_joins · crossing_tees   ← 복원 건수
    fragments · unreachable_heads · head_drops                    ← 남은 손실
    anchor_window                                                 ← 잘린 범위

그리고 뽑힌 망의 «모양» 을 본다 — 가지관이 살아 있으면 차수 1 인 말단이 헤드
수만큼 나온다. 주관만 남았으면 말단이 몇 개 안 된다.

    python scripts/_probe_b1f_branches.py [도면.dxf] [--k 30]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--k", type=int, default=30)
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

    print(f"{dxf.name} ({dxf.stat().st_size / 1024 / 1024:.1f} MB)\n")
    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    cat = {}
    for n in {str(e.get("l") or "0") for e in ents}:
        try:
            cat[n] = A._categorize_layer(n)
        except Exception:  # noqa: BLE001
            cat[n] = "OTHER"

    # 어떤 레이어가 배관으로 인정됐나 — 가지관이 여기 없으면 그 앞에서 끝이다.
    pipe_layers = sorted(n for n, c in cat.items() if c == "PIPE")
    print(f"■ PIPE 로 인정된 레이어 {len(pipe_layers)}개")
    for n in pipe_layers:
        cnt = sum(1 for e in ents if str(e.get("l") or "0") == n)
        print(f"    {n}  ({cnt:,} entity)")
    prom = list(getattr(bundle, "promoted_layers", None) or ())
    print(f"■ 헤드틈 지문으로 승격된 레이어 {len(prom)}개")
    for p in prom:
        print(f"    {p}")

    heads = A.detect_heads(ents, cat)
    print(f"\n■ 헤드 검출 {len(heads):,}개")

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
    print(f"■ 알람밸브(모의) {alarm[0]:.0f}, {alarm[1]:.0f}")

    audit: dict = {}
    sel = A.select_worst30_heads_anchored(
        pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
        head_region=HeadRegion.from_rects(zones), zones=zones,
        k=a.k, audit_out=audit)

    print("\n■ 실배관 복원 — 몇 건이 걸렸나")
    for key, label in (("tee_splits", "T분기 쪼갬"),
                       ("covered_drops", "중복 선분 제거 (R9)"),
                       ("head_gap_joins", "헤드틈 접속 (R7)"),
                       ("crossing_tees", "관통 티 (R8)")):
        v = audit.get(key)
        n = len(v) if isinstance(v, (list, tuple)) else v
        print(f"    {label:<24} {n}")
    hg = audit.get("head_gap_joins") or []
    if hg:
        gaps = [float(g.get("gap_mm", 0)) for g in hg if isinstance(g, dict)]
        if gaps:
            gaps.sort()
            print(f"      틈 중앙값 {gaps[len(gaps) // 2]:.0f}mm · "
                  f"최대 {gaps[-1]:.0f}mm (상한 {A.HEAD_GAP_JOIN_MAX_MM:.0f})")

    print("\n■ 남은 손실")
    fr = audit.get("fragments") or {}
    print(f"    소스 미연결 조각   {fr.get('count')}개 · "
          f"{(fr.get('detached_len_mm') or 0) / 1000:.1f} m")
    uh = (audit.get("heads") or {}).get("unreachable") or []
    print(f"    미도달 헤드        {len(uh)}개")
    hd = audit.get("head_drops")
    if isinstance(hd, dict):
        print(f"    헤드 결합선        {hd.get('count')}개 · "
              f"최대 {hd.get('max_mm')}mm")
    elif hd:
        ds = [float(d.get("len_mm", 0)) for d in hd if isinstance(d, dict)]
        print(f"    헤드 결합선        {len(hd)}개"
              + (f" · 최대 {max(ds):.0f}mm" if ds else ""))
    sa = audit.get("source_attach") or {}
    print(f"    급수원 결합        {sa.get('attached_to')} · "
          f"escalation {sa.get('escalation')} · "
          f"헤드보유 {sa.get('comp_head_count')}")

    # ★소스가 닿는 최대망에 헤드가 몇 개나 들어 있나 — 「가지관을 못 잡는다」의
    #   실체는 대개 여기다. 최대망이 주관뿐이면 헤드가 거의 없다.
    print("\n■ 배관 길이 분포 — 가지관은 짧고 주관은 길다")
    ed_all = list(getattr(sel, "edges", None) or ())
    if ed_all:
        ls = sorted(float(e[2]) for e in ed_all)
        import statistics
        buckets = [(0, 500), (500, 1500), (1500, 3000), (3000, 8000),
                   (8000, 10 ** 9)]
        for lo, hi in buckets:
            n = sum(1 for v in ls if lo <= v < hi)
            lab = f"{lo/1000:g}~{hi/1000:g} m" if hi < 10 ** 9 else "8 m~"
            bar = "█" * int(40 * n / max(1, len(ls)))
            print(f"    {lab:>10}  {n:>4}  {bar}")
        print(f"    중앙값 {statistics.median(ls):.0f}mm · "
              f"최대 {ls[-1]:.0f}mm")

    print("\n■ 뽑힌 망의 «모양» — 가지관이 살아 있나")
    ed = list(getattr(sel, "edges", None) or ())
    deg = Counter()
    for e in ed:
        deg[e[0]] += 1
        deg[e[1]] += 1
    leaves = [n for n, d in deg.items() if d == 1]
    tees = [n for n, d in deg.items() if d >= 3]
    total_m = sum(float(e[2]) for e in ed) / 1000.0
    hs = getattr(sel, "heads", None) or ()
    print(f"    절점 {len(deg)} · 배관 {len(ed)} · 연장 {total_m:.1f} m")
    print(f"    말단(차수1) {len(leaves)} · 분기(차수≥3) {len(tees)} · "
          f"선정 헤드 {len(hs)}")
    if hs:
        xs = [h.pos[0] for h in hs]
        ys = [h.pos[1] for h in hs]
        print(f"    헤드 퍼짐 대각 "
              f"{math.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000:.1f} m")
    # 가지관이 통째로 빠지면 말단이 헤드 수보다 훨씬 적다.
    if len(leaves) < len(hs) * 0.5:
        print("    ★말단이 헤드 수의 절반도 안 된다 — 가지관이 빠졌을 수 있다")
    else:
        print("    말단 수가 헤드 수에 준한다 — 가지관이 붙어 있다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
