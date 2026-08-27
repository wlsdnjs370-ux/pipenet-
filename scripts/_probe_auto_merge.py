# -*- coding: utf-8 -*-
"""[4차 검토] 자동(A) 평면 표 + 실계통도 라이저 → 통합(S740) 실측.

지금까지의 실측은 반쪽이었다:
  · 통합 실측(_verify_module_f_merge.py)은 **G-모양 최소 표** 로만 돌았다
  · 자동 경로 실측(_verify_module_f_auto.py)은 라벨 규약 확인까지만 갔다

즉 «A 의 실제 표(기준점 10 · 오프셋 0)로 stitch 가 도는가» 는 아무도 안 봤다.
A 표의 노드는 10..N, 배관 라벨은 숫자 — 라이저(1·n2..·10 / r1..)와 충돌하지
않아야 한다는 것은 주석의 약속일 뿐, 실측이 없었다.

    python scripts/_probe_auto_merge.py
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "samples" / "dxf" / "LH306동_평면도.dxf"
SYSTEM = ROOT / "data" / "sample_problem" / "대명동201동 계통도.dxf"
FAILS: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    from routes.module_f.auto import detect_head_candidates, parse_plan, run_auto
    from routes.module_f.emit import cross_check, emit_merged
    from routes.module_f.merge import combined_summary, merge_network
    from routes.module_f.subdrawing import (entities_to_world, extract_system,
                                            parse_subdrawing)

    print("[4차] 자동 평면 + 실계통도 통합 실측")

    # ── 자동 평면 표 (A 경로 그대로)
    ents, cat, _ = parse_plan(PLAN)
    heads = detect_head_candidates(ents, cat)
    xs = [h["x"] for h in heads]
    ys = [h["y"] for h in heads]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    mid = min(heads, key=lambda h: (h["x"] - cx) ** 2 + (h["y"] - cy) ** 2)
    pad = 1000.0
    got_auto = run_auto(ents, cat, alarm_xy=(mid["x"], mid["y"]),
                        rects=[[min(xs) - pad, min(ys) - pad,
                                max(xs) + pad, max(ys) + pad]], k=30)
    tbl = got_auto["tables"]
    n_head_nodes = len(tbl.nodes)
    check("자동 평면 표", n_head_nodes > 0,
          f"절점 {n_head_nodes} · 배관 {len(tbl.pipes)} · 노즐 {len(tbl.nozzles)}")

    # ── 실계통도 라이저
    s_ents, _ = parse_subdrawing(SYSTEM)
    w = entities_to_world(s_ents)
    pts = [p for s in w.segs for p in (s[2], s[3])]
    lo = min(pts, key=lambda p: (p[0], p[1]))
    hi = max(pts, key=lambda p: (p[0], p[1]))
    riser = extract_system(s_ents, lo, hi, snap_tolerance_mm=5000)
    n_riser = len(riser["nodes"])
    check("계통도 라이저", n_riser > 1,
          f"절점 {n_riser} · 배관 {len(riser['pipes'])}")

    # ── 라벨 충돌 사전 점검 — 주석의 약속을 실측으로
    head_labels = {str(n.get("label")) for n in tbl.nodes}
    riser_labels = {str(n.get("label")) for n in riser["nodes"]}
    overlap = (head_labels & riser_labels) - {"10"}
    check("노드 라벨은 기준점 10 만 겹친다", not overlap,
          f"겹침 {sorted(overlap)[:6]}" if overlap else "10 뿐")
    head_pipes = {str(p.get("label")) for p in tbl.pipes}
    riser_pipes = {str(p.get("label")) for p in riser["pipes"]}
    p_overlap = head_pipes & riser_pipes
    check("배관 라벨은 안 겹친다", not p_overlap,
          f"겹침 {sorted(p_overlap)[:6]}" if p_overlap else
          f"헤드 {len(head_pipes)} · 라이저 {len(riser_pipes)}")

    # ── S740 — 오프셋 0 (자동)
    try:
        got = merge_network(tbl, riser=riser, mode="lsp_gravity", method="auto")
    except Exception as exc:  # noqa: BLE001
        check("S740 결합 (자동 표)", False, f"{type(exc).__name__}: {exc}")
        return 1
    s = combined_summary(got)
    check("S740 결합 (자동 표)", s["merged"] is True, " · ".join(s["steps"]))
    expect = n_riser + n_head_nodes - 1
    check("절점 = 라이저 + 평면 − 1", s["nodes"] == expect,
          f"{s['nodes']} (기대 {expect})")
    check("노즐 보존", s["nozzles"] == len(tbl.nozzles), str(s["nozzles"]))

    # ── 산출 3종 교차검증
    tmp = Path(tempfile.mkdtemp(prefix="mf_automerge_"))
    try:
        files = emit_merged(got["combined"], tmp, title="자동+계통도 실측")
        for wmsg in files.get("warnings") or ():
            print(f"       ! {wmsg}")
        cc = cross_check(files)
        detail = " · ".join(
            f"{k}: {v.get('total_m')}m·노즐{v.get('nozzles')}"
            if "error" not in v else f"{k}: {v['error'][:40]}"
            for k, v in cc["per_format"].items())
        check("세 형식이 같은 배관망", cc["agree"] is True,
              detail + (f" · {cc['detail']}" if cc.get("detail") else ""))
        check("세 형식 모두 읽힘", len(cc["compared"]) == 3,
              " · ".join(cc["compared"]))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 자동 표로도 제5국면이 끝까지 돈다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
