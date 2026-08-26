# -*- coding: utf-8 -*-
"""[H-4 ~ H-6] 제5국면 S700 을 실도면으로 끝까지 태운다.

여기서 확인하는 것:
  ① S710  급수방식은 사람이 고르고, 안 고르면 결합이 막힌다
  ② S720  계통도 → 입상관 (실도면)
  ③ S740  기준점 10 을 공통 절점으로 결합, 절점 수가 «합 − 1» 이다
  ④        계통도가 없으면 평면도 단독으로 지나간다 (결합 없음도 정상)
  ⑤ S750  SDF 를 원본으로 SLF·KFP·HAS 가 나온다
  ⑥ S760  세 형식이 **같은 배관망을 가리킨다** (교차 검증)
  ⑦ S770  넷이 하나로 압축된다

평면도는 무거우므로(B1F 파싱 13분) 실제 설계 표 대신 **모양이 같은 최소 표**를
쓴다 — 결합이 보는 것은 표의 «모양» 이지 그 표가 어느 도면에서 왔는지가 아니다.
평면도 실측은 tests/test_module_f_complete.py 와 G 의 골든이 이미 덮는다.

    python scripts/_verify_module_f_merge.py
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FAILS: list[str] = []
SYSTEM_DXF = [
    ROOT / "data" / "sample_problem" / "대명동201동 계통도.dxf",
    ROOT / "samples" / "dxf" / "계통도_LH_306_배관망추출.dxf",
]


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


class _G:
    """G 의 `PipeTablesG` 모양 — 급수원 10(=BFS 1) · 헤드 둘."""

    def __init__(self):
        self.nodes = [
            {"label": "1", "x": 0, "y": 0, "elevation": 0.0,
             "io_node": "Input", "pressure_pa": 101325.0},
            {"label": "2", "x": 3000, "y": 0, "elevation": 0.0, "io_node": "No"},
            {"label": "3", "x": 6000, "y": 0, "elevation": 0.0, "io_node": "No"},
            {"label": "4", "x": 6000, "y": 3000, "elevation": 0.0, "io_node": "No"},
        ]
        self.pipes = [
            {"label": "P1", "in": "1", "out": "2", "type": "KSD 3507",
             "dia": 65, "length": 3.0, "elev": 0.0, "c": 120,
             "status": "Normal", "group": "Unset"},
            {"label": "P2", "in": "2", "out": "3", "type": "KSD 3507",
             "dia": 40, "length": 3.0, "elev": 0.0, "c": 120,
             "status": "Normal", "group": "Unset"},
            {"label": "P3", "in": "3", "out": "4", "type": "KSD 3507",
             "dia": 25, "length": 3.0, "elev": 0.0, "c": 120,
             "status": "Normal", "group": "Unset"},
        ]
        self.nozzles = [
            {"label": "1", "in": "3", "out": "@/1", "status": "1",
             "lib": "SP-HEAD", "flow_lmin": 80, "flow_m3s": 80 / 60000.0},
            {"label": "2", "in": "4", "out": "@/2", "status": "1",
             "lib": "SP-HEAD", "flow_lmin": 80, "flow_m3s": 80 / 60000.0},
        ]
        self.fittings = [{"pipe": "P2", "in": "2", "out": "3",
                          "type": "Elbow 90", "count": 1}]
        self.equipment = []
        self.meta = [("제목", "통합 검증")]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    from routes.module_f.emit import cross_check, emit_merged
    from routes.module_f.merge import (
        MergeError, combined_summary, merge_network, to_head_tables)
    from routes.module_f.subdrawing import extract_system, parse_subdrawing

    print("[H-4 ~ H-6] 제5국면 S700 실측")

    # ① S710 — 급수방식을 안 고르면 막힌다
    try:
        merge_network(_G(), mode=None)
        check("급수방식 미선택은 막힌다", False, "통과해 버렸다")
    except MergeError:
        check("급수방식 미선택은 막힌다", True)

    # ④ 계통도 없이 — 평면도 단독으로 지나간다
    got = merge_network(_G(), mode="lsp_gravity")
    s = combined_summary(got)
    check("계통도 없으면 평면도 단독", s["merged"] is False,
          " · ".join(got["steps"]))
    check("그래도 기준점은 10 이 된다",
          any(n["label"] == "10" for n in got["head_tables"].nodes))

    # ② S720 — 실도면에서 입상관
    dxf = next((p for p in SYSTEM_DXF if p.is_file()), None)
    if dxf is None:
        check("계통도 실도면", False, "후보 없음 — 결합 실측 생략")
        return 1 if FAILS else 0
    print(f"\n  계통도: {dxf.name}")
    ents, _ = parse_subdrawing(dxf)
    from routes.module_f.subdrawing import entities_to_world
    w = entities_to_world(ents)
    pts = [p for s2 in w.segs for p in (s2[2], s2[3])]
    lo = min(pts, key=lambda p: (p[0], p[1]))
    hi = max(pts, key=lambda p: (p[0], p[1]))
    try:
        riser = extract_system(ents, lo, hi, snap_tolerance_mm=5000)
    except ValueError as exc:
        check("계통도 경로 추출", False, str(exc))
        return 1
    n_riser = len(riser["nodes"])
    check("S720 입상관 추출", n_riser > 1,
          f"절점 {n_riser} · 배관 {len(riser['pipes'])}")
    check("입상관 AV 가 10", str(riser.get("av_node_label")) == "10",
          str(riser.get("av_node_label")))

    # ③ S740 — 결합
    ht_only = to_head_tables(_G())
    got = merge_network(_G(), riser=riser, mode="lsp_gravity")
    s = combined_summary(got)
    check("S740 결합", s["merged"] is True, " · ".join(s["steps"]))
    expect = n_riser + len(ht_only.nodes) - 1        # AV(10) 는 공통 절점
    check("절점 = 입상관 + 헤드망 − 1 (기준점 공유)", s["nodes"] == expect,
          f"{s['nodes']} (기대 {expect})")
    check("노즐이 보존된다", s["nozzles"] == len(ht_only.nozzles),
          f"{s['nozzles']}")

    # 급수방식 4종이 모두 돈다
    for mode in ("hsp_pump", "lsp_gravity", "lsp_1stage", "llsp_2stage"):
        try:
            g2 = merge_network(_G(), riser=riser, mode=mode)
            check(f"급수방식 {mode}", combined_summary(g2)["merged"] is True)
        except Exception as exc:  # noqa: BLE001
            check(f"급수방식 {mode}", False, f"{type(exc).__name__}: {exc}")

    # ⑤⑥⑦ S750 · S760 · S770
    tmp = Path(tempfile.mkdtemp(prefix="mf_merge_"))
    try:
        files = emit_merged(got["combined"], tmp, title="통합 검증")
        for w2 in files.get("warnings") or ():
            print(f"       ! {w2}")
        check("S750 SDF", bool(files["sdf"]) and Path(files["sdf"]).is_file())
        check("SLF 동봉 (호칭경 대조 자료)", bool(files["slf"]),
              "없음 — PIPENET 에서 관경 Unset 위험" if not files["slf"] else "")
        check("S760 KFP", bool(files["kfp"]))
        check("S760 HAS", bool(files["has"]))
        z = Path(files["zip"])
        check("S770 압축", z.is_file(), f"{z.stat().st_size:,} bytes")
        import zipfile
        with zipfile.ZipFile(z) as zf:
            names = zf.namelist()
        check("압축 안에 형식별 파일", len(names) >= 3, " · ".join(names))

        cc = cross_check(files)
        detail = " · ".join(
            (f"{k}: 연장{v.get('total_m')}m·노즐{v.get('nozzles')}"
             f"(절점{v.get('nodes')})")
            if "error" not in v else f"{k}: {v['error'][:40]}"
            for k, v in cc["per_format"].items())
        check("S760 — 세 형식이 같은 배관망", cc["agree"] is True,
              detail + (f" · {cc['detail']}" if cc.get("detail") else ""))
        check("세 형식 모두 읽혔다", len(cc["compared"]) == 3,
              " · ".join(cc["compared"]))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 제5국면 전 항목 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
