# -*- coding: utf-8 -*-
"""자동(A) 경로의 표로 «.sdf + .slf 저장» 이 실제로 되는가.

화면은 그 단추를 자동 경로에서 영영 안 열어 준다(수동 build 안에서만 푼다).
그런데 서버 쪽(`/design/emit` → `emit_design_files`)은 `sess["design"]` 만
있으면 되고, 자동 경로도 그것을 채운다. 그러면 화면만 고치면 되는 것인지,
아니면 G 의 방출기가 A 의 표를 못 받는 것인지 — 여기서 가른다.

    python scripts/_probe_auto_emit.py
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DXF = ROOT / "samples" / "dxf" / "LH306동_평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    from routes.module_f.common import _boot
    _boot()

    from routes.module_f.auto import (detect_head_candidates, parse_plan,
                                      run_auto)

    if not DXF.is_file():
        print("실도면 없음:", DXF)
        return 1
    print(f"자동 추출 — {DXF.name}")
    ents, cat, _diag = parse_plan(DXF)
    heads = detect_head_candidates(ents, cat)
    xs = [h["x"] for h in heads]
    ys = [h["y"] for h in heads]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    mid = min(heads, key=lambda h: (h["x"] - cx) ** 2 + (h["y"] - cy) ** 2)
    pad = 1000.0
    got = run_auto(ents, cat, alarm_xy=(mid["x"], mid["y"]),
                   rects=[[min(xs) - pad, min(ys) - pad,
                           max(xs) + pad, max(ys) + pad]], k=30)
    tbl = got["tables"]
    print(f"  표 — 절점 {len(tbl.nodes)} · 배관 {len(tbl.pipes)} · "
          f"노즐 {len(tbl.nozzles)}")

    # 화면이 하는 것과 같은 세션 모양을 만든다.
    from routes.module_f.api_design import emit_design_files
    from routes.module_f.jobs import _new_session

    sess = _new_session()
    sess["key"] = DXF.stem
    sess["method"] = "auto"
    sess["design"] = {"got": {}, "tables": tbl, "k": 30,
                      "schedule": None, "marks": {}, "method": "auto"}

    tmp = Path(tempfile.mkdtemp(prefix="mf_autoemit_"))
    try:
        out, err = emit_design_files(sess, tmp)
        if err:
            print(f"\n★ 저장 실패 — {err}")
            print("   → 화면만 고쳐서는 안 된다. 방출기가 A 의 표를 못 받는다.")
            return 1
        sdf = Path(out)
        slf = sdf.with_suffix(".slf")
        print(f"\n  SDF {sdf.name} · {sdf.stat().st_size:,} bytes")
        print(f"  SLF {'있음' if slf.is_file() else '없음'}"
              + (f" · {slf.stat().st_size:,} bytes" if slf.is_file() else ""))

        # 정말 그 배관망인가 — 절점·연장으로 견준다.
        from kfp_sdf_converter import parse_sdf
        net = parse_sdf(str(sdf))
        total = sum(float(getattr(p, "length_m", 0.0) or 0.0)
                    for p in net.pipes.values())
        nz = sum(1 for n in net.nodes.values()
                 if str(getattr(n, "kind", "")).lower() in ("nozzle", "head"))
        print(f"  되읽기 — 절점 {len(net.nodes)} · 배관 {len(net.pipes)} · "
              f"노즐 {nz} · 연장 {total:.3f} m")
        ok = nz == len(tbl.nozzles)
        print(f"\n{'PASS' if ok else 'FAIL'} — 자동 표로도 저장이 된다"
              if ok else "\nFAIL — 노즐 수가 표와 다르다")
        return 0 if ok else 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
