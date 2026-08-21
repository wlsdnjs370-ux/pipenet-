# -*- coding: utf-8 -*-
"""모듈 F 사전 검증 — 모듈 E 파이프라인이 Qt 없이 도는지 확인한다.

웹 인터페이스를 짜기 전에, 찍기·손질·변환 3단이 헤드리스로 실제 값을 내는지
먼저 본다. 여기서 안 돌면 웹으로 감싸도 안 돈다.
"""
from __future__ import annotations

import os
import sys
import time

_ROOT = os.path.dirname(os.path.abspath(__file__))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

WORK = os.path.join(_ROOT, "docs", "import")


def patch_paths():
    """상대경로(docs/import)를 절대경로로 고정 — 서버 cwd 와 무관하게."""
    from services.cad_import.pipeline import disp_cache, handoff
    handoff.import_write_root = lambda: WORK
    handoff.OUT_DIR = handoff.pick_out_dir()
    disp_cache._DISP_CACHE_DIR = WORK
    return handoff


def check_no_qt():
    bad = [m for m in sys.modules if m.startswith("PySide6")]
    print("PySide6 로드됨:", bad or "없음")
    return not bad


def stage_pick(dxf):
    from services.cad_import.pick.session import PickSession
    t0 = time.perf_counter()
    s = PickSession.open(dxf)
    dt = time.perf_counter() - t0
    w = s.world
    print(f"[찍기] key={s.key} 파싱 {dt:.2f}s "
          f"segs={len(w.segs)} circles={len(w.circles)} arcs={len(w.arcs)}")
    bundles = sorted({(ly, c) for ly, c, _a, _b in w.segs})
    print(f"[찍기] 레이어×색 묶음 {len(bundles)}종, 상위 5: {bundles[:5]}")
    # 실제 선분 위 한 점을 찍어 본다 — 클릭 판정이 헤드리스에서 도는지.
    s.select_pipe()
    ly, c, a, b = w.segs[0]
    mid = ((a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0)
    rep = s.click(mid[0], mid[1])
    print(f"[찍기] 클릭 결과: {rep}")
    print(f"[찍기] 완료가능={s.board.complete_materials()} spec keys={list(s.spec())}")
    hl = s.highlight_geom()
    print(f"[찍기] 강조 pipe_segs={len(hl['pipe_segs'])} "
          f"head_circles={len(hl['head_circles'])}")
    return True


def stage_edit_convert(key):
    from services.cad_import.convert.engine import convert_to_kfp, ensure_planar
    from services.cad_import.edit.session import EditSession
    t0 = time.perf_counter()
    es = EditSession.open(key, out_dir=None, load_saved=True, use_cache=True)
    print(f"[손질] open {time.perf_counter() - t0:.2f}s "
          f"pts={len(es.board.pts)} edges={len(es.board.edges)} "
          f"disks={len(es.board.disks)} sources={es.board.sources} "
          f"valves={es.board.valves}")
    geom = es.display_geom()
    print(f"[손질] body_groups={len(geom['body_groups'])} "
          f"heads={len(geom['heads'])} sources={len(geom['sources'])} "
          f"valves={len(geom['valves'])}")
    state = es.flow()
    if state:
        print(f"[손질] 물흐름 wet_heads={len(state['wet_heads'])}"
              f"/{state['total_heads']} wet_edges={len(state['wet_edges'])}")
    else:
        print("[손질] 급수원이 없어 물흐름 생략")

    payload = es.convert_payload()
    print(f"[변환] payload keys={sorted(payload)[:12]} …")
    t0 = time.perf_counter()
    payload = ensure_planar(payload)
    print(f"[변환] 평면그래프 {time.perf_counter() - t0:.2f}s "
          f"kfp={'있음' if payload.get('kfp') else payload.get('_planar_error')}")
    out = os.path.join(_ROOT, "_smoke_out.kfp")
    t0 = time.perf_counter()
    res = convert_to_kfp(payload, out)
    print(f"[변환] {time.perf_counter() - t0:.2f}s ok={res['ok']} "
          f"blockers={[b.get('code') for b in res['blockers']]}")
    if res["ok"]:
        kfp = res["kfp"]
        print(f"[변환] nodes={len(kfp.get('nodes') or [])} "
              f"pipes={len(kfp.get('pipes') or [])} "
              f"stats={res['stats']}")
        print(f"[변환] 파일 {os.path.getsize(out)} bytes")
    else:
        for b in res["blockers"]:
            print("   막힘:", b)
    return res["ok"]


if __name__ == "__main__":
    patch_paths()
    key = "B1F 현장조사 소화설비 평면도"
    dxf = r"C:\Users\admin\Desktop\B1F 현장조사 소화설비 평면도.dxf"
    ok_edit = stage_edit_convert(key)
    ok_pick = stage_pick(dxf) if os.path.isfile(dxf) else print("DXF 없음 — 찍기 생략")
    print("Qt 미사용:", check_no_qt())
    print("RESULT edit/convert =", ok_edit, "| pick =", ok_pick)
