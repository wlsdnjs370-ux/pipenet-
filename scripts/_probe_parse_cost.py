# -*- coding: utf-8 -*-
"""도면 «보여주기» 를 어느 파서로 할 것인가 — 비용을 재고 정한다.

불러오기 뒤 방식을 묻는 사이에도 도면은 화면에 있어야 한다. 그런데 자동과
수동은 파서가 다르다. 「보여주기」에 어느 것을 쓰느냐에 따라 나중에 한 번 더
읽어야 하는 쪽이 갈린다.

    parse_dxf_bundle    자동이 그대로 쓴다 (entities + 레이어 분류)
    parse_dxf_for_view  계통도·기계실이 쓰는 시각화 파서
    PickSession.open    수동(E)이 쓴다 (World + board)

    python scripts/_probe_parse_cost.py [도면.dxf]
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / "samples" / "dxf" / "LH306동_평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from routes.module_f.common import _boot
    _boot()

    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1
    print(f"도면 {dxf.name} · {dxf.stat().st_size / 1024 / 1024:.1f} MB\n")

    from remote30_prototype import parse_dxf_bundle, parse_dxf_for_view

    t = time.perf_counter()
    b = parse_dxf_bundle(dxf)
    t_bundle = time.perf_counter() - t
    print(f"  parse_dxf_bundle    {t_bundle:6.2f}s · 도형 {len(b.entities):,}")

    t = time.perf_counter()
    v = parse_dxf_for_view(dxf, include_hidden_layers=True)
    t_view = time.perf_counter() - t
    print(f"  parse_dxf_for_view  {t_view:6.2f}s · 도형 "
          f"{len(v.get('entities') or ()):,}")

    from services.cad_import.pick.session import PickSession
    t = time.perf_counter()
    ps = PickSession.open(str(dxf))
    t_pick = time.perf_counter() - t
    w = ps.world
    print(f"  PickSession.open    {t_pick:6.2f}s · 선분 {len(w.segs):,} · "
          f"원 {len(w.circles):,}")

    print("\n선택지")
    print(f"  ① bundle 로 보여주기 → 자동은 0 추가 · 수동은 +{t_pick:.1f}s "
          f"(합 {t_bundle + t_pick:.1f}s)")
    print(f"  ② pick 으로 보여주기 → 수동은 0 추가 · 자동은 +{t_bundle:.1f}s "
          f"(합 {t_pick + t_bundle:.1f}s)")
    print(f"  ③ view 로 보여주기   → 둘 다 추가 "
          f"(자동 +{t_bundle:.1f}s · 수동 +{t_pick:.1f}s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
