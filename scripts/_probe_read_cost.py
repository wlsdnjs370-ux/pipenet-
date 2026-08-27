# -*- coding: utf-8 -*-
"""「자동/수동으로 읽기」 뒤의 기다림이 어디서 오는가 — 재고 정한다.

두 갈래가 각각 무엇을 하는지:

    수동  더 읽을 것이 없다 «고 했는데», 화면이 /world 를 한 번 더 받는다.
          이미 메모리에 있는 것을 다시 내려받는 셈이다.
    자동  A 의 파서로 한 번 더 읽는다. A 에는 파일 해시로 디스크 캐시하는
          parse_dxf_bundle_cached 가 있는데 안 쓰고 있었다.

    python scripts/_probe_read_cost.py [도면.dxf]
"""
from __future__ import annotations

import json
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

    # ── 수동: /world 를 다시 받는 값
    from services.cad_import.pick.session import PickSession
    from routes.module_f.world import _world_payload
    t = time.perf_counter()
    ps = PickSession.open(str(dxf))
    t_pick = time.perf_counter() - t
    pay = _world_payload(ps.world)
    blob = json.dumps(pay, ensure_ascii=False)
    t = time.perf_counter()
    for _ in range(3):
        json.dumps(pay, ensure_ascii=False)
    t_ser = (time.perf_counter() - t) / 3
    print("수동 — 「읽기」 뒤에 실제로 드는 것")
    print(f"  PickSession.open       {t_pick:6.2f}s   (열기 때 이미 끝났다)")
    print(f"  /world 응답 크기        {len(blob) / 1024 / 1024:6.2f} MB")
    print(f"  그것을 다시 직렬화      {t_ser:6.2f}s   ← 안 해도 되는 일")
    print(f"  + 브라우저가 그만큼 받아 파싱하고 buildLayers 를 다시 돌린다")

    # ── 자동: A 파서 — 캐시 없이 / 있고
    from remote30_prototype import parse_dxf_bundle, parse_dxf_bundle_cached
    t = time.perf_counter()
    parse_dxf_bundle(dxf)
    t_cold = time.perf_counter() - t
    t = time.perf_counter()
    parse_dxf_bundle_cached(dxf)
    t_warm1 = time.perf_counter() - t
    t = time.perf_counter()
    parse_dxf_bundle_cached(dxf)
    t_warm2 = time.perf_counter() - t
    print("\n자동 — A 파서")
    print(f"  parse_dxf_bundle          {t_cold:6.2f}s   (지금 쓰는 것)")
    print(f"  parse_dxf_bundle_cached   {t_warm1:6.2f}s   (첫 번 — 캐시를 만든다)")
    print(f"  parse_dxf_bundle_cached   {t_warm2:6.2f}s   ← 두 번째부터")
    if t_cold > 0:
        print(f"  같은 도면을 다시 열면 {t_cold / max(t_warm2, 1e-6):.0f}배 빠르다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
