# -*- coding: utf-8 -*-
"""C1 사슬 실도면 벤치마크 — 일회용 (§16 8항, 부록 B).

합성 도면은 사슬이 도는지만 말해 준다. 실 도면에서만 드러나는 것이 둘이다.
부록 C.1 예산을 어느 단계에서 넘기는지, 그리고 어느 단계가 결과를 잘게
부수는지. 그래서 단계마다 시간과 한 줄 요약을 같이 찍는다.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from _c130_smoke import entities_of  # noqa: E402

from core.design.recognize import pipeline as PL  # noqa: E402


def main(paths: list[str], wall_layers: list[str]) -> None:
    for raw in paths:
        path = Path(raw)
        t0 = time.perf_counter()
        ents, bbox = entities_of(path)
        parse = time.perf_counter() - t0

        last = None
        rooms = None
        for msg in PL.recognize(ents, bbox, wall_layers=wall_layers):
            if msg["type"] == "rooms":
                rooms = msg["rooms"]
            last = msg

        print(f"\n{path.name}")
        print(f"  엔티티 {len(ents)} / 파싱 {parse:.1f}s / 인식 {last['seconds']:.1f}s "
              f"/ 합계 {parse + last['seconds']:.1f}s")
        print(f"  wall={last['wall_layers']}({last['wall_source']}) "
              f"blocked={last['blocked']} counts={last['counts']}")
        for stage in last["stages"]:
            print(f"    {stage['name']:<10}{stage['seconds']:>7.2f}s  {stage['summary']}")
            for line in stage["provenance"]:
                print(f"        + {line}")
        if rooms:
            top = sorted(rooms, key=lambda r: -r["area_m2"])[:5]
            for room in top:
                print(f"    · {room['area_m2']:8.1f}㎡  {room.get('name') or '(이름 없음)':<12}"
                      f"가상 {room['virtual_edge_ratio']:.0%}  {room['flags']}")


if __name__ == "__main__":
    args = sys.argv[1:]
    split = args.index("--wall") if "--wall" in args else len(args)
    main(args[:split], args[split + 1:])
