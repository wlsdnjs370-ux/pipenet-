# -*- coding: utf-8 -*-
"""유도 스텁 길이의 기준을 고른다 — 무엇에 비례해야 흔들림이 가장 작은가.

방향이 없는 노즐은 길이를 지어내야 한다. 지어낼 값은 방향이 있는 진짜 노즐들의
분포에서 온다. 후보 셋을 같은 표본으로 재서 변동계수로 고른다.

  ① 모델 단위 절댓값        기호(화살표·머리)와 같은 기준
  ② 도면 span 비율          지금 코드가 쓰는 기준
  ③ 그 도면의 관로 길이 대비  도면마다 축척이 달라도 따라가는 기준
"""
from __future__ import annotations

import math
import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def spread(vals: list[float]) -> str:
    v = sorted(vals)
    return (f"중앙 {statistics.median(v):9.4f}  4분위 "
            f"{statistics.quantiles(v, n=4)[0]:8.4f}/{statistics.quantiles(v, n=4)[2]:8.4f}"
            f"  변동계수 {statistics.pstdev(v) / statistics.mean(v):.3f}")


def main() -> None:
    units, ratios, per_pipe, per_file = [], [], [], []
    files = 0
    for sdf in sorted(LIB.rglob("*.sdf")):
        try:
            model = load_display_model(sdf)
        except Exception:                                   # noqa: BLE001
            continue
        coords = {n.label: (n.x, n.y) for n in model.nodes}
        if not model.nozzles:
            continue
        xs = [n.x for n in model.nodes]
        ys = [n.y for n in model.nodes]
        span = max(max(xs) - min(xs), max(ys) - min(ys))
        # 관로 길이는 그 도면의 "한 칸"이다 — 축척이 달라도 같이 커지고 작아진다.
        plen = [math.dist(coords[p.input_node], coords[p.output_node])
                for p in model.pipes
                if p.input_node in coords and p.output_node in coords]
        plen = [d for d in plen if d > 1e-9]
        if span <= 0 or not plen:
            continue
        median_pipe = statistics.median(plen)

        got = []
        for z in model.nozzles:
            base, tip = coords.get(z.input_node), coords.get(z.output_node)
            if base is None or tip is None:
                continue
            d = math.dist(base, tip)
            if d > 1e-9:
                got.append(d)
        if not got:
            continue
        files += 1
        units.extend(got)
        ratios.extend(d / span for d in got)
        per_pipe.extend(d / median_pipe for d in got)
        per_file.append(statistics.median(got))

    print(f"도면 {files}장  방향이 있는 노즐 {len(units)}개")
    print(f"  ① 모델 단위      {spread(units)}")
    print(f"  ② 도면 span 비율 {spread(ratios)}")
    print(f"  ③ 관로 길이 대비 {spread(per_pipe)}")
    print(f"  (도면별 중앙값끼리만) {spread(per_file)}")


if __name__ == "__main__":
    main()
