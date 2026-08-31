# -*- coding: utf-8 -*-
"""찍기 묶음 «접점» 힌트 — 만들기 전에 갈리는지 잰다.

BLOCKED §27 이 남긴 다음 수다. 선분 길이로는 배관과 격자선이 안 갈렸지만
(`_probe_bundle_role_calib.py`), **연결성**은 갈렸다 — 계통도의 층 구획선 `4`
는 큰 덩이를 51→331 노드로 키우면서 원래 배관 노드를 **0개** 데려온다.

여기서 재는 것은 그보다 싼 형태다:

    「이 묶음의 끝점이 **지금 배관으로 찍은 것**과 몇 군데나 맞닿나」

묶음마다 그래프를 다시 세우지 않는다. 씨앗의 끝점을 격자 해시에 한 번 넣고,
후보 끝점을 조회할 뿐이라 전체가 선분 수에 선형이다.

★평가는 **떼어놓고 맞히기**로 한다. 이름 사전이 «맞는» 도면에서 진짜 배관
  레이어를 하나 빼고 씨앗을 만든 뒤, 그 뺀 것이 접점 순위 1등으로 돌아오는지
  본다. 그게 안 되면 이 지표도 못 쓴다 — 계통도에서만 그럴듯한 것은 우연일 수
  있기 때문이다.

    python scripts/_probe_bundle_touch_calib.py [도면.dxf ...]
"""
from __future__ import annotations

import math
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = [
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 계통도.dxf",
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 기계실.dxf",
    ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf",
    ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
]
EPS_MM = 50.0          # 끝점이 «맞닿았다» 고 볼 거리


def segs_of(en) -> list:
    t = str(en.get("t") or "")
    if t == "L":
        x1, y1, x2, y2 = en["p"]
        return [((x1, y1), (x2, y2))]
    if t == "PL":
        pts = en["p"]
        return [((a[0], a[1]), (b[0], b[1])) for a, b in zip(pts, pts[1:])]
    return []


class Hash:
    """끝점 격자 해시 — 셀 크기 = eps. 이웃 9칸만 본다."""

    def __init__(self, eps=EPS_MM):
        self.eps = eps
        self.g = defaultdict(list)

    def add(self, x, y):
        self.g[(int(x // self.eps), int(y // self.eps))].append((x, y))

    def near(self, x, y) -> bool:
        gx, gy = int(x // self.eps), int(y // self.eps)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for (px, py) in self.g.get((gx + dx, gy + dy), ()):
                    if math.hypot(px - x, py - y) <= self.eps:
                        return True
        return False


def touches(seed_pts, cand_segs) -> int:
    """후보 선분의 끝점 중 씨앗과 맞닿는 것의 수."""
    h = Hash()
    for (x, y) in seed_pts:
        h.add(x, y)
    n = 0
    for (a, b) in cand_segs:
        if h.near(a[0], a[1]) or h.near(b[0], b[1]):
            n += 1
    return n


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    for dxf in [Path(x) for x in sys.argv[1:]] or DEF:
        if not dxf.is_file():
            print(f"\n■ {dxf.name} — 파일 없음")
            continue
        bundle = A.parse_dxf_bundle_cached(dxf)
        layers = {ly.get("name"): (ly.get("auto_category") or "OTHER")
                  for ly in (bundle.layers or [])}
        by = defaultdict(list)
        for en in bundle.entities:
            ss = segs_of(en)
            if ss:
                by[str(en.get("l") or "0")] += ss

        pipe_ly = [ly for ly in by if layers.get(ly) == "PIPE"]
        print(f"\n{'=' * 86}")
        print(f"■ {dxf.name}")
        print(f"   PIPE 로 분류된 레이어 {len(pipe_ly)}개 — {pipe_ly}")
        print("=" * 86)

        def rank(seed_lys, tag, expect=None):
            seed = [p for ly in seed_lys for s in by[ly] for p in s]
            if not seed:
                print(f"   {tag} — 씨앗이 비어 건너뜀")
                return
            rows = []
            for ly, ss in by.items():
                if ly in seed_lys or len(ss) < 10:
                    continue
                t = touches(seed, ss)
                rows.append((t / len(ss), t, len(ss), ly))
            rows.sort(reverse=True)
            print(f"\n   {tag}  (씨앗 끝점 {len(seed):,})")
            print(f"     {'레이어':<20}{'분류':>8}{'선분':>8}{'접점':>8}{'비율':>8}")
            for i, (r, t, n, ly) in enumerate(rows[:8]):
                mark = ""
                if expect and ly == expect:
                    mark = f"   ★뺀 것 — 순위 {i + 1}"
                print(f"     {ly[:19]:<20}{layers.get(ly, 'OTHER'):>8}"
                      f"{n:>8,}{t:>8,}{r * 100:>7.0f}%{mark}")
            if expect and expect not in [x[3] for x in rows[:8]]:
                pos = next((i + 1 for i, x in enumerate(rows)
                            if x[3] == expect), None)
                print(f"     ★뺀 것 «{expect}» 는 순위 {pos} — 8위 밖")

        # ⑴ 있는 그대로 — PIPE 를 씨앗으로.
        if pipe_ly:
            rank(set(pipe_ly), "⑴ PIPE 를 씨앗으로")

        # ⑵ 떼어놓고 맞히기 — 가장 큰 PIPE 레이어를 빼고 되찾는지 본다.
        if len(pipe_ly) >= 2:
            held = max(pipe_ly, key=lambda ly: len(by[ly]))
            rank(set(pipe_ly) - {held}, f"⑵ «{held}» 를 빼고 씨앗", expect=held)
    print("\n  판정 기준 — 뺀 배관이 1등으로 돌아오면 쓸 만하다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
