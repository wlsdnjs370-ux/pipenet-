# -*- coding: utf-8 -*-
"""모듈 F 자동 차선의 «시간이 어디로 가나» — 최적화 전에 재는 자.

리팩터는 줄 수를 줄이지만 최적화는 시간을 줄인다. 어디가 느린지 모르고 손대면
안 느린 데를 고치게 된다. 그래서 자동 차선의 단계마다 벽시계를 찍는다:

    열기(캐시 적중) · 배관 지정 반영 · 헤드 검출 · 망 검출 · 이음자리

「배관 지정 반영」(`apply_pipe_overrides`)은 자동의 **모든** 입구에서 돈다.
찍은 묶음이 있으면 entity 전체를 다시 만드는 O(n) 이라, 큰 도면에서 요청마다
얹히는 값을 재 둔다.

    python scripts/_probe_module_f_hot.py [도면.dxf ...] [--reps 3]
"""
from __future__ import annotations

import argparse
import statistics
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
]


def timed(fn, reps):
    ts = []
    out = None
    for _ in range(reps):
        t0 = time.perf_counter()
        out = fn()
        ts.append((time.perf_counter() - t0) * 1000)
    return out, statistics.median(ts)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="*")
    ap.add_argument("--reps", type=int, default=3)
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    from routes.module_f.auto import (_bundle_key, apply_pipe_overrides,
                                      detect_head_candidates, junction_marks,
                                      parse_plan, run_network)

    targets = [Path(x) for x in a.dxf] or DEFAULT
    for dxf in targets:
        if not dxf.is_file():
            print(f"■ {dxf.name} — 파일 없음\n")
            continue
        print(f"■ {dxf.name}")

        (ents, cat, diag), t_open = timed(lambda: parse_plan(dxf), a.reps)
        print(f"    entity {len(ents):,} · 레이어 {len(cat):,}")
        print(f"    {'열기(캐시 적중)':<20} {t_open:>9.1f} ms")

        # 배관 지정 없음 — 이른 반환 경로
        _, t_ov0 = timed(lambda: apply_pipe_overrides(ents, cat, None), a.reps)
        print(f"    {'배관 지정 (없음)':<20} {t_ov0:>9.1f} ms")

        # 배관 지정 1묶음 — entity 전체를 다시 만드는 경로
        first = None
        for e in ents:
            k = _bundle_key(e)
            if k:
                first = k
                break
        picks = [{"layer": first[0], "color": first[1]}] if first else []
        _, t_ov1 = timed(lambda: apply_pipe_overrides(ents, cat, picks), a.reps)
        flag = "   ★요청마다 얹힌다" if t_ov1 > 20 else ""
        print(f"    {'배관 지정 (1묶음)':<20} {t_ov1:>9.1f} ms{flag}")

        heads, t_head = timed(lambda: detect_head_candidates(ents, cat), a.reps)
        hp = [(h["x"], h["y"]) for h in (heads or ())]
        print(f"    {'헤드 검출':<20} {t_head:>9.1f} ms   후보 {len(hp):,}")
        if not hp:
            print()
            continue

        cx = sum(q[0] for q in hp) / len(hp)
        cy = sum(q[1] for q in hp) / len(hp)
        al = min(hp, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
        try:
            net, t_net = timed(
                lambda: run_network(ents, cat, alarm_xy=al, rects=None,
                                    prune=False), 1)
            segs = (net or {}).get("segs") or []
            print(f"    {'망 검출':<20} {t_net:>9.1f} ms   선분 {len(segs):,}")
            if segs:
                _, t_j = timed(lambda: junction_marks(segs), 1)
                print(f"    {'이음자리 판정':<20} {t_j:>9.1f} ms")
        except Exception as exc:  # noqa: BLE001
            print(f"    망 검출 실패 — {type(exc).__name__}: {exc}")
        print()

    print("  «배관 지정» 이 크면 세션에 캐시할 값어치가 있다 —")
    print("  같은 묶음으로 여러 요청이 잇달아 들어오기 때문이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
