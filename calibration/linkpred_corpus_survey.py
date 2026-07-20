# -*- coding: utf-8 -*-
"""link-prediction 학습 코퍼스 확대 후보 인벤토리.

v2 는 스프링클러 REMOTE 43건만 학습한다. 손상→후보→피처 파이프라인은 기하-only·
시스템 무관이라, 답안 SDF 의 다른 유형(옥내소화전·자연감압·ESFR·K-160·USER 전체망)
을 넣으면 위상 다양성이 오른다. 단 각 파일이 학습에 쓸 만한가:
  · 유효 위상(노드>0·파이프>0·start/end 매칭)
  · 차수≥2 junction (un-snap break 양성 생성 가능)
  · 차수3 junction (T-tap 합성 양성 생성 가능)
  · 좌표 spread (degenerate=동일점 붕괴 도면 배제)
유형 키워드별로 집계한다.

실행:
    python calibration/linkpred_corpus_survey.py
"""
from __future__ import annotations

import glob
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

from kfp_sdf_converter import parse_sdf  # noqa: E402

GLOB = str(_ROOT / "수리계산 참고용 도서" / "**" / "*.sdf")

TYPE_KEYS = [
    "옥내소화전", "자연감압", "물분무", "ESFR", "K-160", "K160",
    "드렌처", "포소화", "스프링클러",
]


def _type_of(name: str) -> str:
    for k in TYPE_KEYS:
        if k in name:
            return k
    return "기타"


def _stats(net):
    n_nodes = len(net.nodes)
    n_pipes = len(net.pipes)
    deg = defaultdict(int)
    for p in net.pipes.values():
        deg[p.start] += 1
        deg[p.end] += 1
    d2 = sum(1 for v in deg.values() if v >= 2)
    d3 = sum(1 for v in deg.values() if v == 3)
    d4 = sum(1 for v in deg.values() if v >= 4)
    xs = [n.x for n in net.nodes.values()]
    ys = [n.y for n in net.nodes.values()]
    spread = max(max(xs) - min(xs), max(ys) - min(ys)) if xs else 0.0
    return n_nodes, n_pipes, d2, d3, d4, spread


def main():
    files = sorted(glob.glob(GLOB, recursive=True))
    print("=" * 96)
    print(f"답안 SDF 코퍼스 인벤토리 — 총 {len(files)}건")
    print("=" * 96)

    by_type = defaultdict(lambda: {
        "n": 0, "usable": 0, "with_d3": 0, "pipes": 0, "d3": 0, "fail": 0})
    usable_files = []
    for f in files:
        name = Path(f).name
        is_remote = "REMOTE" in name.upper()
        t = _type_of(name) + ("/REMOTE" if is_remote else "/USER")
        rec = by_type[t]
        rec["n"] += 1
        try:
            net = parse_sdf(f)
            nn, npi, d2, d3, d4, spread = _stats(net)
        except Exception:
            rec["fail"] += 1
            continue
        usable = npi >= 5 and nn >= 5 and spread > 1e-6 and d2 >= 1
        if usable:
            rec["usable"] += 1
            rec["pipes"] += npi
            rec["d3"] += d3
            if d3 >= 1:
                rec["with_d3"] += 1
            usable_files.append((f, t, npi, d3))

    print(f"\n{'유형/구역':28}{'건수':>5}{'사용가능':>7}{'T분기보유':>9}"
          f"{'평균파이프':>9}{'평균차수3':>9}{'파싱실패':>8}")
    print("-" * 96)
    for t in sorted(by_type):
        r = by_type[t]
        u = r["usable"] or 1
        print(f"{t:28}{r['n']:>5}{r['usable']:>7}{r['with_d3']:>9}"
              f"{r['pipes']/u:>9.0f}{r['d3']/u:>9.1f}{r['fail']:>8}")

    print("-" * 96)
    tot_use = len(usable_files)
    tot_d3 = sum(1 for *_x, d3 in usable_files if d3 >= 1)
    print(f"사용가능 합계 {tot_use}건 · 그중 T분기(차수3) 보유 {tot_d3}건")
    return 0


if __name__ == "__main__":
    sys.exit(main())
