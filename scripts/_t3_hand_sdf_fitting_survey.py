# -*- coding: utf-8 -*-
"""작업 3 근거 실측 — 실무자가 손으로 만든 PIPENET SDF 의 부속 계상 관행.

새 규칙(직류티 제외)이 현업과 맞는지 보려면 **손으로 만든 망**에서
"분기 노드 수" 대비 "실제로 달린 티 수" 를 세는 것이 직접 증거다.
  · 옛 규칙(차수≥3 전량)   → 티 = 분기 노드의 하류 관로 수 합 (분기당 2~3)
  · 새 규칙(분류티만)      → 티 = 분기당 하류 수 - 1 (직진 1개 제외)

SDF 스키마: <Pipe label input output><Fittings><Fitting type count/></Fittings></Pipe>
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
ROOTS = (REPO / "data/sample_problem", REPO / "assets")


def survey(path: Path) -> str | None:
    try:
        root = ET.parse(path).getroot()
    except Exception as exc:
        return f"{path.name[:52]:54s} 파싱 실패 — {exc}"

    pipes = []
    fittings: Counter[str] = Counter()
    for el in root.iter("Pipe"):
        i, o = el.get("input"), el.get("output")
        if not (i and o):
            continue
        pipes.append((i, o))
        for f in el.iter("Fitting"):
            fittings[f.get("type", "?")] += int(f.get("count") or 1)
    if not pipes:
        return None

    deg: Counter[str] = Counter()
    outs: dict[str, int] = defaultdict(int)
    for i, o in pipes:
        deg[i] += 1
        deg[o] += 1
        outs[i] += 1
    branch = [n for n in deg if deg[n] >= 3]
    old_rule = sum(outs[n] for n in branch)
    new_rule = sum(max(outs[n] - 1, 0) for n in branch)
    return (f"{path.name[:52]:54s} 관로{len(pipes):4d} 분기{len(branch):3d}"
            f" | 손 tee {fittings['tee']:4d}"
            f" | 옛 {old_rule:4d} 새 {new_rule:4d} | {dict(fittings)}")


files = sorted({p for r in ROOTS for p in r.rglob("*.sdf")})
print(f"대상 SDF {len(files)}개\n")
for p in files:
    line = survey(p)
    if line:
        print(line)
