# -*- coding: utf-8 -*-
"""r30_combined 의 region_auto 가 응답 시점까지 «같은 함수 안» 인가.

파이썬은 이름이 없으면 실행할 때 터진다 — 라우트 하나가 통째로 500 이 된다.
자리만 눈으로 맞추지 말고 AST 로 확인한다.

    python scripts/_verify_r30_scope.py
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "routes" / "r30_combined.py"

MARKS = [
    ("region_auto 초기화", "region_auto = False"),
    ("영역 자동 생성",     "if alarm_xy and not zones:"),
    ("경로 판정",          "use_anchored = bool(alarm_xy) and bool(zones)"),
    ("응답 anchored",      '"anchored": use_anchored,'),
    ("응답 region_auto",   '"region_auto": region_auto,'),
]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    src = TARGET.read_text("utf-8")
    tree = ast.parse(src)

    fns = [n for n in ast.walk(tree)
           if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]

    def owner(line: int) -> str:
        cand = [f for f in fns if f.lineno <= line <= (f.end_lineno or f.lineno)]
        if not cand:
            return "<모듈>"
        return min(cand, key=lambda f: (f.end_lineno or f.lineno) - f.lineno).name

    seen, fail = [], 0
    for label, needle in MARKS:
        if needle not in src:
            print(f"  없음  {label} — {needle!r}")
            fail += 1
            continue
        line = src[:src.index(needle)].count("\n") + 1
        fn = owner(line)
        seen.append(fn)
        print(f"  line {line:5d}  {fn:34s}  {label}")

    if len(set(seen)) == 1 and not fail:
        print(f"\nPASS — 다섯 자리 모두 {seen[0]}() 안")
        return 0
    print(f"\nFAIL — 함수가 갈렸다: {sorted(set(seen))}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
