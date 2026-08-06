# -*- coding: utf-8 -*-
"""비트동일 회귀 검증 — characterization 케이스 실측을 data/_w_baseline/ 과 대조.

golden 파일은 사용자 육안확인 게이트로 재생성이 보류돼 있어 stale 이다(BLOCKED.md #2).
그래서 "작업 시작 시점 HEAD 의 실제 산출"을 baseline 으로 동결하고 그것과 대조한다.

    python scripts/_w_baseline_verify.py          # 대조
    python scripts/_w_baseline_verify.py --freeze # baseline 재동결
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests" / "characterization"))

import golden_cases as gc  # noqa: E402

BASE = REPO / "data" / "_w_baseline"


def _dump(obj) -> str:
    return json.dumps(obj, sort_keys=True, ensure_ascii=False, indent=2)


def main() -> int:
    freeze = "--freeze" in sys.argv
    BASE.mkdir(parents=True, exist_ok=True)
    bad = []
    for name in sorted(gc.CASES):
        try:
            actual = _dump(gc.CASES[name]())
        except Exception as exc:  # 도메인 워크트리 결손 자산 — 전후 동일하면 통과
            actual = _dump({"__error__": f"{type(exc).__name__}: {exc}"})
        f = BASE / f"{name}.json"
        if freeze:
            f.write_text(actual, encoding="utf-8")
            print(f"  FROZEN  {name}")
            continue
        if not f.exists():
            print(f"  MISSING {name}")
            bad.append(name)
        elif f.read_text(encoding="utf-8") != actual:
            print(f"  DIFF    {name}")
            bad.append(name)
        else:
            print(f"  ok      {name}")
    if freeze:
        return 0
    print(f"\n{len(gc.CASES) - len(bad)}/{len(gc.CASES)} BIT-IDENTICAL")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
