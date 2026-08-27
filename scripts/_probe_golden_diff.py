# -*- coding: utf-8 -*-
"""golden 과 지금 출력이 «무엇이» 다른지만 뽑는다.

특성화 테스트는 「달라졌다」까지만 말한다. 다시 기록할지 되돌릴지 정하려면
어느 값이 얼마나 움직였는지를 봐야 한다.

    python scripts/_probe_golden_diff.py combined_build__plane_daemyeong
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def walk(a, b, path=""):
    """두 JSON 을 나란히 걸으며 다른 잎만 낸다."""
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            yield from walk(a.get(k, "<없음>"), b.get(k, "<없음>"),
                            f"{path}.{k}" if path else str(k))
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            yield (f"{path}[]", f"{len(a)}개", f"{len(b)}개")
        for i, (x, y) in enumerate(zip(a, b)):
            yield from walk(x, y, f"{path}[{i}]")
    elif a != b:
        yield (path, a, b)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    name = sys.argv[1]

    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    gold = json.loads(
        (ROOT / "tests" / "characterization" / "golden" / f"{name}.json")
        .read_text("utf-8"))

    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "_golden_cases",
        ROOT / "tests" / "characterization" / "golden_cases.py")
    gc = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gc)
    CASES = gc.CASES
    fn = CASES.get(name)
    if fn is None:
        print("케이스 없음:", name)
        print("있는 것:", ", ".join(sorted(CASES)))
        return 1
    now = json.loads(json.dumps(fn(), ensure_ascii=False, sort_keys=True))

    diffs = list(walk(gold, now))
    print(f"{name} — 다른 잎 {len(diffs)}개\n")
    for path, g, n in diffs[:120]:
        print(f"  {path}")
        print(f"      golden {g!r}")
        print(f"      지금   {n!r}")
    if len(diffs) > 120:
        print(f"  … 외 {len(diffs) - 120}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())
