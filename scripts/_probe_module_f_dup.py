# -*- coding: utf-8 -*-
"""모듈 F 라우트의 반복 관용구를 «센다» — 리팩터 전에 숫자를 잡는다.

눈으로 「중복이 많아 보인다」는 근거가 못 된다. 실제로 몇 줄이 몇 번
되풀이되는지, 없애면 몇 줄이 주는지 세고 시작한다.

    python scripts/_probe_module_f_dup.py
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PKG = ROOT / "routes" / "module_f"

# 세션 꺼내기 + 실패 응답의 정형 — 이 다섯 줄이 라우트마다 앞머리에 붙는다.
GUARD = re.compile(
    r"[ \t]*try:\n"
    r"[ \t]*(?P<call>\w+(?:, why)? = _\w+\([^\n]*\))\n"
    r"[ \t]*except ValueError as exc:\n"
    r"[ \t]*return _fail\(str\(exc\), 410\)\n"
    r"(?:[ \t]*if why:\n[ \t]*return _fail\(why, 409\)\n)?")


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    files = sorted(PKG.glob("*.py"))
    tot_lines = tot_routes = tot_guard = tot_guard_lines = 0
    rows = []
    calls: Counter = Counter()
    for f in files:
        t = f.read_text(encoding="utf-8")
        n_lines = len(t.splitlines())
        n_routes = len(re.findall(r"@app\.(?:get|post|route)", t))
        g = list(GUARD.finditer(t))
        g_lines = sum(m.group(0).count("\n") for m in g)
        for m in g:
            calls[m.group("call").split("=")[-1].strip().split("(")[0]] += 1
        tot_lines += n_lines
        tot_routes += n_routes
        tot_guard += len(g)
        tot_guard_lines += g_lines
        if n_routes or g:
            rows.append((f.name, n_lines, n_routes, len(g), g_lines))

    print("■ 라우트 앞머리 «세션 꺼내기 + 실패» 관용구\n")
    print(f"  {'파일':<18} {'줄':>6} {'라우트':>7} {'관용구':>7} {'차지한 줄':>9}")
    print("  " + "-" * 52)
    for nm, nl, nr, ng, gl in rows:
        print(f"  {nm:<18} {nl:>6,} {nr:>7} {ng:>7} {gl:>9}")
    print("  " + "-" * 52)
    print(f"  {'합계':<18} {tot_lines:>6,} {tot_routes:>7} "
          f"{tot_guard:>7} {tot_guard_lines:>9}")

    print(f"\n  관용구가 쓰는 헬퍼:")
    for k, v in calls.most_common():
        print(f"    {v:>3}회  {k}")

    # 데코레이터 하나로 바꾸면 관용구 줄이 통째로 사라진다(호출 1줄만 남김).
    saved = tot_guard_lines - tot_guard
    print(f"\n  데코레이터로 걷으면 {tot_guard_lines} → {tot_guard} 줄 "
          f"(**{saved}줄 감소**, 패키지의 {saved / max(1, tot_lines) * 100:.1f}%)")

    # ── 그 밖의 반복 — 완전히 같은 줄이 여러 파일에 몇 번 나오나
    seen: Counter = Counter()
    for f in files:
        for ln in f.read_text(encoding="utf-8").splitlines():
            s = ln.strip()
            if len(s) > 24 and not s.startswith("#"):
                seen[s] += 1
    dup = [(c, s) for s, c in seen.items() if c >= 4]
    dup.sort(reverse=True)
    print(f"\n■ 4번 이상 똑같이 되풀이되는 줄 — 상위 12")
    for c, s in dup[:12]:
        print(f"    {c:>3}회  {s[:78]}")
    print(f"\n    그런 줄 {len(dup)}종 · 합계 {sum(c for c, _ in dup):,}줄")
    return 0


if __name__ == "__main__":
    sys.exit(main())
