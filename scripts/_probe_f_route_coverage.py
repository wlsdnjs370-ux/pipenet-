# -*- coding: utf-8 -*-
"""모듈 F 라우트 «시험 통과 여부» — 오류가 숨을 수 있는 자리를 센다.

리팩터링에서 가장 위험한 것은 **시험이 한 번도 안 지나간 코드**를 건드리는
것이다. 그런 자리는 옮기다 이름 하나만 어긋나도 `NameError` 가 «실행할 때» 나고,
등록만 보는 라우트 인벤토리 시험은 그것을 못 잡는다(이 저장소에서 실제로 겪은
종류다 — 자유이름이 스코프를 옮기면 정적으로는 멀쩡해 보인다).

그래서 쪼개기 전에 **어느 라우트가 시험에 덮이나** 를 먼저 잰다.

방법: 앱을 먼저 띄워 `view_functions` 를 세는 껍데기로 감싼 뒤 pytest 를
같은 프로세스에서 돌린다. 시험이 `대조 서버` 를 다시 import 해도 파이썬이
모듈을 캐시하므로 같은 앱 객체를 쓴다.

    python scripts/_probe_f_route_coverage.py [pytest 인자…]
"""
from __future__ import annotations

import os
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    app = srv.app

    hits: Counter = Counter()
    for name, fn in list(app.view_functions.items()):
        def wrap(fn=fn, name=name):
            def inner(*a, **k):
                hits[name] += 1
                return fn(*a, **k)
            inner.__name__ = getattr(fn, "__name__", name)
            return inner
        app.view_functions[name] = wrap()

    import pytest
    args = sys.argv[1:] or ["tests", "-q", "-x", "--no-header"]
    print(f"■ pytest {' '.join(args)} — 라우트 호출을 세는 중…\n", flush=True)
    code = pytest.main(args)

    # 모듈 F 라우트만 추린다.
    rules = {}
    for r in app.url_map.iter_rules():
        if "/api/module-f/" in r.rule:
            rules[r.endpoint] = r.rule
    covered = sorted(e for e in rules if hits[e])
    naked = sorted(e for e in rules if not hits[e])

    print(f"\n{'=' * 74}")
    print(f"■ 모듈 F 라우트 {len(rules)}개 — 시험이 부른 것 {len(covered)} · "
          f"안 부른 것 {len(naked)}")
    print("=" * 74)
    print(f"\n   ★시험이 한 번도 안 부른 라우트 {len(naked)}개")
    for e in naked:
        print(f"      {rules[e]:<44} {e}")

    # ★«부른 것» 을 «덮은 것» 으로 읽으면 안 된다.
    #
    #   전 라우트 두드리기(`tests/test_module_f_routes.py`)는 세션 없이 부르므로
    #   대부분 `route_session` 이 410 으로 막고 **핸들러 본문은 안 지난다.**
    #   그래서 «한 번만 불린» 라우트는 사실상 그 얕은 두드리기 하나뿐일 수
    #   있다 — 그것을 가려 준다.
    only_smoke = sorted(e for e in rules if hits[e] == 1)
    print(f"\n   ⚠ 딱 한 번만 불린 라우트 {len(only_smoke)}개")
    print("     (전수 두드리기 한 번뿐일 가능성이 높다 — 본문은 아직 안 지났을 수 있다)")
    for e in only_smoke:
        print(f"      {rules[e]}")
    print(f"\n   부른 것 (호출 수 순)")
    for e, n in hits.most_common():
        if e in rules:
            print(f"      {n:>5}회  {rules[e]}")
    print(f"\n   pytest 종료코드 {code}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
