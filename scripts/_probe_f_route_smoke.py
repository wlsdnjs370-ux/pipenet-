# -*- coding: utf-8 -*-
"""모듈 F 라우트 전수 «두드리기» — 시험으로 옮기기 전에 위험한 것을 가린다.

라우트 커버리지를 재 보니 63개 중 **15개만** 시험이 지난다
(`_probe_f_route_coverage.py`). 나머지는 옮기다 이름 하나 어긋나도 아무도
모른다 — 등록만 보는 인벤토리 시험은 «실행할 때» 나는 `NameError` 를 못 잡는다.

그래서 전 라우트를 한 번씩 두드리는 시험을 만들려는데, 그 전에 **두드려도
안전한지** 부터 본다. 스트리밍(SSE)·파일 쓰기·무거운 잡이 섞여 있어서
아무 생각 없이 부르면 시험이 멈추거나 파일을 흘린다.

    python scripts/_probe_f_route_smoke.py
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SLOW_MS = 1500          # 이보다 오래 걸리면 시험에 넣기 전에 봐야 한다


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    rules = []
    for r in srv.app.url_map.iter_rules():
        if "/api/module-f/" not in r.rule:
            continue
        for m in ("GET", "POST"):
            if m in r.methods:
                rules.append((m, r.rule, r.endpoint))
    rules.sort(key=lambda t: (t[1], t[0]))

    print(f"■ 모듈 F 라우트 {len(rules)}개를 세션 없이 한 번씩 두드린다")
    print(f"   {'메서드':<6}{'경로':<44}{'상태':>6}{'ms':>7}  비고")
    bad, slow = [], []
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        for meth, rule, ep in rules:
            # 경로 매개변수는 아무 값이나 — 존재하지 않는 키가 정상 응답이다.
            path = rule.replace("<key>", "__none__").replace("<path:", "<")
            path = path.replace("<name>", "__none__")
            if "<" in path:
                path = path.split("<")[0] + "__none__"
            t0 = time.time()
            try:
                if meth == "GET":
                    rv = c.get(path + "?sid=__none__")
                else:
                    rv = c.post(path, json={"sid": "__none__"})
                code = rv.status_code
                note = ""
            except Exception as exc:  # noqa: BLE001 — 던지는 것 자체가 결함이다
                code, note = -1, f"★{type(exc).__name__}: {exc}"
            ms = (time.time() - t0) * 1000
            if code >= 500 or code < 0:
                bad.append((meth, rule, code, note))
            if ms > SLOW_MS:
                slow.append((meth, rule, ms))
            print(f"   {meth:<6}{rule[:43]:<44}{code:>6}{ms:>7.0f}  {note}")

    print(f"\n■ 요약")
    print(f"   500 이상이거나 예외를 던진 것 {len(bad)}개")
    for meth, rule, code, note in bad:
        print(f"      {meth} {rule} → {code} {note}")
    print(f"   {SLOW_MS}ms 넘게 걸린 것 {len(slow)}개")
    for meth, rule, ms in slow:
        print(f"      {meth} {rule} → {ms:.0f}ms")
    if not bad and not slow:
        print("   [OK] 전부 통제된 응답 · 전부 빠르다 — 시험으로 옮겨도 된다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
