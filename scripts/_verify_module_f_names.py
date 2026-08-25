# -*- coding: utf-8 -*-
"""모듈 F 패키지 — 함수 본문이 쓰는 «자유 이름» 이 그 모듈에 실제로 있나.

라우트를 다른 파일로 옮길 때 나는 사고가 있다: 등록은 되는데 **본문이 쓰던
이름이 새 스코프에 없어** 그 경로를 실제로 부를 때에만 NameError 가 난다.
라우트 인벤토리는 «등록됐나» 만 보므로 이걸 못 잡는다.

여기서는 `symtable` 로 각 모듈의 모든 스코프를 훑어, 전역으로 읽히는 이름 중
그 모듈에 바인딩도 없고 빌트인도 아닌 것을 찾는다. import 만 고치면 되는
종류의 실수를 실행 없이 잡는다.
"""
from __future__ import annotations

import builtins
import io
import os
import symtable
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(ROOT, "routes", "module_f")
BUILTINS = set(dir(builtins)) | {"__file__", "__name__", "__doc__"}

problems: list[str] = []


def walk(table, top_names, path, where):
    for sym in table.get_symbols():
        name = sym.get_name()
        # is_global() 만이 «모듈 전역에서 찾는다» 는 뜻이다. 바깥 함수의 지역을
        # 물려받는 클로저 변수(is_free)는 정상이므로 걸러내야 한다 — 이걸 빼면
        # `sweep()` 이 `pairs` 를 못 찾는다는 식의 거짓 경보가 45건 쏟아진다.
        if not sym.is_global() or name in BUILTINS or name in top_names:
            continue
        # 전역으로 읽는데 이 모듈에 없다 → 부를 때 NameError 가 난다.
        problems.append(f"{path}: {where} 안에서 «{name}» 을 찾을 수 없다")
    for child in table.get_children():
        walk(child, top_names, path, f"{where} > {child.get_name()}()")


def main() -> int:
    files = sorted(f for f in os.listdir(PKG) if f.endswith(".py"))
    print(f"[모듈 F 자유 이름 검사] {len(files)}개 파일\n")
    for f in files:
        p = os.path.join(PKG, f)
        src = io.open(p, encoding="utf-8").read()
        top = symtable.symtable(src, p, "exec")
        top_names = {s.get_name() for s in top.get_symbols()}
        before = len(problems)
        for child in top.get_children():
            walk(child, top_names, f"routes/module_f/{f}", f"{child.get_name()}()")
        mark = "OK  " if len(problems) == before else "FAIL"
        print(f"  [{mark}] {f:<16} 최상위 이름 {len(top_names)}개")
    print()
    if problems:
        for p in problems:
            print("  !!", p)
        print(f"\n자유 이름 {len(problems)}건 — import 를 고쳐야 한다.")
        return 1
    print("자유 이름 문제 없음.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
