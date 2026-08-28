# -*- coding: utf-8 -*-
"""api_edit.py 의 라우트 앞머리를 `route_session` 으로 옮긴다 — 기계적으로.

같은 블록이 열한 번 나온다. 손으로 열한 번 고치면 한 번은 틀린다. 그래서
프로그램이 한다 — 그리고 무엇을 바꿨는지 세어서 보여 준다.

바꾸는 것:
    @app.post(...)                          @app.post(...)
    def f():                        →       @route_session(_edit_session, post=True)
        <앞머리 5~7줄>                       def f(sess, body):
        <본문>                                   es = sess["edit"]      ← es 를 쓰면
                                                 <본문>

    python scripts/_refactor_edit_routes.py [--check]
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "routes" / "module_f" / "api_edit.py"

GUARD_ES = """        body = request.get_json(silent=True) or {}
        try:
            sess, es, bad = _edit_session(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if bad:
            return bad
"""
GUARD_PLAIN = """        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
"""
GUARD_GET = """        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
"""
DEF = re.compile(r"^(    )def (module_f_\w+)\(\):$", re.M)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()

    text = SRC.read_text(encoding="utf-8").replace("\r\n", "\n")
    # 함수 단위로 자른다 — 라우트 데코레이터가 경계다.
    parts = re.split(r"(?=^    @app\.)", text, flags=re.M)
    out, changed = [], []
    for blk in parts:
        m = DEF.search(blk)
        if not m:
            out.append(blk)
            continue
        name = m.group(2)
        if GUARD_ES in blk:
            guard, deco = GUARD_ES, "    @route_session(_edit_session, post=True)\n"
        elif GUARD_PLAIN in blk:
            guard, deco = GUARD_PLAIN, "    @route_session(_edit_session, post=True)\n"
        elif GUARD_GET in blk:
            guard, deco = GUARD_GET, "    @route_session(_edit_session)\n"
        else:
            out.append(blk)
            continue
        body = blk.replace(guard, "")
        body = body.replace(f"    def {name}():\n",
                            deco + f"    def {name}(sess, body):\n")
        # 앞머리를 걷어내면 `es` 를 정의하던 줄도 사라진다 — 쓰는 함수에만 되살린다.
        rest = body.split(f"def {name}(sess, body):\n", 1)[1]
        if re.search(r"\bes\b", rest):
            # 도크스트링이 있으면 그 뒤에 넣는다.
            dm = re.match(r'(\s*""".*?"""\n)', rest, re.S)
            head = dm.group(1) if dm else ""
            tail = rest[len(head):]
            rest2 = head + '        es = sess["edit"]\n' + tail
            body = body.split(f"def {name}(sess, body):\n", 1)[0] \
                + f"def {name}(sess, body):\n" + rest2
        out.append(body)
        changed.append(name)

    new = "".join(out)
    print(f"■ 바꾼 라우트 {len(changed)}개")
    for n in changed:
        print(f"    {n}")
    print(f"\n  {len(text.splitlines()):,}줄 → {len(new.splitlines()):,}줄 "
          f"({len(text.splitlines()) - len(new.splitlines()):+,})")
    if a.check:
        print("  --check 라 쓰지 않았다.")
        return 0
    SRC.write_text(new, encoding="utf-8", newline="\n")
    print(f"  썼다 — {SRC.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
