# -*- coding: utf-8 -*-
"""한 파일의 diff 를 훅 단위로 갈라 낸다 — «항목당 커밋 1개» 를 지키기 위한 도구.

작업 지시서는 항목 병합을 금지한다. 그런데 한 파일(`static/module_f.js`)에
서로 다른 항목의 손질이 섞여 쌓이는 일은 흔하다. `git add -p` 는 이 환경에서
대화형이라 못 쓰므로, 훅의 «옛 줄 번호» 로 갈라 두 개의 패치를 만든다.

    python scripts/_split_patch.py <파일> <경계줄> <앞.patch> <뒤.patch>

경계줄 이상에서 시작하는 훅이 «뒤» 로 간다. 되돌릴 때는 `git apply -R 뒤.patch`.
"""
from __future__ import annotations

import re
import subprocess
import sys


def main() -> int:
    path, cut, out_a, out_b = sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4]
    d = subprocess.run(["git", "diff", "-U3", "--", path],
                       capture_output=True, text=True, encoding="utf-8").stdout
    lines = d.splitlines(keepends=True)
    head, hunks, cur = [], [], None
    for ln in lines:
        if ln.startswith("@@"):
            cur = [ln]
            hunks.append(cur)
        elif cur is None:
            head.append(ln)
        else:
            cur.append(ln)

    def old_start(h):
        return int(re.match(r"@@ -(\d+)", h[0]).group(1))

    a = [h for h in hunks if old_start(h) < cut]
    b = [h for h in hunks if old_start(h) >= cut]
    for out, sel in ((out_a, a), (out_b, b)):
        with open(out, "w", encoding="utf-8", newline="") as fh:
            fh.write("".join(head))
            for h in sel:
                fh.write("".join(h))
    print(f"  훅 {len(hunks)}개 → 앞 {len(a)} · 뒤 {len(b)}")
    for h in b:
        print(f"    뒤: {h[0].rstrip()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
