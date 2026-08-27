# -*- coding: utf-8 -*-
"""A 화면(remote30_prototype.html)의 JS 를 실제로 파싱해 본다.

템플릿은 서버가 자동 리로드해도 «구문이 깨지면» 화면 전체가 죽는다. 붙여 넣은
줄 하나 때문에 A 가 통째로 멈추는 일을 여기서 먼저 잡는다.

    python scripts/_verify_r30_js.py
"""
from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    html = (ROOT / "templates" / "remote30_prototype.html").read_text("utf-8")
    blocks = re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", html, re.S)
    if not blocks:
        print("script 블록을 못 찾았다")
        return 1

    fails = 0
    for i, src in enumerate(blocks, start=1):
        if not src.strip():
            continue
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False,
                                         encoding="utf-8") as fh:
            fh.write(src)
            tmp = fh.name
        r = subprocess.run(["node", "--check", tmp],
                           capture_output=True, text=True, encoding="utf-8")
        Path(tmp).unlink(missing_ok=True)
        n = src.count("\n") + 1
        if r.returncode:
            fails += 1
            print(f"  블록 {i} ({n:,}줄) — 구문 오류")
            print("   ", (r.stderr or "").strip()[:600])
        else:
            print(f"  블록 {i} ({n:,}줄) OK")

    # 이번에 넣은 줄이 실제로 붙어 있는가 — 구문만 맞고 빠지면 소용없다.
    for needle in ('data.anchored === false', 'data.region_auto',
                   '+ _pathLine +'):
        ok = needle in html
        print(f"  {'OK ' if ok else '없음'} {needle}")
        fails += 0 if ok else 1

    print("PASS" if not fails else f"FAIL — {fails}건")
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
