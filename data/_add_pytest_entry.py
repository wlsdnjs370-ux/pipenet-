# -*- coding: utf-8 -*-
"""스크립트형 «시험» 4개에 pytest 진입점을 붙인다.

이름은 `test_*.py` 인데 `assert` 없이 `main()` + `FAILS` 로 보고해서 pytest 가
**0건**을 수집했다. 표방하는 것이 「전 구간 골든」·「산출 3종 8조합」처럼 가장
넓은 것들인데 자동 실행에는 없었다.

`main()` 은 그대로 두고(사람이 직접 돌리는 길을 없애지 않는다) 그것을 부르는
시험 함수 하나를 덧붙인다. 반환값 0 이 아니면 실패다.
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

ENTRY = '''

# ─────────────────────────────────────────────────────────────────────────
def test_{name}():
    """pytest 진입점 — `main()` 을 그대로 태운다.

    ★이 파일은 오랫동안 `test_*.py` 이면서 **pytest 수집 0건**이었다.
      `assert` 대신 `FAILS`+`return 1` 로 보고하는 스크립트였기 때문이다.
      「{title}」를 표방하면서 자동 실행에는 없었다 — 안 도는 시험은
      통과가 아니라 없는 것이다.

      `main()` 은 그대로 둔다(사람이 직접 돌리는 길을 없애지 않는다).
      여기서는 그 반환값만 본다.
    """
    import pytest

    if not DXF.exists():
        pytest.skip(f"표본 도면 없음: {{DXF}}")
    assert main() == 0, "실패 상세는 위 출력 참고"
'''

ENTRY_NODXF = '''

# ─────────────────────────────────────────────────────────────────────────
def test_{name}():
    """pytest 진입점 — `main()` 을 그대로 태운다.

    ★이 파일은 오랫동안 `test_*.py` 이면서 **pytest 수집 0건**이었다.
      `assert` 대신 `FAILS`+`return 1` 로 보고하는 스크립트였기 때문이다.
      「{title}」를 표방하면서 자동 실행에는 없었다 — 안 도는 시험은
      통과가 아니라 없는 것이다.
    """
    assert main() == 0, "실패 상세는 위 출력 참고"
'''

TARGETS = {
    "test_module_f_complete.py": ("완성품_전_구간_골든", "F-7 완성품 전 구간 골든", True),
    "test_module_f_design.py": ("design_라우트가_열려_있다", "F-2 design/ HTTP 노출", False),
    "test_module_f_outputs.py": ("산출_3종이_그것뿐이다", "F-4 산출 3종 8조합", False),
    "test_module_f_suggest.py": ("찍기_제안과_제외_사유", "F-5 찍기 후보 제안", True),
}

for fn, (name, title, has_dxf) in TARGETS.items():
    p = os.path.join("tests", fn)
    s = io.open(p, encoding="utf-8").read()
    if f"def test_{name}" in s:
        print(f"  이미 있음 {fn}")
        continue
    tmpl = ENTRY if has_dxf else ENTRY_NODXF
    s = s.rstrip("\n") + "\n" + tmpl.format(name=name, title=title)
    io.open(p, "w", encoding="utf-8").write(s)
    print(f"  진입점 추가 {fn}  → test_{name}")
