# -*- coding: utf-8 -*-
"""[리팩터링 절단기] api_design.py 의 선정·인계·신선도 덩이 → selection.py.

집 규약(2026-08-25 패키지 분리와 같음): 본문은 **잘라 옮긴다** — 재타이핑 0.
블록 경계는 def 줄로 못박고, 앵커가 정확히 한 번씩만 나오는지 assert 한다.

    python data/_refactor_selection_split.py
"""
from __future__ import annotations

import io
import re

SRC = "routes/module_f/api_design.py"
DST = "routes/module_f/selection.py"

HEADER = '''# -*- coding: utf-8 -*-
"""선정 인계 · 표 신선도 — 라우트가 아닌 «판단» 만 사는 곳.

2026-09-11 리팩터링으로 api_design.py 에서 잘라 옮겼다(재타이핑 0 —
`data/_refactor_selection_split.py` 가 절단기). 두 가지 이유다:

  · api_design.py 가 1,597줄까지 자랐다 — 이 주에 자란 것이 전부 이 덩이다
    (두 화면 선정일치 §2-1~§2-5 · 영역 가두기 · 표 신선도).
  · `remote30._design_corridor` 가 이 판단들을 쓰려고 **라우트 파일을**
    import 하고 있었다. 순수 판단이 라우트 등록 파일에 살면 그런 의존이 는다.

여기 있는 것은 전부 순수 함수다 — Flask 도, 엔진 부팅(_boot)도 모듈 레벨에서
끌지 않는다(엔진 참조는 함수 안 지연 import 뿐). 그래서 시험이 서버 없이 바로
import 해 쓴다.

★이름 계약: api_design.py 가 이 이름들을 **다시 내보낸다** — 시험·탐침이
  `routes.module_f.api_design` 에서 import 하기 때문이다. 가른 것 때문에
  도구가 깨지면 안 된다(패키지 분리 때의 `__init__.py` 재수출과 같은 규약).
"""
from __future__ import annotations
'''

REIMPORT = '''
# [리팩터링 2026-09-11] 선정 인계·표 신선도는 selection.py 로 갈라냈다.
#   여기서 다시 내보내는 이유는 시험·탐침이 이 모듈에서 그 이름들을 import
#   하기 때문이다 — 가른 것 때문에 도구가 깨지면 안 된다(§패키지 분리 규약).
from routes.module_f.selection import (  # noqa: F401 — 재수출 포함
    _EDIT_ONLY_WORST_KEYS, _adopt_final_worst, _classify_excluded,
    _design_stale, _handoff_after_table, _head_row, _in_rect, _load_map,
    _selection_sig, _worst_handoff_note, zone_confined_pool)
'''


def cut(text: str, start_marker: str, end_marker: str) -> tuple[str, str]:
    """start_marker 줄부터 end_marker 줄 «직전» 까지를 (남은 글, 잘린 글) 로."""
    for m in (start_marker, end_marker):
        n = text.count(m)
        assert n == 1, f"앵커가 {n}번 나온다(1이어야): {m!r}"
    i = text.index(start_marker)
    j = text.index(end_marker)
    assert i < j, "앵커 순서가 뒤집혔다"
    return text[:i] + text[j:], text[i:j]


def main() -> None:
    s = io.open(SRC, encoding="utf-8").read()
    n0 = s.count("\n")

    s, block_a = cut(s, "def _load_map(got: dict) -> dict:",
                     "def _view_opts(cfg: dict) -> dict:")
    s, block_b = cut(s, "def _classify_excluded(sess: dict, got: dict, board, probe=None) -> dict:",
                     "def emit_design_files(sess: dict, UPLOAD_DIR, cfg: dict | None = None):")

    # 재수출 import — jobs import 바로 다음 줄에.
    anchor = "from routes.module_f.jobs import _job_running, _run_job, route_session\n"
    assert s.count(anchor) == 1, "import 앵커가 안 보인다"
    s = s.replace(anchor, anchor + REIMPORT)

    # 잘린 자리의 3연속 이상 빈 줄을 2줄로.
    s = re.sub(r"\n{4,}", "\n\n\n", s)

    io.open(SRC, "w", encoding="utf-8", newline="\n").write(s)
    out = HEADER + "\n\n" + block_a.rstrip() + "\n\n\n" + block_b.rstrip() + "\n"
    io.open(DST, "w", encoding="utf-8", newline="\n").write(out)

    import ast
    ast.parse(io.open(SRC, encoding="utf-8").read())
    ast.parse(io.open(DST, encoding="utf-8").read())
    n1 = io.open(SRC, encoding="utf-8").read().count("\n")
    n2 = io.open(DST, encoding="utf-8").read().count("\n")
    print(f"api_design.py {n0} → {n1} 줄 · selection.py {n2} 줄 · 구문 OK")


if __name__ == "__main__":
    main()
