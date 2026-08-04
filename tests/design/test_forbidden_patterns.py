# -*- coding: utf-8 -*-
"""지시서 §13.1 — 금칙어 회귀.

여기 걸리는 패턴들은 전부 "한 번 저지르면 조용히 통과하는" 실수다. 방사시간
20분 고정처럼 30층 건물에서만 틀리는 값은 리뷰로 잡히지 않는다. 그래서 문법이
아니라 **문자열**로 막는다.

패턴에 걸렸을 때 표현만 바꿔 회피하지 마라. 걸린 줄이 진짜로 층수 분기를
빠뜨렸는지부터 확인하고, 정말 오탐이면 이 파일의 `ALLOWLIST` 에 사유와 함께
등록하라.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

FORBIDDEN = [
    (r"×\s*20\s*분",                         "G1 방사시간 20분 고정"),
    (r"discharge_minutes\s*=\s*20\b",        "G1 하드코딩"),
    (r"water_supply.*\*\s*1\.6(?!\s*if)",    "G1 층수 분기 없는 1.6 고정"),
    (r"activate_103[bB]\s*=\s*True",         "G2 103B 자동 활성"),
    (r'rack_storage["\']?\s*:\s*.*103B',     "G2 랙크식 자동 분기"),
    (r"branch_heads_max",                    "G3 per_side 누락 변수명"),
    (r"beams.*0\.6|0\.6.*beams",             "G4 보를 60cm 룰에 포함"),
    (r"별표\s*\+\s*1|min_dn\(.*\)\s*\+\s*1", "G5 관경 자동 상향"),
    (r"zone_grid_relief_m2\s*[:=]\s*3700",   "D1 결정 전 구현"),
    (r"tau_water.*(FAIL|재설계|redesign)",    "D2 결정 전 구현"),
]

SCAN_PATHS = ["core/design/", "routes/r30_design.py", "templates/design_workbench.html"]

# (경로 접미사, 줄번호 무관 원문 조각, 사유) — 오탐만 등록한다.
ALLOWLIST: list[tuple[str, str, str]] = []

_SCAN_SUFFIXES = {".py", ".html", ".js", ".css", ".md", ".json"}


def _scan_files() -> list[Path]:
    files: list[Path] = []
    for entry in SCAN_PATHS:
        target = _ROOT / entry
        if not target.exists():
            continue  # 아직 만들지 않은 PR 산출물
        if target.is_dir():
            files.extend(
                p for p in target.rglob("*")
                if p.is_file() and p.suffix in _SCAN_SUFFIXES
                and "__pycache__" not in p.parts
            )
        else:
            files.append(target)
    return files


def _allowed(path: Path, line: str) -> bool:
    rel = path.relative_to(_ROOT).as_posix()
    return any(rel.endswith(suffix) and fragment in line
               for suffix, fragment, _reason in ALLOWLIST)


@pytest.mark.parametrize("pattern,label", FORBIDDEN, ids=[l for _p, l in FORBIDDEN])
def test_no_forbidden_pattern(pattern, label):
    rx = re.compile(pattern)
    hits = []
    for path in _scan_files():
        for lineno, line in enumerate(
                path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
            if rx.search(line) and not _allowed(path, line):
                hits.append(f"{path.relative_to(_ROOT).as_posix()}:{lineno}: {line.strip()}")
    assert not hits, f"[{label}] 금칙어 {len(hits)}건\n" + "\n".join(hits)


def test_scan_target_exists():
    """SCAN_PATHS 가 전부 사라지면 이 테스트는 아무것도 검사하지 않으면서 통과한다."""
    assert _scan_files(), "검사 대상이 하나도 없다 — SCAN_PATHS 를 확인하라"
