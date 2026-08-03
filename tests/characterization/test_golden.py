# -*- coding: utf-8 -*-
"""Remote30 특성화 golden 대조 테스트.

동작 보존(behavior-preserving) 리팩토링 안전망. 각 케이스의 현재 출력을
golden/<name>.json 과 1바이트 단위로 대조한다.

실행::

    # 일반 대조 (리팩토링 후 회귀 검출)
    python -m pytest tests/characterization -q

    # baseline 동결/갱신 (리팩토링 전 1회, 또는 의도된 동작 변경 시)
    GOLDEN_UPDATE=1 python -m pytest tests/characterization -q
    # (PowerShell)  $env:GOLDEN_UPDATE=1; python -m pytest tests/characterization -q

갱신 모드는 각 케이스를 2회 실행해 결정적(stable)임을 확인한 뒤에만 저장한다.
비결정적 케이스는 저장을 거부하고 실패시켜, 흔들리는 스냅샷이 섞이지 않게 한다.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

import golden_cases as gc

GOLDEN_DIR = Path(__file__).resolve().parent / "golden"
UPDATE = os.environ.get("GOLDEN_UPDATE") == "1"


def _golden_file(name: str) -> Path:
    return GOLDEN_DIR / f"{name}.json"


def _dump(obj) -> str:
    return json.dumps(obj, sort_keys=True, ensure_ascii=False, indent=2)


@pytest.mark.parametrize("name", sorted(gc.CASES.keys()))
def test_golden(name):
    run = gc.CASES[name]
    actual = run()

    if UPDATE:
        # 결정성 검증 — 2회 실행 동일해야 baseline 으로 동결
        second = run()
        assert _dump(actual) == _dump(second), (
            f"케이스 '{name}' 가 비결정적이라 golden 으로 동결 불가 "
            f"(두 번 실행 결과 불일치)"
        )
        GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
        _golden_file(name).write_text(_dump(actual), encoding="utf-8")
        return

    gf = _golden_file(name)
    # skip 이 아니라 fail — golden 파일이 사라지면 안전망도 같이 사라지는데,
    # skip 은 초록으로 보여서 회귀 검출이 꺼진 걸 아무도 모른다.
    assert gf.exists(), f"golden 없음 — GOLDEN_UPDATE=1 로 먼저 생성: {name}"
    expected = json.loads(gf.read_text(encoding="utf-8"))
    assert _dump(actual) == _dump(expected), (
        f"케이스 '{name}' 출력이 golden 과 다름 — 동작이 변경됨(회귀 의심)"
    )
