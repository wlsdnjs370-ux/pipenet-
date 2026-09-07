# -*- coding: utf-8 -*-
"""꺼진 안전망을 **늘 보이게** 한다 — 「안 돌았다」가 「통과」로 보이지 않도록.

■ 왜 필요한가 (2026-09-07 실측)

  시험 여럿이 저장소에 없는 파일(큰 도면·사람의 작업 산물)을 전제로 skip 한다.
  이 개발 기계에는 그 파일이 있어 다 돌지만, 새로 clone 하면 다르다:

      이 기계    1,745 통과 ·   1 skip
      새 clone   1,542 통과 · 115 skip     ← 그런데 **여전히 초록**

  200여 건이 안 도는데 화면이 초록이면 사람은 「괜찮다」로 읽는다. 빨간불은
  누구나 알아채지만 초록불은 아무도 안 본다. 이 저장소가 이미 그 원칙을
  적어 둔 곳이 있다 — `tests/characterization/test_golden.py` 의
  「skip 이 아니라 fail — golden 이 사라지면 안전망도 같이 사라지는데, skip 은
  초록으로 보여서 회귀 검출이 꺼진 걸 아무도 모른다」.

■ 절충 (2026-09-07 결정)

  ⑴ 작은 픽스처는 저장소에 넣었다 — 결과물 6종(374KB)·대명동 도면 3종(17.4MB).
     이것으로 모듈 D 47건과 모듈 F 행위 시험 다수가 어디서든 산다.
  ⑵ 큰 것(B1F 두 장 · 232MB)은 안 넣는다. 대신 그것 때문에 꺼진 시험이 **몇
     건인지** 실행 끝에 한 줄로 보고한다. 초록이되 꺼진 것은 늘 보인다.

  즉 fail 로 바꾸지 않는다(픽스처 없는 환경에서 늘 빨갛기만 하면 그 빨강도
  이내 무시된다). 대신 **숫자를 세어 말한다.**
"""
from __future__ import annotations

import pytest

# «환경 때문에 꺼졌다» 로 셀 사유. 시험이 `pytest.skip(...)` 에 적는 문구와
# 맞춘다 — 사유 문구가 곧 분류다.
_ENV_MARKS = (
    "없음", "없다", "not found", "missing", "제출용", "저장본",
    "표본 도면", "도면 없음", "원본", "픽스처",
)


def _is_env_skip(reason: str) -> bool:
    r = str(reason or "")
    return any(m in r for m in _ENV_MARKS)


def pytest_configure(config):
    config._env_skips = []


def pytest_runtest_logreport(report):
    if report.skipped and report.when == "setup":
        # longrepr = (파일, 줄, "Skipped: 사유")
        reason = ""
        try:
            reason = str(report.longrepr[2])
        except Exception:  # noqa: BLE001 — 형식이 달라도 보고는 계속한다
            reason = str(report.longrepr)
        cfg = getattr(report, "_cfg", None) or _CONFIG.get("cfg")
        if cfg is not None and _is_env_skip(reason):
            cfg._env_skips.append((report.nodeid, reason))


_CONFIG: dict = {}


def pytest_cmdline_main(config):
    _CONFIG["cfg"] = config


def pytest_terminal_summary(terminalreporter, exitstatus, config):
    """실행 끝에 «환경 때문에 꺼진 안전망» 을 한 줄로 말한다."""
    env = getattr(config, "_env_skips", [])
    if not env:
        return
    tr = terminalreporter
    tr.write_sep("=", "환경 때문에 꺼진 검사", yellow=True)
    tr.write_line(
        f"  {len(env)}건이 «저장소에 없는 파일» 때문에 안 돌았습니다 — "
        f"통과가 아니라 **안 돈 것**입니다.")
    seen: dict = {}
    for nodeid, reason in env:
        key = reason.replace("Skipped: ", "").strip()[:60]
        seen.setdefault(key, []).append(nodeid)
    for reason, ids in sorted(seen.items(), key=lambda kv: -len(kv[1])):
        tr.write_line(f"    · {len(ids):>4}건  {reason}")
        tr.write_line(f"           예: {ids[0]}")
    tr.write_line(
        "  큰 도면(B1F 232MB)은 일부러 저장소에 안 넣습니다 — "
        "`.gitignore` 의 「시험 픽스처」 항목 참고.")


@pytest.fixture(scope="session")
def env_fixture_note():
    """시험이 «왜 꺼졌는지» 를 사람 말로 적을 때 쓰는 짧은 안내."""
    return ("저장소에 없는 큰 도면이 필요합니다 — "
            "`.gitignore` 의 「시험 픽스처」 항목 참고.")
