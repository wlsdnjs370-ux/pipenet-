# -*- coding: utf-8 -*-
"""정적 파일을 고쳤으면 `?v=` 도 같이 올렸는가 — «반쯤 새 화면» 막기.

■ 2026-09-11 실측으로 밟은 함정

  같은 프로세스·같은 파이썬인데 자원마다 반영 방식이 다르다:

      .py    재기동해야 반영         (import 캐시)
      .html  auto_reload 로 즉시 반영 (TEMPLATES_AUTO_RELOAD=True)
      .js    브라우저 캐시에 걸림     (`?v=` 가 그대로면 옛 것을 쓴다)

  그래서 JS 를 고치고 `?v=` 를 안 올리면 **템플릿은 오늘 것, JS 는 어제 것**
  이 된다. 실제로 그날 「고른 것 중 빠짐」 체크박스가 화면에 떠 있는데
  눌러도 아무 일이 없는 상태를 만들었다 — 「고쳤다는데 안 된다」의 정체다.

■ 이 시험이 잡는 것

  템플릿의 `?v=YYYYMMDD-...` 날짜가, 그 템플릿이 물고 있는 정적 파일을
  **마지막으로 고친 커밋 날짜보다 오래됐으면** 실패한다.

  ★git 을 쓸 수 없는 환경(얕은 클론·tarball)에서는 조용히 건너뛴다.
    시험이 환경 때문에 빨간불이 되면 사람이 시험을 끄게 된다.
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
_REF = re.compile(r'/static/(?P<file>[A-Za-z0-9_.\-]+)\?v=(?P<ver>[0-9]{8})')


def _last_commit_date(rel: str) -> str | None:
    """그 파일을 마지막으로 고친 커밋 날짜(YYYYMMDD). 못 읽으면 None."""
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%cd", "--date=format:%Y%m%d",
             "--", rel],
            cwd=str(_ROOT), capture_output=True, text=True, timeout=30)
    except (OSError, subprocess.SubprocessError):
        return None
    if out.returncode != 0:
        return None
    v = (out.stdout or "").strip()
    return v if re.fullmatch(r"[0-9]{8}", v) else None


def _pairs():
    """(템플릿, 정적파일, 버전) — 버전이 박힌 참조만 모은다."""
    for tpl in sorted((_ROOT / "templates").glob("*.html")):
        try:
            text = tpl.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue
        for m in _REF.finditer(text):
            yield tpl, m.group("file"), m.group("ver")


def test_참조된_정적파일이_실제로_있다():
    missing = [(t.name, f) for t, f, _v in _pairs()
               if not (_ROOT / "static" / f).is_file()]
    assert not missing, f"템플릿이 없는 정적 파일을 가리킨다: {missing}"


def test_정적파일을_고쳤으면_캐시버스터도_올렸다():
    """★이 시험이 없어서 오늘 «템플릿은 새 것 · JS 는 캐시된 옛 것» 을 냈다."""
    if _last_commit_date("static") is None:
        pytest.skip("git 이력을 읽을 수 없는 환경 — 건너뛴다")
    stale = []
    for tpl, fname, ver in _pairs():
        touched = _last_commit_date(f"static/{fname}")
        if touched is None:
            continue                      # 아직 커밋 안 된 새 파일
        if ver < touched:
            stale.append(
                f"{tpl.name}: {fname} 은 {touched} 에 고쳤는데 ?v={ver} 그대로")
    assert not stale, (
        "정적 파일을 고치고 캐시버스터를 안 올렸다 — 브라우저가 옛 것을 쓴다:\n  "
        + "\n  ".join(stale))
