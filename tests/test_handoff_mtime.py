# -*- coding: utf-8 -*-
"""handoff 캐시 판정 — mtime 은 빠른 길이지 거부권이 아니다.

실측 회귀(B1F): 원본 DXF 의 내용이 한 바이트도 안 바뀌었는데 mtime 만 174초
어긋나 캐시가 기각됐고, 그 안에 있던 치수 텍스트 3,168행(치수로 읽히는 것
533개)이 함께 사라져 관경이 **100% 별표1 폴백**이 됐다. 화면에는 경고 한 줄
없이 관경표가 규약값으로만 채워졌다.

여기서 못박는 계약:
  ① mtime 이 같으면 sha256 을 계산하지 않는다 (115MB 해싱 회피)
  ② mtime 이 달라도 **내용이 같으면** 캐시는 유효하다
  ③ 내용이 다르면 mtime 이 같아도(크기까지 같아도) 기각한다
"""
from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "cad_project_editor_g"))

from services.cad_import.pipeline import handoff  # noqa: E402


@pytest.fixture()
def dxf(tmp_path):
    p = tmp_path / "원본.dxf"
    p.write_bytes(b"0\nSECTION\n2\nENTITIES\n0\nENDSEC\n0\nEOF\n")
    return p


def _meta_for(path) -> dict:
    m = dict(handoff._source_meta(str(path)))
    m["format"] = handoff.FORMAT
    m["prep_sha256"] = handoff._compatible_prep_digest()
    return m


def test_그대로면_통과(dxf):
    assert handoff._meta_matches(_meta_for(dxf), str(dxf)) is True


def test_mtime만_달라도_내용이_같으면_통과(dxf):
    """같은 파일을 다시 올리면 mtime 만 새로 찍힌다 — 캐시는 여전히 유효하다."""
    meta = _meta_for(dxf)
    meta["source_mtime_ns"] = str(int(meta["source_mtime_ns"]) - 174_000_000_000)
    assert handoff._meta_matches(meta, str(dxf)) is True


def test_mtime이_같으면_sha를_계산하지_않는다(dxf, monkeypatch):
    """빠른 길 — 115MB 해싱을 건너뛴다. 부르면 실패로 본다."""
    meta = _meta_for(dxf)

    def _boom(_p):
        raise AssertionError("mtime 이 같은데 sha256 을 계산했다")

    monkeypatch.setattr(handoff, "_sha256_file", _boom)
    assert handoff._meta_matches(meta, str(dxf)) is True


def test_내용이_바뀌면_기각(dxf):
    meta = _meta_for(dxf)
    meta["source_sha256"] = "0" * 64
    meta["source_mtime_ns"] = str(int(meta["source_mtime_ns"]) - 1)
    assert handoff._meta_matches(meta, str(dxf)) is False


def test_크기가_다르면_기각(dxf):
    meta = _meta_for(dxf)
    meta["source_size"] = str(int(meta["source_size"]) + 1)
    assert handoff._meta_matches(meta, str(dxf)) is False


def test_준비지문이_다르면_기각(dxf):
    """엔진의 stage1 준비 코드가 바뀌면 캐시는 못 쓴다 — 이건 유지한다."""
    meta = _meta_for(dxf)
    meta["prep_sha256"] = "다른 지문"
    assert handoff._meta_matches(meta, str(dxf)) is False


def test_포맷이_다르면_기각(dxf):
    meta = _meta_for(dxf)
    meta["format"] = "stage1-world-sqlite-v1"
    assert handoff._meta_matches(meta, str(dxf)) is False


def test_다른_경로면_기각(dxf, tmp_path):
    other = tmp_path / "다른곳.dxf"
    other.write_bytes(dxf.read_bytes())
    meta = _meta_for(dxf)
    assert handoff._meta_matches(meta, str(other)) is False
