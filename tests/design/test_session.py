# -*- coding: utf-8 -*-
"""지시서 §C.2 — 세션 저장·낙관적 잠금·감사 로그.

여기가 깨지면 "어떤 제약으로 설계했는지" 를 나중에 재구성할 수 없다. 설계 도서는
날인 문서라 그 재구성 불가능성이 곧 감리 대응 불가다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design import session as S  # noqa: E402


def test_세션_id_는_추측할_수_없는_uuid4_다(tmp_path):
    sess = S.DesignSession.create(tmp_path)
    assert S.is_session_id(sess.sid)
    assert (tmp_path / sess.sid / "meta.json").is_file()


@pytest.mark.parametrize("bad", ["..", "../etc", "a/b", "", "1234", "%2e%2e"])
def test_세션_id_가_아니면_열리지_않는다(tmp_path, bad):
    """세션 id 가 그대로 디렉터리 이름이 되므로 여기가 유일한 경로 조작 방어선이다."""
    with pytest.raises(S.SessionNotFound):
        S.DesignSession.open(tmp_path, bad)


def test_쓰기는_버전을_올린다(tmp_path):
    sess = S.DesignSession.create(tmp_path)
    assert sess.write("building.json", {"rooms": []}) == 1
    assert sess.write("building.json", {"rooms": [1]}) == 2
    data, version = sess.read("building.json")
    assert version == 2 and data["rooms"] == [1]


def test_버전이_어긋나면_현재_내용과_함께_거절한다(tmp_path):
    """두 탭이 같은 세션을 열 수 있다. 나중 쓰기가 먼저 쓴 것을 조용히 지우면 안 된다."""
    sess = S.DesignSession.create(tmp_path)
    sess.write("building.json", {"rooms": ["a"]})
    sess.write("building.json", {"rooms": ["a", "b"]})
    with pytest.raises(S.VersionConflict) as exc:
        sess.write("building.json", {"rooms": ["a", "c"]}, if_version=1)
    assert exc.value.current == 2
    assert exc.value.data["rooms"] == ["a", "b"]


def test_constraints_는_한_번_쓰면_불변이다(tmp_path):
    sess = S.DesignSession.create(tmp_path)
    sess.write("constraints.json", {"scenario_head_count": 20})
    with pytest.raises(S.ImmutableArtifact):
        sess.write("constraints.json", {"scenario_head_count": 30})


def test_없는_산출물은_조용히_비지_않고_알린다(tmp_path):
    sess = S.DesignSession.create(tmp_path)
    with pytest.raises(S.ArtifactNotFound):
        sess.read("building.json")


def test_감사_로그는_덧붙이기만_한다(tmp_path):
    sess = S.DesignSession.create(tmp_path)
    sess.audit("jinwon", "GATE", "use_override", {"room": "R-1F-012"})
    sess.audit("system", "C2", "constraints_built", {"n": 20})
    entries = sess.audit_entries()
    assert [e["event"] for e in entries] == ["created", "use_override", "constraints_built"]
    assert entries[1]["detail"]["room"] == "R-1F-012"


def test_같은_단계를_동시에_두_번_돌리지_않는다(tmp_path):
    sess = S.DesignSession.create(tmp_path)
    with sess.stage_lock("c4"):
        with pytest.raises(S.StageBusy):
            with sess.stage_lock("c4"):
                pass
    with sess.stage_lock("c4"):
        pass


def test_상태에_산출물_버전과_만료가_담긴다(tmp_path):
    sess = S.DesignSession.create(tmp_path, operator="jinwon")
    sess.write("building.json", {"rooms": []})
    status = sess.status()
    assert status["meta"]["operator"] == "jinwon"
    assert status["meta"]["expires_at"] > status["meta"]["created_at"]
    assert status["artifacts"]["building.json"]["version"] == 1
    assert status["artifacts"]["design.json"] is None
