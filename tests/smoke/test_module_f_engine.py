# -*- coding: utf-8 -*-
"""[F-0] 모듈 F 가 무는 엔진이 G 인지 못박는 스모크.

두 편집기 트리(cad_project_editor · cad_project_editor_g)는 services/domain
패키지 이름이 같다. 어느 쪽이 import 되는지가 sys.path 순서 우연에 걸리면,
F 는 겉으로 멀쩡히 돌면서 **옛 엔진으로 계산한다** — 그 오염은 화면에 안
보인다. 그래서 «services 가 물리적으로 어느 트리 파일인가» 를 직접 본다.

실행::

    python -m pytest tests/smoke -q
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def test_module_f_engine_is_g():
    from routes.module_f.common import (
        EDITOR_ROOT, IMPORT_WORK_ROOT, RETIRED_E_ROOT, _boot)

    # 선언부터 — 재지정이 되돌아가면 여기서 잡힌다.
    assert EDITOR_ROOT.name == "cad_project_editor_g"
    assert RETIRED_E_ROOT.name == "cad_project_editor"

    _boot()

    # ① services 실체가 G 트리다.
    import services
    svc = str(Path(services.__file__).resolve())
    assert svc.startswith(str(EDITOR_ROOT.resolve())), svc

    # ② 은퇴한 E 루트는 path 에 없다.
    assert str(RETIRED_E_ROOT) not in sys.path

    # ③ 작업폴더도 G 다 — 데스크톱 E 와 웹 F 가 같은 캐시를 헤집지 않는다.
    assert str(IMPORT_WORK_ROOT).startswith(str(EDITOR_ROOT))
    from services.cad_import.pipeline import handoff
    assert str(Path(handoff.pick_out_dir()).resolve()).startswith(
        str(EDITOR_ROOT.resolve()))

    # ④ F 의 심장 — G 의 design/ 이 이 엔진에서 실제로 열린다.
    from services.cad_import.design import restrict, worst  # noqa: F401


def test_migrated_keys_visible():
    """E 작업폴더에만 있던 키가 재지정 후에도 보인다(마이그레이션 증명)."""
    from routes.module_f.common import IMPORT_WORK_ROOT
    spec_dir = IMPORT_WORK_ROOT / "0단계_새찍기"
    for key in ("3F", "apt", "MF-301(지하3층 소방시설(기계) 평면도)"):
        assert (spec_dir / f"{key}_찍은스펙.json").is_file(), key
    # 설명 그림도 함께 왔다 — 없으면 /api/module-f/diagram/* 가 404 다.
    for png in ("_kfp_sample_가지.png", "_kfp_sample_상향식.png",
                "_kfp_sample_하향식.png", "_kfp_sample_상하향식.png",
                "_kfp_sample_알람밸브.png"):
        assert (IMPORT_WORK_ROOT / png).is_file(), png
