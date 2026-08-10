# -*- coding: utf-8 -*-
"""Route 테이블 인벤토리 스모크 — 블루프린트 추출 안전망.

Phase 2 리팩토링(단일 파일 → 도메인 블루프린트 분할)은 route 를 다른 파일로
옮긴다. 옮기는 과정에서 route 가 누락/중복되거나 URL·엔드포인트명·HTTP 메서드가
바뀌면 외부 동작이 깨진다. 이 테스트는 앱의 전체 URL map 을 frozen baseline
(route_inventory.json)과 1바이트 단위로 대조해 그런 회귀를 즉시 잡는다.

실행::

    python -m pytest tests/smoke -q

    # baseline 동결/갱신 (의도된 route 추가·변경 시에만)
    ROUTE_INVENTORY_UPDATE=1 python -m pytest tests/smoke -q
    # (PowerShell)  $env:ROUTE_INVENTORY_UPDATE=1; python -m pytest tests/smoke -q
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
BASELINE = Path(__file__).resolve().parent / "route_inventory.json"
UPDATE = os.environ.get("ROUTE_INVENTORY_UPDATE") == "1"


def _load_app():
    core = str(REPO_ROOT / "core")
    for p in (str(REPO_ROOT), core):
        if p not in sys.path:
            sys.path.insert(0, p)
    # 모듈 C 라우트는 이 플래그로 갈린다. 플래그가 .env(비추적)에 있으면 인벤토리가
    # 기계마다 달라진다 — 여기서 못박아 baseline 이 항상 전량을 덮게 한다.
    # load_dotenv 는 이미 있는 환경변수를 덮지 않으므로 이 설정이 이긴다.
    os.environ["DESIGN_WORKBENCH_ENABLED"] = "1"
    spec = importlib.util.spec_from_file_location(
        "server_app_smoke", str(REPO_ROOT / "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.app


def _inventory(app) -> list:
    rows = []
    for r in app.url_map.iter_rules():
        rows.append({
            "rule": r.rule,
            "endpoint": r.endpoint,
            "methods": sorted(r.methods),
        })
    rows.sort(key=lambda d: (d["rule"], d["endpoint"]))
    return rows


def _dump(obj) -> str:
    return json.dumps(obj, sort_keys=True, ensure_ascii=False, indent=2)


def test_route_inventory():
    actual = _inventory(_load_app())
    if UPDATE:
        BASELINE.write_text(_dump(actual), encoding="utf-8")
        return
    assert BASELINE.exists(), (
        "route_inventory.json 없음 — ROUTE_INVENTORY_UPDATE=1 로 먼저 동결")
    expected = json.loads(BASELINE.read_text(encoding="utf-8"))
    assert _dump(actual) == _dump(expected), (
        "route 테이블이 baseline 과 다름 — 블루프린트 추출 중 route 회귀 의심")
