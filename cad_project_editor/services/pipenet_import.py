"""변환이 부르는 기본 라이브러리 채움만. PIPENET .sdf 변환기는 넣지 않는다."""
from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def _ensure_default_libraries(editor) -> None:
    """빈(신규) 에디터에 앱 기본 라이브러리를 채운다.

    .kfp는 라이브러리를 내장하므로, 여기서 채우지 않으면 생성 파일이 빈
    라이브러리를 갖게 되어 열었을 때 배관 표준·부속 등가길이(NFPA 원판)
    조회가 전부 실패한다. 경로 SSOT는 LibraryService._base_path()
    (%LOCALAPPDATA%/K-Fire — 최초 호출 시 공장 기본 JSON 자동 복사).
    """
    import json

    lib = editor._lib
    base = lib._base_path()

    def _read(name: str):
        path = base / name
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception as exc:
            logger.warning("[PIPENET import] 기본 라이브러리 로드 실패 (%s): %s", path, exc)
            return None

    if not editor.pipe_library:
        editor.pipe_library = _read("pipe_library.json") or {}
    if not editor.nozzle_library:
        editor.nozzle_library = _read("nozzle_library.json") or {}
    if not editor.pump_library:
        editor.pump_library = _read("pump_library.json") or {"pumps": []}
    if not lib.fittings_library_v3:
        try:
            lib.load_fittings_v3(base / "fittings_library_v3.json")
        except Exception as exc:
            logger.warning("[PIPENET import] v3 부속 라이브러리 로드 실패: %s", exc)
