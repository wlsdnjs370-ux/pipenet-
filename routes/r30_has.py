# -*- coding: utf-8 -*-
"""Remote 30 HAS(하스) 파서 도메인 라우트.

대조 서버.py 에서 register(app, ...) 로 등록. 공유 헬퍼·전역은 main 에 두고 참조 주입 — 라우트 본문 원본 그대로, 엔드포인트명 보존."""
from __future__ import annotations

from flask import jsonify


def register(app, *, _common_network_to_geometry, _err500, _save_upload):

    @app.post("/api/remote30/has/parse")
    def remote30_has_parse():
        """Remote 30 프로토타입 — .has(HASS) 파일 불러오기 → 통합 모드 geometry.

        parse_has 로 CommonNetwork 를 만든 뒤, 통합(combined) 렌더러가 쓰는 geometry
        스키마(nodes/pipes/*_labels)로 변환해 반환한다. 라이저·기계실 구분 정보는 .has 에
        없으므로(우리 export 가 비움) 비워두고, 수원(IoNode=1)·펌프만 강조한다.
        """
        try:
            has_path = _save_upload("has_file", {".has"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        from has_converter import parse_has
        try:
            net = parse_has(has_path)
        except Exception as exc:  # noqa: BLE001
            return _err500(exc, message=f"HAS 파싱 실패: {str(exc)[:280]}")

        geometry = _common_network_to_geometry(net)
        src = next((n for n in geometry["nodes"] if n["io"] == "Input"), None)
        return jsonify({
            "ok": True,
            "filename": has_path.name,
            "nodes": len(geometry["nodes"]),
            "pipes": len(geometry["pipes"]),
            "source_label": src["label"] if src else None,
            "geometry": geometry,
        })
