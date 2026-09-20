"""Module F file/HTTP boundary for audited KFP exports."""
from __future__ import annotations

import hashlib
import io
import json
import os
import tempfile
from pathlib import Path

from flask import jsonify, request, send_file

from core.kfp_editability import prepare_kfp


def finish_kfp(path: str | Path, payload: dict | None = None) -> tuple[dict, dict]:
    """Set a precision grid on a newly exported file, preserving calculation data."""
    path = Path(path)
    if payload is None:
        payload = json.loads(path.read_text(encoding="utf-8"))
    result, report = prepare_kfp(payload)
    if result is None:
        raise ValueError("KFP 출력 검사: " + "; ".join(report.errors[:8]))
    temp = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent,
                                         prefix=path.name + ".", suffix=".tmp", delete=False) as f:
            temp = Path(f.name)
            json.dump(result, f, ensure_ascii=False, indent=2, allow_nan=False)
        os.replace(temp, path)
    finally:
        if temp is not None and temp.exists():
            temp.unlink()
    return result, report.to_dict()


def send_kfp(path: str | Path, filename: str):
    """Download a session-owned KFP, or inspect a distinct legacy editing copy.

    Editing downloads require the hash returned by inspection so a recalculation
    or another tab cannot silently replace the file the user just reviewed.
    Existing sessions are inspected too; no reconversion is required.
    """
    mode = request.args.get("variant", "precise")
    if mode not in ("precise", "solver-edit"):
        return jsonify(ok=False, message="지원하지 않는 KFP 출력 모드입니다."), 400
    try:
        raw = Path(path).read_bytes()
        digest = hashlib.sha256(raw).hexdigest()
        payload, report = prepare_kfp(json.loads(raw), mode=mode)
    except (OSError, ValueError, TypeError) as exc:
        return jsonify(ok=False, message=f"KFP 파일을 읽을 수 없습니다: {exc}"), 422
    if request.args.get("inspect") == "1":
        return jsonify(ok=True, source_sha256=digest, report=report.to_dict())
    if payload is None:
        return jsonify(ok=False, message="; ".join(report.errors[:8]), report=report.to_dict()), 422
    if mode == "solver-edit":
        if request.args.get("source_sha256") != digest:
            return jsonify(ok=False, message="편집용 파일의 변경량을 다시 확인하세요. 원본이 바뀌었거나 검사 결과가 없습니다."), 409
        filename = Path(filename).stem + "_편집용_1cm.kfp"
    content = json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False).encode("utf-8")
    return send_file(io.BytesIO(content), as_attachment=True, download_name=filename,
                     mimetype="application/json", max_age=0)
