# -*- coding: utf-8 -*-
"""Remote 30 GNN(fire-dxf2sdf) 도메인 라우트.

대조 서버.py 에서 register(app, ...) 로 등록. 공유 헬퍼·전역은 main 에 두고 참조 주입 — 라우트 본문 원본 그대로, 엔드포인트명 보존."""
from __future__ import annotations

import json

from flask import jsonify, request, send_file
from werkzeug.utils import secure_filename


def register(app, *, FIRE_DXF2SDF_DIR, FIRE_DXF2SDF_OUTPUT_DIR, UV_EXECUTABLE, _save_upload):

    @app.post("/api/remote30/gnn/run")
    def remote30_gnn_run():
        """DXF 업로드 → fire-dxf2sdf Phase 1-3 실행 → JSON 반환.

        Form fields:
            dxf_file: required (.dxf)
            k: int, default 30
            alarm_x, alarm_y: float (둘 다 또는 둘 다 없음 — 없으면 auto)
            selection_method: simple_greedy | branch_aware | branch_cluster
        """
        import secrets
        import subprocess

        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        if dxf_path is None:
            return jsonify({"ok": False, "message": "DXF 파일이 필요합니다."}), 400

        # 파라미터
        try:
            k = int(request.form.get("k", "30"))
            if k <= 0:
                raise ValueError("k must be positive")
        except (TypeError, ValueError):
            return jsonify({"ok": False, "message": "k 값이 잘못됨"}), 400

        method = request.form.get("selection_method", "simple_greedy")
        if method not in {"simple_greedy", "branch_aware", "branch_cluster"}:
            return jsonify({"ok": False, "message": f"알 수 없는 선정 방식: {method}"}), 400

        alarm_x = request.form.get("alarm_x")
        alarm_y = request.form.get("alarm_y")
        if (alarm_x is None) != (alarm_y is None):
            return jsonify(
                {"ok": False, "message": "alarm_x 와 alarm_y 는 함께 지정해야 합니다."}
            ), 400

        # 실행 디렉토리
        run_id = secrets.token_hex(6)
        out_dir = FIRE_DXF2SDF_OUTPUT_DIR / run_id
        out_dir.mkdir(parents=True, exist_ok=True)

        # uv subprocess
        if not UV_EXECUTABLE.is_file():
            return jsonify(
                {
                    "ok": False,
                    "message": f"uv 실행 파일 없음: {UV_EXECUTABLE}",
                }
            ), 500
        if not FIRE_DXF2SDF_DIR.is_dir():
            return jsonify(
                {
                    "ok": False,
                    "message": f"fire-dxf2sdf 디렉토리 없음: {FIRE_DXF2SDF_DIR}",
                }
            ), 500

        json_out = out_dir / "full.json"
        cmd = [
            str(UV_EXECUTABLE),
            "--project",
            str(FIRE_DXF2SDF_DIR),
            "run",
            "python",
            "-m",
            "fire_dxf2sdf.pipeline",
            "--stage",
            "3",
            "--input",
            str(dxf_path),
            "--k",
            str(k),
            "--selection-method",
            method,
            "--metrics-out",
            str(out_dir),
            "--json-out",
            str(json_out),
            "--log-level",
            "WARNING",
        ]
        if alarm_x is not None and alarm_y is not None:
            cmd.extend(["--alarm-x", str(alarm_x), "--alarm-y", str(alarm_y)])

        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=180,
                encoding="utf-8",
                errors="replace",
            )
        except subprocess.TimeoutExpired:
            return jsonify(
                {
                    "ok": False,
                    "message": "처리 시간 초과 (180초). 더 작은 도면으로 시도하세요.",
                }
            ), 504

        if proc.returncode != 0:
            # exit code 2 = DxfParseError, 3 = FireDxf2SdfError, 1 = SystemExit
            stderr_tail = (proc.stderr or "")[-1500:]
            return jsonify(
                {
                    "ok": False,
                    "exit_code": proc.returncode,
                    "message": "fire-dxf2sdf 실행 실패",
                    "stderr": stderr_tail,
                }
            ), 422 if proc.returncode in (2, 3) else 500

        # JSON 결과 로딩
        if not json_out.is_file():
            return jsonify(
                {
                    "ok": False,
                    "message": "JSON 결과 파일이 생성되지 않음",
                    "stdout": (proc.stdout or "")[-500:],
                }
            ), 500
        try:
            summary = json.loads(json_out.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            return jsonify(
                {"ok": False, "message": f"JSON 파싱 실패: {exc}"}
            ), 500

        # PNG 경로 확인
        png_candidates = list(out_dir.glob("*_stage3.png"))
        png_url: str | None = None
        if png_candidates:
            png_name = png_candidates[0].name
            png_url = f"/api/remote30/gnn/result/{run_id}/{png_name}"

        return jsonify(
            {
                "ok": True,
                "run_id": run_id,
                "summary": summary,
                "png_url": png_url,
                "json_url": f"/api/remote30/gnn/result/{run_id}/full.json",
            }
        )

    @app.get("/api/remote30/gnn/result/<run_id>/<path:filename>")
    def remote30_gnn_result(run_id: str, filename: str):
        """GNN 파이프라인 결과 파일 (PNG / JSON) 반환."""
        safe_run_id = secure_filename(run_id)
        if not safe_run_id or safe_run_id != run_id:
            return "잘못된 run_id 입니다.", 400
        target = FIRE_DXF2SDF_OUTPUT_DIR / safe_run_id / filename
        # 경로 traversal 방지
        try:
            target.resolve().relative_to(FIRE_DXF2SDF_OUTPUT_DIR.resolve())
        except ValueError:
            return "잘못된 경로", 400
        if not target.is_file():
            return "결과 파일 없음", 404
        if filename.endswith(".png"):
            return send_file(target, mimetype="image/png")
        if filename.endswith(".json"):
            return send_file(target, mimetype="application/json")
        return send_file(target, as_attachment=True)
