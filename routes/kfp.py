# -*- coding: utf-8 -*-
"""11번 모듈 — KFP ↔ SDF 변환기 라우트.

`대조 서버.py` 에서 `register(app)` 로 등록. 공유 헬퍼 의존 없음
(kfp_sdf_converter / remote30_prototype 는 함수 내부에서 지역 import).
엔드포인트명(kfp_sdf_converter_page/kfp_sdf_preview/kfp_sdf_convert) 보존.
"""
from __future__ import annotations

from flask import (Response, jsonify, make_response, render_template, request)


def register(app) -> None:
    @app.get("/kfp-sdf-converter")
    def kfp_sdf_converter_page():
        """11번 모듈 — K-solver .kfp ↔ PIPENET .sdf 양방향 변환기."""
        response = make_response(render_template("kfp_sdf_converter.html"))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return response

    @app.post("/api/kfp_sdf/preview")
    def kfp_sdf_preview():
        """업로드된 .kfp 또는 .sdf → 미리보기 + (KFP 의 경우) round-trip 통계.

        Form: file (multipart). 응답: {ok, preview, round_trip?}.
        임시 파일에 저장 후 kfp_sdf_converter.parse_for_preview 호출.
        """
        import tempfile, os
        from pathlib import Path as _Path
        import kfp_sdf_converter as _conv

        f = request.files.get("file")
        if f is None:
            return jsonify({"ok": False, "message": "file 누락"}), 400
        suffix = _Path(f.filename or "").suffix.lower() or ".kfp"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = tmp.name
        try:
            preview = _conv.parse_for_preview(tmp_path)
            out = {"ok": True, "preview": preview}
            if preview["format"] == "kfp":
                try:
                    out["round_trip"] = _conv.round_trip_check(tmp_path)
                except Exception as _e:
                    out["round_trip"] = {"error": str(_e)}
            return jsonify(out)
        except Exception as exc:
            return jsonify({"ok": False, "message": f"파싱 실패: {exc}"}), 400
        finally:
            try: os.unlink(tmp_path)
            except OSError: pass

    @app.post("/api/kfp_sdf/convert")
    def kfp_sdf_convert():
        """양방향 변환 + 결과 파일 즉시 다운로드.

        Form: file (multipart), direction ("to_sdf" | "to_kfp").
        응답: 파일 내용 (application/octet-stream).
        """
        import tempfile, os, json as _json
        from pathlib import Path as _Path
        import kfp_sdf_converter as _conv

        f = request.files.get("file")
        direction = (request.form.get("direction") or "").strip()
        if f is None:
            return jsonify({"ok": False, "message": "file 누락"}), 400
        if direction not in ("to_sdf", "to_kfp"):
            return jsonify({"ok": False, "message": "direction 은 to_sdf 또는 to_kfp"}), 400

        suffix = _Path(f.filename or "").suffix.lower() or ".bin"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = tmp.name
        try:
            if direction == "to_sdf":
                # PIPENET native 모드 + SLF 동봉 ZIP — PIPENET 계산/아이소 모두 통과.
                # 단순 .sdf 만 보내면 SLF 없어서 내경(Internal) "Unset" 이슈 발생.
                import zipfile, io
                from remote30_prototype import resolve_standard_slf
                stem = _Path(f.filename or "converted").stem or "converted"
                # SLF 파일명을 SDF 의 <User-lib file=...> 에 매핑.
                # 같은 ZIP 안에 stem.slf 동봉되므로 사용자가 풀면 PIPENET 자동 인식.
                xml = _conv.convert_kfp_to_sdf(tmp_path, use_pipenet_native=True,
                                                 project_title=stem,
                                                 slf_filename=f"{stem}.slf")
                # ZIP 으로 .sdf + .slf 묶기 — PIPENET 권장 (zip 다운 후 압축해제 사용)
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr(f"{stem}.sdf", xml)
                    try:
                        slf_src = resolve_standard_slf()
                        zf.write(slf_src, f"{stem}.slf")
                    except Exception as _slf_err:
                        # SLF 없어도 .sdf 에 6 schedule 임베드 되어 있으므로 동작은 함
                        zf.writestr("_slf_missing.txt",
                                    f"SLF 동봉 실패: {_slf_err}\n"
                                    "→ .sdf 안에 6 schedule 임베드 되어 있어 PIPENET 동작은 정상.")
                return Response(buf.getvalue(), mimetype="application/zip",
                                headers={"Content-Disposition": f'attachment; filename="{stem}.zip"'})
            else:
                kfp = _conv.convert_sdf_to_kfp(tmp_path)
                data = _json.dumps(kfp, ensure_ascii=False, indent=2).encode("utf-8")
                return Response(data, mimetype="application/json",
                                headers={"Content-Disposition": 'attachment; filename="converted.kfp"'})
        except Exception as exc:
            import traceback
            return jsonify({"ok": False, "message": f"변환 실패: {exc}",
                            "traceback": traceback.format_exc()[-2000:]}), 400
        finally:
            try: os.unlink(tmp_path)
            except OSError: pass
