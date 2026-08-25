# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 3단계 변환(.kfp/.sdf/.slf)과 내려받기."""
from __future__ import annotations

import os
import zipfile
from pathlib import Path

from flask import jsonify, request, send_file

from routes.module_f.common import GROUP_DIAGRAM, _boot, _fail
from routes.module_f.jobs import _job_running, _job_view, _run_job, _sess
from routes.module_f.remote30 import _emit_pipenet, _restrict_to_worst


def register(app, *, UPLOAD_DIR):
    # ─────────────────────────────────────────── 3. 변환
    @app.get("/api/module-f/convert/fields")
    def module_f_convert_fields():
        try:
            _boot()
        except Exception as exc:  # noqa: BLE001
            return _fail(str(exc), 500)
        from services.cad_import.dto import (
            BRANCH_FIELDS, COMBO_FIELDS, FLEX_FIELDS, PENDANT_FIELDS,
            SHARED_FIELDS, UPRIGHT_FIELDS, VALVE_FIELDS, default_dto)
        groups = [
            ("메인 → 가지", BRANCH_FIELDS),
            ("상향식", UPRIGHT_FIELDS),
            ("하향식", PENDANT_FIELDS),
            ("상하향식", COMBO_FIELDS),
            ("후렉시블", FLEX_FIELDS),
            ("공통", SHARED_FIELDS),
            ("알람밸브", VALVE_FIELDS),
        ]
        return jsonify({
            "ok": True,
            "defaults": default_dto(),
            "groups": [{"title": t,
                        "diagram": GROUP_DIAGRAM.get(t),
                        "fields": [{"key": k, "label": lb, "placeholder": ph}
                                   for k, lb, ph, _d in fs]}
                       for t, fs in groups],
        })

    def _src_view(src, i):
        tag = src.get("tag") if isinstance(src, dict) else None
        xy = src.get("xy") if isinstance(src, dict) else src
        return {"tag": tag or f"Z{i + 1}", "index": i + 1,
                "xy": [round(float(v), 1) for v in (xy or [0, 0])[:2]]}

    @app.post("/api/module-f/convert/run")
    def module_f_convert_run():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        dto = body.get("dto") or {}
        selected = body.get("selected_source")
        remote_only = bool(body.get("remote_only"))
        want_sdf = body.get("emit_sdf", True)
        if remote_only and not sess.get("worst"):
            return _fail("최불리 헤드를 먼저 선정해야 그 범위로 변환할 수 있습니다.")
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.", 409)

        def job():
            from services.cad_import.convert.engine import (
                convert_to_kfp, ensure_planar)
            from services.cad_import.convert.planar import pick_convert_sources
            from services.cad_import.convert.preflight import (
                preflight_kfp_convert)
            from services.cad_import.dto import (
                default_dto, dto_to_convert_kwargs)

            payload = es.convert_payload()
            if remote_only:
                payload = _restrict_to_worst(payload, es.board, sess["worst"])
            if selected is not None:
                payload["selected_source"] = selected
            srcs = payload.get("sources") or ()
            if len(srcs) > 1:
                picked, err = pick_convert_sources(srcs, selected)
                if err:
                    return {"ok": False, "blockers": [
                        {"code": err[0], "message": err[1]}],
                        "sources": [_src_view(s, i)
                                    for i, s in enumerate(srcs)]}
                payload["sources"] = picked

            pf = preflight_kfp_convert(payload)
            if not pf["ok"]:
                print(f"[변환] 사전검사 막힘 {len(pf['blockers'])}건")
                return {"ok": False, "blockers": list(pf["blockers"]),
                        "diagnostics": list(pf.get("diagnostics") or [])}

            print("[변환] 평면 그래프를 만드는 중…")
            payload = ensure_planar(payload)
            if payload.get("kfp") is None and not payload.get("kfp_path"):
                return {"ok": False, "blockers": [{
                    "code": payload.get("_planar_code") or "planar_kfp_missing",
                    "message": payload.get("_planar_error")
                    or "평면 그래프 .kfp 가 없습니다."}]}

            merged = default_dto()
            for k, v in (dto or {}).items():
                if k in merged:
                    merged[k] = v
            print("[변환] 수직 전개 후 .kfp 를 씁니다…")
            out_dir = Path(UPLOAD_DIR) / "module_f"
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / f"{sess['id']}.kfp"
            res = convert_to_kfp(payload, str(out_path),
                                 **dto_to_convert_kwargs(merged))
            if not res["ok"]:
                return {"ok": False, "blockers": list(res["blockers"]),
                        "diagnostics": list(res.get("diagnostics") or [])}
            kfp = res["kfp"]
            sess["kfp"] = kfp
            sess["kfp_path"] = str(out_path)
            sess["sdf_path"] = None
            sess["slf_path"] = None
            stats = dict(res.get("stats") or {})
            summary = {
                "nodes": len(kfp.get("nodes_meta_runtime") or {}),
                "pipes": len(kfp.get("pipe_data") or {}),
                "bytes": out_path.stat().st_size,
                "filename": f"{sess['key'] or 'cad'}_변환.kfp",
                "remote_only": remote_only,
                "heads": len(payload.get("hcov") or []),
            }
            print(f"[변환] 완료 · 노드 {summary['nodes']} · "
                  f"배관 {summary['pipes']} · {summary['bytes']:,} bytes")

            if want_sdf:
                summary["sdf"] = _emit_pipenet(sess, kfp, out_dir)
            return {"ok": True, "stats": stats, "summary": summary,
                    "diagnostics": list(res.get("diagnostics") or [])}

        _run_job(sess, "KFP 변환", job)
        return jsonify({"ok": True})

    @app.get("/api/module-f/convert/result")
    def module_f_convert_result():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        job = sess.get("job") or {}
        return jsonify({"ok": True, "job": _job_view(sess),
                        "result": job.get("result")})

    @app.get("/api/module-f/download")
    def module_f_download():
        """`what=kfp|sdf|set` — 낱개 또는 한 벌(zip)."""
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        what = (request.args.get("what") or "kfp").lower()
        stem = sess.get("key") or "cad"
        kfp = sess.get("kfp_path")
        sdf = sess.get("sdf_path")
        slf = sess.get("slf_path")

        if what == "kfp":
            if not kfp or not os.path.isfile(kfp):
                return _fail("아직 변환된 .kfp 가 없습니다.", 404)
            return send_file(kfp, as_attachment=True,
                             download_name=f"{stem}_변환.kfp",
                             mimetype="application/json")
        if what == "sdf":
            if not sdf or not os.path.isfile(sdf):
                return _fail("아직 생성된 .sdf 가 없습니다.", 404)
            return send_file(sdf, as_attachment=True,
                             download_name=f"{stem}.sdf",
                             mimetype="application/xml")
        if what != "set":
            return _fail(f"내려받을 대상이 아닙니다: {what}")

        if not kfp or not os.path.isfile(kfp):
            return _fail("아직 변환 결과가 없습니다.", 404)
        out_dir = Path(UPLOAD_DIR) / "module_f"
        zip_path = out_dir / f"{sess['id']}_set.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(kfp, f"{stem}_변환.kfp")
            if sdf and os.path.isfile(sdf):
                z.write(sdf, f"{stem}.sdf")
            # SDF 는 라이브러리(.slf) 없이는 PIPENET 이 열지 못한다 — 같이 담는다.
            if slf and os.path.isfile(slf):
                z.write(slf, f"{stem}.slf")
        return send_file(str(zip_path), as_attachment=True,
                         download_name=f"{stem}_수리계산입력.zip",
                         mimetype="application/zip")
