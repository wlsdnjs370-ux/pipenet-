# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 3단계 변환(.kfp/.sdf/.slf)과 내려받기."""
from __future__ import annotations

import os
import zipfile
from pathlib import Path

from flask import jsonify, request, send_file

from routes.module_f.common import GROUP_DIAGRAM, _boot, _fail
from routes.module_f.jobs import (_job_running, _job_view, _run_job, _sess,
                                  route_session)
from routes.module_f.remote30 import _restrict_to_worst


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
    @route_session(post=True)
    def module_f_convert_run(sess, body):
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        dto = body.get("dto") or {}
        selected = body.get("selected_source")
        # [F-4 · D3] 산출 3종 체크 — 전체망 .kfp / 최불리 .kfp / 최불리 .sdf.
        # 옛 호출부(remote_only)는 그 뜻대로 옮겨 읽는다. 기존 emit_sdf →
        # _emit_pipenet(전체망 문법 재직렬화) 경로는 은퇴했다 — 설계구역 없는
        # SDF 는 수리계산 입력이 아니라는 것이 확정 결정이다(D3).
        outputs = body.get("outputs")
        if outputs is None:
            remote_only = bool(body.get("remote_only"))
            outputs = {"full_kfp": not remote_only, "worst_kfp": remote_only,
                       "worst_sdf": False}
        outputs = {k: bool(outputs.get(k)) for k in
                   ("full_kfp", "worst_kfp", "worst_sdf")}
        if not any(outputs.values()):
            return _fail("산출물을 하나도 고르지 않았습니다.")
        sess["convert_outputs"] = outputs      # 다음에도 같은 선택으로 뜬다

        # 최불리 계열은 선정이 있어야 한다 — 막지 말고 수리계산 패널로 안내.
        worst = sess.get("worst") or (
            (sess.get("design") or {}).get("got") or {}).get("worst")
        if (outputs["worst_kfp"] or outputs["worst_sdf"]) and not worst:
            return jsonify({
                "ok": False, "code": "worst_required",
                "message": "최불리 선정이 아직입니다 — 수리계산 패널에서 "
                           "「표 확정」을 먼저 눌러 주세요."})
        if outputs["worst_sdf"] and not sess.get("design"):
            return jsonify({
                "ok": False, "code": "worst_required",
                "message": "수리계산 입력 표가 아직입니다 — 수리계산 패널에서 "
                           "「표 확정」을 먼저 눌러 주세요."})
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

            merged = default_dto()
            for k, v in (dto or {}).items():
                if k in merged:
                    merged[k] = v
            out_dir = Path(UPLOAD_DIR) / "module_f"
            out_dir.mkdir(parents=True, exist_ok=True)

            def convert_one(restrict_worst):
                """한 판 변환 준비 — 전체망이든 최불리든 같은 경로를 탄다."""
                payload = es.convert_payload()
                if restrict_worst is not None:
                    payload = _restrict_to_worst(payload, es.board,
                                                 restrict_worst)
                if selected is not None:
                    payload["selected_source"] = selected
                srcs = payload.get("sources") or ()
                if len(srcs) > 1:
                    picked, err = pick_convert_sources(srcs, selected)
                    if err:
                        return None, {"ok": False, "blockers": [
                            {"code": err[0], "message": err[1]}],
                            "sources": [_src_view(s, i)
                                        for i, s in enumerate(srcs)]}
                    payload["sources"] = picked
                pf = preflight_kfp_convert(payload)
                if not pf["ok"]:
                    print(f"[변환] 사전검사 막힘 {len(pf['blockers'])}건")
                    return None, {"ok": False,
                                  "blockers": list(pf["blockers"]),
                                  "diagnostics":
                                  list(pf.get("diagnostics") or [])}
                print("[변환] 평면 그래프를 만드는 중…")
                payload = ensure_planar(payload)
                if payload.get("kfp") is None and not payload.get("kfp_path"):
                    return None, {"ok": False, "blockers": [{
                        "code": payload.get("_planar_code")
                        or "planar_kfp_missing",
                        "message": payload.get("_planar_error")
                        or "평면 그래프 .kfp 가 없습니다."}]}
                return payload, None

            summary = {"outputs": dict(outputs)}
            stats = {}

            if outputs["full_kfp"]:
                payload, err = convert_one(None)
                if err:
                    return err
                print("[변환] 전체망 — 수직 전개 후 .kfp 를 씁니다…")
                out_path = out_dir / f"{sess['id']}.kfp"
                res = convert_to_kfp(payload, str(out_path),
                                     **dto_to_convert_kwargs(merged))
                if not res["ok"]:
                    return {"ok": False, "blockers": list(res["blockers"]),
                            "diagnostics":
                            list(res.get("diagnostics") or [])}
                kfp = res["kfp"]
                sess["kfp"] = kfp
                sess["kfp_path"] = str(out_path)
                stats = dict(res.get("stats") or {})
                summary["full"] = {
                    "nodes": len(kfp.get("nodes_meta_runtime") or {}),
                    "pipes": len(kfp.get("pipe_data") or {}),
                    "bytes": out_path.stat().st_size,
                    "filename": f"{sess['key'] or 'cad'}_변환.kfp",
                }
                print(f"[변환] 전체망 완료 · 노드 {summary['full']['nodes']}"
                      f" · 배관 {summary['full']['pipes']} · "
                      f"{summary['full']['bytes']:,} bytes")

            if outputs["worst_kfp"]:
                payload, err = convert_one(worst)
                if err:
                    return err
                n_k = len((worst or {}).get("heads") or [])
                print(f"[변환] 최불리 K{n_k} — 수직 전개 후 .kfp 를 씁니다…")
                # 파일명으로 전체망본과 구분한다 — 같은 이름이면 어느 쪽인지
                # 열어 보기 전엔 모른다.
                out_path = out_dir / f"{sess['id']}_최불리K{n_k}.kfp"
                res = convert_to_kfp(payload, str(out_path),
                                     **dto_to_convert_kwargs(merged))
                if not res["ok"]:
                    return {"ok": False, "blockers": list(res["blockers"]),
                            "diagnostics":
                            list(res.get("diagnostics") or [])}
                kfp_w = res["kfp"]
                sess["worst_kfp_path"] = str(out_path)
                if not stats:
                    stats = dict(res.get("stats") or {})
                summary["worst"] = {
                    "k": n_k,
                    "nodes": len(kfp_w.get("nodes_meta_runtime") or {}),
                    "pipes": len(kfp_w.get("pipe_data") or {}),
                    "bytes": out_path.stat().st_size,
                    "filename": f"{sess['key'] or 'cad'}_최불리K{n_k}.kfp",
                }
                print(f"[변환] 최불리 완료 · 노드 {summary['worst']['nodes']}"
                      f" · 배관 {summary['worst']['pipes']} · "
                      f"{summary['worst']['bytes']:,} bytes")

            if outputs["worst_sdf"]:
                from routes.module_f.api_design import emit_design_files
                out, err = emit_design_files(sess, UPLOAD_DIR)
                if err:
                    return {"ok": False, "blockers": [
                        {"code": "design_emit_failed", "message": err}]}
                slf = out.with_suffix(".slf")
                summary["design"] = {
                    "sdf": out.name, "bytes": out.stat().st_size,
                    "slf": slf.name,
                }
                print(f"[변환] 최불리 SDF · {out.name} · "
                      f"{out.stat().st_size:,} bytes (+.slf)")

            return {"ok": True, "stats": stats, "summary": summary,
                    "diagnostics": []}

        _run_job(sess, "KFP 변환", job)
        return jsonify({"ok": True})

    @app.get("/api/module-f/convert/result")
    @route_session()
    def module_f_convert_result(sess, body):
        job = sess.get("job") or {}
        return jsonify({"ok": True, "job": _job_view(sess),
                        "result": job.get("result")})

    @app.get("/api/module-f/download")
    @route_session()
    def module_f_download(sess, body):
        """`what=kfp|sdf|set` — 낱개 또는 한 벌(zip)."""
        what = (request.args.get("what") or "kfp").lower()
        stem = sess.get("key") or "cad"
        kfp = sess.get("kfp_path")
        sdf = sess.get("sdf_path")
        slf = sess.get("slf_path")

        if what == "worst-kfp":
            wk = sess.get("worst_kfp_path")
            if not wk or not os.path.isfile(wk):
                return _fail("아직 변환된 최불리 .kfp 가 없습니다.", 404)
            return send_file(wk, as_attachment=True,
                             download_name=f"{stem}_"
                             + os.path.basename(wk).split("_", 1)[-1],
                             mimetype="application/json")
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
        if what == "design":
            # [F-2] 수리계산 입력 한 벌 — SDF 는 옆의 SLF 와 한 쌍이다(파일명
            # 참조라 따로 열면 관경이 Unset). 그래서 낱개가 아니라 zip 으로만 준다.
            dsdf = sess.get("design_sdf_path")
            dslf = sess.get("design_slf_path")
            if not dsdf or not os.path.isfile(dsdf):
                return _fail("아직 만든 수리계산 입력이 없습니다.", 404)
            out_dir = Path(UPLOAD_DIR) / "module_f"
            zip_path = out_dir / f"{sess['id']}_design.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(dsdf, os.path.basename(dsdf))
                if dslf and os.path.isfile(dslf):
                    z.write(dslf, os.path.basename(dslf))
            return send_file(str(zip_path), as_attachment=True,
                             download_name=f"{stem}_수리계산입력_설계.zip",
                             mimetype="application/zip")
        if what != "set":
            return _fail(f"내려받을 대상이 아닙니다: {what}")

        wk = sess.get("worst_kfp_path")
        dsdf = sess.get("design_sdf_path")
        dslf = sess.get("design_slf_path")
        have = [q for q in (kfp, wk, dsdf) if q and os.path.isfile(q)]
        if not have:
            return _fail("아직 변환 결과가 없습니다.", 404)
        out_dir = Path(UPLOAD_DIR) / "module_f"
        zip_path = out_dir / f"{sess['id']}_set.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            if kfp and os.path.isfile(kfp):
                z.write(kfp, f"{stem}_변환.kfp")
            if wk and os.path.isfile(wk):
                z.write(wk, f"{stem}_" + os.path.basename(wk).split("_", 1)[-1])
            if dsdf and os.path.isfile(dsdf):
                z.write(dsdf, os.path.basename(dsdf))
                # SDF 는 .slf 없이는 PIPENET 이 못 연다 — 같이 담는다.
                if dslf and os.path.isfile(dslf):
                    z.write(dslf, os.path.basename(dslf))
            # 은퇴한 전체망 문법 재직렬화 SDF — 남아 있으면 그대로 담아 준다.
            if sdf and os.path.isfile(sdf):
                z.write(sdf, f"{stem}.sdf")
                if slf and os.path.isfile(slf):
                    z.write(slf, f"{stem}.slf")
        return send_file(str(zip_path), as_attachment=True,
                         download_name=f"{stem}_수리계산입력.zip",
                         mimetype="application/zip")
