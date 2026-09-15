# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 3단계 변환(.kfp/.sdf/.slf)과 내려받기."""
from __future__ import annotations

import os
from pathlib import Path

from flask import jsonify, request, send_file

from routes.module_f.common import GROUP_DIAGRAM, _boot, _fail
from routes.module_f.jobs import (_job_running, _job_view, _run_job, route_session)
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
        # 옛 호출부(remote_only)는 그 뜻대로 옮겨 읽는다. 전체망을 PIPENET 문법
        # 으로 재직렬화하던 경로는 은퇴했다 — 설계구역 없는 SDF 는 수리계산
        # 입력이 아니라는 것이 확정 결정이다(D3). 그 함수도 2026-08-31 에
        # 지웠다(`remote30.py` 의 정리 주석 참조).
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

        # 최불리 계열은 «앞 단계» 의 산출을 재료로 쓴다 — 막지 말고 그리로
        # 안내한다. (2026-09: 단계 순서를 손질 → 수리계산 → 변환 으로 바로
        # 잡았다. 종전에는 변환이 4, 수리계산이 5 여서 앞 단계가 뒤 단계의
        # 산출을 요구하는 회로였다 — 이 메시지가 그 증거였다.)
        worst = sess.get("worst") or (
            (sess.get("design") or {}).get("got") or {}).get("worst")
        if (outputs["worst_kfp"] or outputs["worst_sdf"]) and not worst:
            return jsonify({
                "ok": False, "code": "worst_required",
                "message": "최불리 선정이 아직입니다 — 앞 단계 «손질» 에서 "
                           "「최불리 선정」을 먼저 누르세요."})
        if (outputs["worst_sdf"] or outputs["worst_kfp"]) and \
                not sess.get("design"):
            return jsonify({
                "ok": False, "code": "worst_required",
                "message": "수리계산 입력 표가 아직입니다 — 앞 단계 «수리계산» "
                           "에서 「표 확정」을 먼저 누르세요."})
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
                # ★[가지치기·부속판정 §3-5 · D5] **표의 망**에서 낸다.
                #
                #   종전에는 여기서 제한 전개를 한 번 더 돌렸다 —— 같은 K 인데
                #   `.sdf`(표에서 남)와 다른 망이 나왔다. 실측(대명동 K=30):
                #   배관 347 vs 242 · 관경 전부 기본값 25A · 부속 0건 ·
                #   등가길이 0. 이제 둘이 한 표에서 난다.
                from services.cad_import.design.emit import emit_design_kfp
                d_ = sess.get("design") or {}
                n_k = len((worst or {}).get("heads") or [])
                print(f"[변환] 최불리 K{n_k} — 표의 망으로 .kfp 를 씁니다…")
                out_path = out_dir / f"{sess['id']}_최불리K{n_k}.kfp"
                res = emit_design_kfp(d_["tables"], d_["got"], str(out_path))
                kfp_w = res["kfp"]
                sess["worst_kfp_path"] = str(out_path)
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

        # 잡 이름은 화면 진행표시에 그대로 뜬다. 이 단계는 .kfp 만 내는 것이
        # 아니라 .sdf(+.slf) 도 낸다 — 한 형식 이름을 단계 이름으로 쓰면
        # 나머지 산출이 없는 것처럼 읽힌다.
        _run_job(sess, "수리계산 입력 변환", job)
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
        """`what` 하나에 파일 하나 — **묶지 않는다**.

        ★[오너 2026-09-14] 「zip 으로 한꺼번에」 단추를 없애고 형태별로 따로
          받게 한다. SDF 만은 `.slf`(호칭경 대조 자료) 가 **같은 폴더에**
          있어야 PIPENET 이 관경을 찾는다 — 그래서 화면이 두 번 내려받고,
          여기서는 낱개로만 준다. 묶어 주면 사람이 압축을 풀어야 한다.

            what = kfp        전체망 .kfp
                   worst-kfp  최불리(설계) .kfp
                   design     설계 .sdf
                   design-slf 설계 .slf   ← .sdf 와 한 쌍
                   design-has 설계 .has
        """
        what = (request.args.get("what") or "kfp").lower()
        stem = sess.get("key") or "cad"
        kfp = sess.get("kfp_path")

        def _send(path, name, mime, missing):
            if not path or not os.path.isfile(path):
                return _fail(missing, 404)
            return send_file(path, as_attachment=True,
                             download_name=name, mimetype=mime)

        if what == "worst-kfp":
            wk = sess.get("worst_kfp_path")
            return _send(wk, (f"{stem}_" + os.path.basename(wk).split("_", 1)[-1]
                              if wk else ""), "application/json",
                         "아직 변환된 최불리 .kfp 가 없습니다.")
        if what == "kfp":
            return _send(kfp, f"{stem}_변환.kfp", "application/json",
                         "아직 변환된 .kfp 가 없습니다.")
        if what == "design":
            dsdf = sess.get("design_sdf_path")
            return _send(dsdf, (os.path.basename(dsdf) if dsdf else ""),
                         "application/xml",
                         "아직 만든 수리계산 입력이 없습니다.")
        if what == "design-slf":
            dslf = sess.get("design_slf_path")
            return _send(dslf, (os.path.basename(dslf) if dslf else ""),
                         "application/xml",
                         "아직 만든 .slf(호칭경 대조 자료)가 없습니다 — "
                         "이것이 없으면 PIPENET 에서 관경이 Unset 이 됩니다.")
        if what == "design-has":
            dhas = sess.get("design_has_path")
            return _send(dhas, (os.path.basename(dhas) if dhas else ""),
                         "application/json",
                         "아직 만든 .has 가 없습니다.")
        return _fail(f"내려받을 대상이 아닙니다: {what}")
