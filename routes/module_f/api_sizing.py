"""Module F optional hydraulic proposals: separate from live merge outputs."""
from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
import json
import uuid

from flask import Flask, jsonify, send_file
from routes.module_f.common import _fail
from routes.module_f.jobs import route_session, _run_job, _job_running
from routes.module_f.hydraulic_sizing import fingerprint, input_summary, prepare
from routes.module_f.cancellation import checkpoint


def _tables(sess):
    got = sess.get("merged") or {}
    if got.get("combined") is None:
        raise ValueError("먼저 통합 배관망을 결합하세요.")
    return got["combined"]


def register(app: Flask, *, UPLOAD_DIR: str | Path) -> None:
    """Register read-only preview and non-destructive background proposals."""
    @app.get("/api/module-f/merge/sizing/state")
    @route_session()
    def module_f_sizing_state(sess, body):
        try:
            tables = _tables(sess)
            basis = input_summary(tables)
            report = sess.get("sizing_report")
            stale = bool(report and report["fingerprint"] != basis["fingerprint"])
            return jsonify(ok=True,basis=basis,report=report,stale=stale,
                           review_defaults=json.loads((Path(__file__).resolve().parents[2]/"configs"/"hydraulic_review.json").read_text(encoding="utf-8")),
                           files=list(sess.get("sizing_files") or {}) if not stale else [])
        except ValueError as exc:
            return _fail(str(exc),400)

    @app.post("/api/module-f/merge/sizing/run")
    @route_session(post=True)
    def module_f_sizing_run(sess, body):
        if _job_running(sess):
            return _fail("진행 중인 작업이 끝난 뒤 실행하세요.",409)
        try:
            tables = _tables(sess)
            revision = fingerprint(tables)
            if body.get("fingerprint") != revision:
                return _fail("결합망이 바뀌었습니다. 역산 입력을 다시 불러오세요.",409)
            options = body.get("options") or {}
            if not isinstance(options,dict):
                raise ValueError("역산 설정 형식 오류")
            net, settings, warnings = prepare(tables,options)
        except (ValueError,TypeError) as exc:
            return _fail(str(exc),400)
        def job():
            from src.pipenet_converter.hydraulics.sizing import propose
            result = propose(net,settings,checkpoint=checkpoint,progress=print)
            checkpoint()
            if revision != fingerprint(_tables(sess)):
                raise ValueError("계산 중 결합망이 변경되었습니다. 결과를 적용하지 않았습니다.")
            result.update(fingerprint=revision,model=asdict(net),warnings=warnings,
                          options=options,proposal_id=uuid.uuid4().hex)
            originals = {str(p['label']):p.get('dia') or None for p in tables.pipes}
            for row in result.get('pipes',[]):
                row['before_dn'] = originals[row['label']]
            sess["sizing_report"] = result
            sess["sizing_files"] = {}
            print(f"[역산] {result['status']} · 필요 유량 {result['pump']['flow_lpm']:.2f} L/min"
                  f" · 양정 {result['pump']['differential_head_m']:.2f}m · 원본 변경 없음")
            return dict(status=result['status'],proposal_id=result['proposal_id'])
        _run_job(sess,"관경·펌프 역산",job)
        return jsonify(ok=True,sid=sess['id'])

    @app.post("/api/module-f/merge/sizing/emit")
    @route_session(post=True)
    def module_f_sizing_emit(sess, body):
        if _job_running(sess):
            return _fail("진행 중인 작업이 끝난 뒤 출력하세요.",409)
        report = sess.get('sizing_report')
        try:
            if not report or report['fingerprint'] != fingerprint(_tables(sess)):
                raise ValueError("유효한 역산 검토안이 없습니다. 다시 계산하세요.")
            if body.get('proposal_id') != report['proposal_id']:
                raise ValueError("검토안이 변경되었습니다. 결과를 다시 확인하세요.")
            if not report['feasible']:
                raise ValueError("조건 미달/미수렴 상태입니다. 진단 결과를 먼저 확인하세요.")
        except ValueError as exc:
            return _fail(str(exc),409)
        def job():
            from routes.module_f.sizing_export import emit_proposal
            out = Path(UPLOAD_DIR)/'module_f_sizing'/sess['id']/report['proposal_id']
            files = emit_proposal(sess['merged'],report,out)
            checkpoint()
            sess['sizing_files'] = files
            return dict(files=list(files))
        _run_job(sess,"역산 검토안 SDF 생성",job)
        return jsonify(ok=True,sid=sess['id'])

    @app.get("/api/module-f/merge/sizing/download")
    @route_session()
    def module_f_sizing_download(sess, body):
        report = sess.get('sizing_report')
        try:
            if not report or report['fingerprint'] != fingerprint(_tables(sess)):
                raise ValueError("결합망 변경으로 지난 결과의 다운로드를 차단했습니다.")
        except ValueError as exc:
            return _fail(str(exc),409)
        what = body.get('what','report')
        if what == 'report':
            exported = (sess.get('sizing_files') or {}).get('report')
            if exported and Path(exported).is_file():
                # Same proposal plus the export-only node/pipe provenance map.
                return send_file(exported,as_attachment=True)
            response = jsonify(report)
            response.headers['Content-Disposition'] = 'attachment; filename="hydraulic_review.json"'
            return response
        path = (sess.get('sizing_files') or {}).get(what)
        if what not in {'sdf','slf','sdf_iso','slf_iso'} or not path or not Path(path).is_file():
            return _fail("검토안 SDF를 먼저 생성하세요.",404)
        return send_file(path,as_attachment=True)
